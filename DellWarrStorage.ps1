<#
.SYNOPSIS
    The script pulls from a CSV file containing Dell service tags, retrieves warranty information and storage details for each device using the Dell API, and exports the results to a new CSV file.
    The new CSV file will include the original columns along with new columns for ship date, warranty expiration, and storage information.
    .DESCRIPTION
    This script reads service tags, retrieves warranty information (ship date and end date) and storage details (size and disk type) for each device using the Dell API.
    .PARAMETER ServiceTag
    The service tag of the Dell device for which warranty information is to be retrieved.
    .PARAMETER ApiKey
    Your Dell API key for authentication. Make sure it is valid and has the necessary permissions.
    .PARAMETER ApiSecret
    Your Dell API secret for authentication. Make sure it is valid and has the necessary permissions.
    .PARAMETER AccessToken
    The access token generated from the API key and secret, used for making authenticated requests to the Dell API.
    .par
    .EXAMPLE
    Get-DellWarrantyInfo -ServiceTag "ABC123" -ApiKey "l780b2d312099e4576ba345b4718b5ca46" -ApiSecret "f73e956e648a42cd962e5ca4c58ec660"
    Retrieves warranty information and storage information for the device with service tag "ABC123".
#>

# ---------------------------------------------------------------------------------------------------------------
# Define the path to the CSV file imports and exports containing service tags
$importFilePath = "AYC_InvShip.csv"
$exportFilePath = "AYC_InvShipRUN.csv"
# Your API key and secret
# ***MAKE SURE THESE ARE YOUR ACTIVE API CREDENTIALS***
# These credentials are used to authenticate with the Dell API
# MAKE SURE TO KEEP THEM SECURE AND DO NOT SHARE THEM PUBLICLY
$apiKey = "INSERT_YOUR_API_KEY_HERE"
$apiSecret = "INSERT_YOUR_API_SECRET_HERE"
# ---------------------------------------------------------------------------------------------------------------

# Get-DellWarrrantyInfo function retrieves warranty information for a Dell device using its service tag.
# It uses the Dell API to get the ship date and warranty end date.
function Get-DellWarrantyInfo {
    param (
        [string]$ServiceTag,
        [string]$ApiKey,
        [string]$ApiSecret
    )

    try {
        
        # Use the OAuth 2.0 token from main script
        # This token is generated using the API key and secret provided
        $accessToken = $tokenResponse.access_token

        # The warrenty API uses *asset entitlements* to retrieve warranty information
        # This endpoint provides ship date and warranty end date for the device with the given service tag
        $warrantyResponse = Invoke-RestMethod -Uri "https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements?servicetags=$ServiceTag" -Headers @{
            Authorization = "Bearer $accessToken"
        }
        # Initialize variables for ship date and warranty end date
        $shipDate = $null
        $warrantyEndDate = $null

        # Check if the response contains valid data
        # If the response is valid, extract the ship date and warranty end date
        if ($warrantyResponse -and $warrantyResponse[0]) {
            $shipDate = $warrantyResponse[0].shipDate
            $entitlements = $warrantyResponse[0].entitlements
            if ($entitlements) {
                $endDates = $entitlements | ForEach-Object { $_.endDate }
                if ($endDates) {
                    $warrantyEndDate = ($endDates | Sort-Object | Select-Object -Last 1)
                }
            }
        }

        # Return a hashtable containing ship date and warranty expiration date
        # This will be used later to populate the CSV file
        # If no ship date or warranty end date is found, they will be returned as null
        return @{
            ShipDate = $shipDate
            WarrantyExpiration = $warrantyEndDate
        }
    } catch {
        Write-Output "An error occurred for Service Tag: $ServiceTag - $_"
        return $null
    }
}  #end of Get-DellWarrantyInfo function


# Get-DellStorageInfo function retrieves storage information for a Dell device using its service tag.
# It uses the Dell API to get the storage size and drive type.
function Get-DellStorageInfo {
    param (
        [string]$ServiceTag,
        [string]$AccessToken
    )
    try {
        # The warranty API uses *asset components* to retrieve storage information
        # This endpoint provides details about the storage components for the device with the given service tag
        $response = Invoke-RestMethod -Uri "https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-components?servicetag=$ServiceTag" -Headers @{
            Authorization = "Bearer $AccessToken"
        }
        # Check if the response contains components
        # If the response is valid, look for storage components (SSD and HDD)
        # The components may contain itemDescription and partDescription which can be used to identify storage types
        # The function will return the size and type of the first matching storage component found
        if ($response.components) {
            $storageComponents = $response.components | Where-Object {
                ($_.itemDescription -match 'SSD|HDD|Solid State Drive|SSDR') -or
                ($_.partDescription -match 'SSD|HDD|Solid State Drive|SSDR')
            }
            # If no storage components found, return an empty string
            # If storage components are found, extract size and type
            # The size is extracted from the itemDescription or partDescription using regex 
            foreach ($comp in $storageComponents) {
                $desc = "$($comp.itemDescription) $($comp.partDescription)"
                $number = ""
                $unit = ""
                if ($desc -match '(\d+)\s*GB') {
                    $number = $matches[1]
                    $unit = "GB"
                } elseif ($desc -match '(\d+)\s*TB') {
                    $number = $matches[1]
                    $unit = "TB"
                } elseif ($desc -match '(\d+)\s*G\b') {
                    $number = $matches[1]
                    $unit = "GB"
                } elseif ($desc -match '(\d+)\s*T\b') {
                    $number = $matches[1]
                    $unit = "TB"
                }
                # Look for type
                if ($desc -match 'SSD|Solid State Drive|SSDR') {
                    $type = "SSD"
                } elseif ($desc -match 'HDD|Hard Drive|HD') {
                    $type = "HDD"
                } else {
                    $type = ""
                }
                if ($number -and $unit -and $type) {
                    return "$number $unit $type"
                }
            }
            # If nothing matched, return the first storage component's description as fallback
            # This has not been superuseful. 
            if ($storageComponents.Count -gt 0) {
                $first = $storageComponents | Select-Object -First 1
                return "$($first.itemDescription) $($first.partDescription)".Trim()
            }
        }
        return ""
    # If no storage components found, return an empty string
    } catch {
        Write-Output "Failed to get storage info for $ServiceTag - $_"
        return ""
    }
} #end of Get-DellStorageInfo function


# Start of the main script

# Read serial numbers from $importFilePath (as defined at the top of the script)
# This file should contain a column named 'Serial Number' with the service tags of Dell devices
$serialNumbers = Import-Csv -Path $importFilePath



# Get OAuth 2.0 access token once for all requests
$tokenResponse = Invoke-RestMethod -Uri "https://apigtwb2c.us.dell.com/auth/oauth/v2/token" -Method Post -Body @{
    client_id = $apiKey
    client_secret = $apiSecret
    grant_type = "client_credentials"
}
$accessToken = $tokenResponse.access_token

# Capture the original column order
$originalColumns = @()
$seen = @{}
foreach ($col in $serialNumbers[0].PSObject.Properties.Name) {
    $key = ($col -replace '\s+', '').ToLower()
    if (-not $seen.ContainsKey($key)) {
        $originalColumns += $col
        $seen[$key] = $true
    }
}
# Prepare results array
$results = @()

# Loop through each row in the CSV file
# For each service tag, retrieve warranty information (ship date and warranty expiration) and storage information
foreach ($row in $serialNumbers) {
    $serviceTag = $row.'Serial Number'
    if ($serviceTag) {
        $info = Get-DellWarrantyInfo -ServiceTag $serviceTag -ApiKey $apiKey -ApiSecret $apiSecret
        if ($info) {
            if ($info.ShipDate) {
                $shipDate = (Get-Date $info.ShipDate).ToString("MM/dd/yyyy")
            } else {
                $shipDate = ""
            }
            if ($info.WarrantyExpiration) {
                $warrantyExpiration = (Get-Date $info.WarrantyExpiration).ToString("MM/dd/yyyy")
            } else {
                $warrantyExpiration = ""
            }
        } else {
            $shipDate = ""
            $warrantyExpiration = ""
        }
        $storage = Get-DellStorageInfo -ServiceTag $serviceTag -AccessToken $accessToken

        # Build output row in the same order as imported, appending Storage at the end
        # Create a new object with the original columns and the new columns for ship date, warranty expiration, and storage
        # This ensures that the new columns are added without disrupting the original column order
        $props = @{}
        foreach ($col in $originalColumns) {
            $props[$col] = $row.$col
        }
        $props['Ship Date'] = $shipDate
        $props['Warranty Expiration'] = $warrantyExpiration
        $props['Storage'] = $storage

        $newRow = [PSCustomObject]$props
        $results += $newRow
    }
}

# Define new columns to append
$newColumns = @('Ship Date', 'Warranty Expiration', 'Storage')

# Remove any new columns that already exist in the original columns (case-insensitive, ignoring whitespace)
$originalColumnKeys = @{}
foreach ($col in $originalColumns) {
    $key = ($col -replace '\s+', '').ToLower()
    $originalColumnKeys[$key] = $col
}
$finalNewColumns = @()
foreach ($col in $newColumns) {
    $key = ($col -replace '\s+', '').ToLower()
    if (-not $originalColumnKeys.ContainsKey($key)) {
        $finalNewColumns += $col
    }
}

# Build column order: original columns + only truly new columns
$allColumns = $originalColumns + $finalNewColumns

# Export the results to a new CSV file with columns in the correct order
# The export file path is defined at the top of the script
$results | Select-Object $allColumns | Export-Csv -Path $exportFilePath -NoTypeInformation

Write-Output "The new CSV file with serial numbers, ship dates, warranty expiration, and storage info has been created successfully."
#End of script
