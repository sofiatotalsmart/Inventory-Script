# Inventory-Script
This repo contains the PowerShell script to automate inventory CSV files with Dell's ship date, warranty end date, and storage information by using the serial numbers.

For easiest use: Create a folder and add the powershell file into the folder along with the CSV file. 
Open Visual Studio Code and open this folder you have created. 
From here you can insert the inventory CSV file name, your desired CSV file export name, as well as the API key and API secret from Dell into the powershell script.
This will be in lines 23, 24, 39, and 30. 
After entering these 4 peices of infomration you can click run and recive your exported CSV file. 
When the script is complete the message belw will appear:
"The new CSV file with serial numbers, ship dates, warranty expiration, and storage info has been created successfully."
