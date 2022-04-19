What you will find in this Powershell directory is mainly an inventory macro-enabled spreadsheet that I put together. Every quarter, our department has to conduct an inventory audit to ensure everything is still under our custody as well as verify serial number accuracy. Drives are classified so they are also inventoried.

With that in mind, I created buttons, wrote VBA in Excel, and wrote Powershell scripts, all in conjunction for the following functions:

1. Check Online status. The button executes a Powershell script that takes all computer names in the inventory and tests for connectivity. Once every machine has been tested, the results are imported into the spreadsheet.

2. Conduct Inventory. The button executes a Powershell script that takes all computer names in the inventory and uses WMI to get the serial number of the computer, as well as the drive serial number. The results are then imported into Excel and then the results are color coded by whether or not it matched what is currently on the inventory.

3. Clear/Hide/Unhide Automated Fields. Self-explanatory. Three buttons, clear (clear all imported results), hide (hide automated fields), unhide (unhide automated fields).

Lastly, if you'd like to see the VBA code, I've included it in the directory as rawVBA.txt so that you don't have to download the spreadsheet and dig into it.
