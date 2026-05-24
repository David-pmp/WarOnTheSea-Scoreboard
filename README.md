This project organizes the array of data involved in the World War 2 naval combat simulator "War on the Sea", available from Steam, and produced by Killer Fish.

The WarOnTheSeaScoreboard.ps1 file is a PowerShell script that pulls in the most recent campaign save from the user's campaign save folder, 
and presents the information in that into an Excel spreadsheet that shows the status of every ship and every land location in that campaign 
with the aid of the various campaign setup and data files.  Also shows status by type and ship class.

The Excel document is unique to that save file and is written out to either the user's local Documents folder, or their OneDrive folder.

This script has been developed against Kiko's new Pacific mod for WotS.  It should at least some summary ability for any save file running with its current mod in place.

HOW TO USE:
Download the WarOnTheSeaScoreboard.ps1 script to your downloads folder.

Open up the downloads folder in Windows Explorer

Find the script in the downloads folder.

Right click on the file and select 'run as powershell'

If the above doesn't work, you can also try:
Right click on the file and select 'copy as path'

Open Powershell.exe (Start menu and search for Powershell)

Type in "." and a space and then paste in the path.  Press Enter to run.

