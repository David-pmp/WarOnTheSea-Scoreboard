#  This script generates a list of ships and locations in the most recent War On the Sea campaign file, 
#  and then gets further ship and air data from default and override folders, override taking priority.

#  Returns an Excel Spreadsheet with all the data, a csv file to show the full list of ships sunk in the game, a csv file to show ships lost by class,
#  and a csv to show the ships lost by category.

# As of May 21, 2026 this has only been tested against Kiko's new Pacific mod 1942HD running on 1.08h version of War On The Sea

# This script is written to run in PowerShell 5 without any additional modules, but does REQUIRE Microsoft Excel to get the final spreadsheet.


#Campaign save data pulled from AppData\LocalLow\Killerfish Games\War on the Sea\save\campaign\<SAVENAME.txt>

# Working out whether to save the final Excel document in user's local folder or a OneDrive folder.  Local takes precedence
$OutputPath = ($env:Userprofile + "\Documents\")
if (!(test-path $OutputPath)){
    $OutputPath = ($env:OneDrive + "\Documents\")
    if (!(test-path $OutputPath)){
        Write-Host "Unable to find path to documents folder for final output. Exiting" -ForegroundColor Red
        Exit-PSSession
    }
}

# location for campaign files to test with.
$TestingPath 

$SteamPath = (Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Valve\Steam).InstallPath
$WarOnTheSeaPath = "$SteamPath\steamapps\common\War on the Sea\WarOnTheSea_Data\StreamingAssets"
$campaignPath =  ($env:LOCALAPPDATA + "\..\LocalLow\Killerfish Games\War on the Sea\save\campaign\")
$seaUnitsOverridePath = "$WarOnTheSeaPath\override\language\english\unit\sea"
$seaUnitsDefaultPath =  "$WarOnTheSeaPath\default\language\english\unit\sea"
$seaUnitsDataDefaultPath ="$WarOnTheSeaPath\default\unit\sea\"
$seaUnitsDataOverridePath  =  "$WarOnTheSeaPath\override\unit\sea\"

$airUnitsOverridePath = "$WarOnTheSeaPath\override\language\english\unit\air"
$airUnitsDefaultPath =  "$WarOnTheSeaPath\default\language\english\unit\air"

# Get the list of save files in the campaign folder
$campaignFiles = get-childitem -Path $campaignPath 
$LastSavedFile = ($campaignFiles | where-object VersionInfo -ne $null | Sort-Object LastWriteTime -Descending)[0]

# Get the content of the last saved game file.
$CampaignSaveData = get-content  ($CampaignPath + $LastSavedFile.Name) | ConvertFrom-Json

#Get the campaignID and faction sides.
# Nations0 is always the player's side.
$CampaignID = $CampaignSaveData.campaignID
$Nations0 = $CampaignSaveData.nations0
$Nations1 = $CampaignSaveData.nations1


# Mobile formation data for later use.
$PlayerMobileData = $CampaignSaveData.playerMobileObjectSaveData |convertfrom-json
$EnemyMobileSaveData = $CampaignSaveData.enemyMobileObjectSaveData |convertfrom-json

$TimeArray = $CampaignSaveData.currentDatetime
$CampaignDate =  get-Date -year $TimeArray[5] -Month $TimeArray[4] -Day $TimeArray[3] -Hour $TimeArray[2] -Minute $timearray[1] -second $timeArray[0]
$CampaignDateFileFormat = $TimeArray[5]

$TimeArray = $CampaignSaveData.startDate
$CampaignStartDate =  get-Date -year $TimeArray[5] -Month $TimeArray[4] -Day $TimeArray[3] -Hour $TimeArray[2] -Minute $timearray[1] -second $timeArray[0]
$CampaignDay = ($CampaignDate - $CampaignStartDate).Days

$campaignLandLocationsFile =  ("$WarOnTheSeaPath\override\campaign\$campaignID" + "\mapLandLocations.txt")

$campaignSeaUnitsFile = ("$WarOnTheSeaPath\override\campaign\$campaignID" + "\seaUnits.txt")
# if we can't get the the campaign's id in the override folder then this is a default game campaign.
if (!(test-path $campaignSeaUnitsFile)){
    $campaignSeaUnitsFile = ("$WarOnTheSeaPath\default\campaign\$campaignID" + "\seaUnits.txt")
    $campaignLandLocationsFile =  ("$WarOnTheSeaPath\default\campaign\$campaignID" + "\mapLandLocations.txt")
}


$CampaignLandLocations = $CampaignSaveData.mapLocationSaveData | convertfrom-json
#Creatng a list of the player's home ports for future use.
$HomePorts =  $CampaignLandLocations| where-object CreateShips |Where-Object currentFaction -eq 0 |Select-Object locationID, locationName | sort-object


Write-Host "Processing Air Unit files" -ForegroundColor Green

$AirOverrideFiles = get-childitem -Path $airUnitsOverridePath |  where-object Name -like "*.txt" |select-object Name
$AirDefaultFiles = get-childitem -Path $airUnitsDefaultPath | where-object Name -like "*.txt" |select-object Name
$AllAir = new-Object System.Collections.Hashtable

foreach ($AirFile in $AirDefaultFiles){
 
    #Write-host $AirFile.Name
    $AirData = get-content ("$airUnitsDefaultPath\" + $AirFile.Name) | convertFrom-json
    $AirObj = new-object -typename PSCustomObject
     $AirID = $AirFile.Name.substring(0,$AirFile.Name.indexOf(".txt"))
    add-member -InputObject $AirObj -MemberType NoteProperty -Name "AirID" -Value $AirID
     
    add-member -InputObject $AirObj -MemberType NoteProperty -Name "Name" -Value $AirData.unitName
   # Write-host $AirObj
    $null = $AllAir.add($AirID, $AirData.unitName)
    Clear-Variable AirData
}

foreach ($AirFile in $AirOverrideFiles){
    
    # Write-host $AirFile.Name
    $AirData = get-content ("$airUnitsOverridePath\" + $AirFile.Name) | convertFrom-json
    $AirObj = new-object -typename PSCustomObject
    $AirID = $AirFile.Name.substring(0, $AirFile.Name.indexOf(".txt"))
    add-member -InputObject $AirObj -MemberType NoteProperty -Name "AirID" -Value $AirID
     
    add-member -InputObject $AirObj -MemberType NoteProperty -Name "Name" -Value $AirData.unitName
    # Write-host $AirObj
    if ($AllAir[$AirID]){
        $null  = $AllAir.Item($AirID) =  $AirData.unitName
    }
    else{
        $null = $AllAir.add($AirID, $AirData.unitName)
    }
    Clear-Variable AirData
}

# Get the potential list of ships from the campaignSeaUnitsFile
$CampaignShipClasses = get-content  $campaignSeaUnitsFile  | ConvertFrom-Json

# get the list of class files from the Sea units folder in the Override path
Write-Host "Processing sunk ships list from the campaign save file. " -ForegroundColor Green
$ShipClassFiles = get-childitem -Path $seaUnitsOverridePath | where-object VersionInfo -ne $null
$AllShips = new-object System.Collections.ArrayList;

$sunkenShips =  new-object System.Collections.ArrayList;
for ($SunkIteration = 0; $sunkIteration -lt $CampaignSaveData.sunkShipClasses.count; $Sunkiteration++){
     $sunkObj = New-object -typename PSCustomObject
     add-member -InputObject $sunkObj -MemberType NoteProperty -Name "ShipClassID" -Value $CampaignSaveData.sunkShipClasses[$SunkIteration]
     
     add-member -InputObject $sunkObj -MemberType NoteProperty -Name "ShipClassInstance" -Value $CampaignSaveData.sunkShipInstances[$SunkIteration]

     $sunkDate = get-date -Day $CampaignSaveData.daySunk[$SunkIteration] -Month $CampaignSaveData.monthSunk[$SunkIteration] -Year $CampaignSaveData.yearSunk[$SunkIteration]
     add-member -InputObject $sunkObj -MemberType NoteProperty -Name "SunkDate" -Value $sunkDate.Date
    $sunkenShips.add($sunkObj) | Out-Null

}

# Get the player surface task forces into an ArrayList
$PlayerTaskForces =  new-object System.Collections.ArrayList;
foreach($TaskForce in $PlayerMobileData){ 
     

     if (($TaskForce.createdByDisplayName -eq "")){   # This task force wasn't created by another object so it's a sea unit.
        #$TaskForce.mobileName
        for ($TFIteration = 0; $TFIteration -lt $TaskForce.unitPrefabs.count; $TFIteration++){
            $NewTF = New-object -typename PSCustomObject
            add-member -InputObject $NewTF -MemberType NoteProperty -Name "ShipClassID" -Value $TaskForce.unitPrefabs[$TFIteration]
     
            add-member -InputObject $NewTF -MemberType NoteProperty -Name "ShipClassInstance" -Value $TaskForce.unitInstances[$TFIteration]

            add-member -InputObject $NewTF -MemberType NoteProperty -Name "TFName" -Value $TaskForce.mobileName
            $PlayerTaskForces.add($NewTF) | Out-Null
        }
    }

}

#Main loop through the master list of ships.
$ShipClassHT = @{};

Write-Host "Processing campaign ship list" -foregroundcolor Green
$ShipClasses  = new-object System.Collections.arraylist

foreach ($ShipClassID in $CampaignShipClasses.unitID){
    $sunkInClass = ($sunkenShips |where-object ShipClassID -EQ $ShipClassID).count
  


    $ShipClassFile = "$ShipClassID.txt"
   # $ShipClass
    if( Test-path ( "$seaUnitsOverridePath\$ShipClassFile")){
        try{
            $ShipClassFileContent = Get-Content -Path "$seaUnitsOverridePath\$ShipClassFile"
            $ShipClassInfo = $ShipClassFileContent.substring(0, $ShipClassFileContent.LastIndexOf('}')+1) | ConvertFrom-Json
        }
        catch{
            Write-host "Error on JSON conversion for Override file $ShipClassFile"

        }
    }
    else{
        if (Test-path  "$seaUnitsDefaultPath\$ShipClassFile" ){
            try{
                 $ShipClassFileContent = Get-Content -Path "$seaUnitsDefaultPath\$ShipClassFile"
                  $ShipClassInfo = $ShipClassFileContent.substring(0, $ShipClassFileContent.LastIndexOf('}')+1) | ConvertFrom-Json
            }
            catch{
                 Write-host "Error on JSON conversion for Default file $ShipClassFile"
            }   
         
        }
        else{
            Write-host "$ShipClassID unit info not found"
            
        }


    }
   
    if ($ShipClassInfo){
        $ShipClassName = $ShipClassInfo.unitName
        $ShipNamesInClass = $ShipClassInfo.namesInClass
        $ShipInstancesInClass = $ShipClassInfo.instancesInClass
    }
    else {
        $ShipClassName = "Not Found"
        $ShipNamesInClass = @();
        $ShipInstancesInClass =  @();

    }
    $ShipClassHT.add($ShipClassID,$ShipClassName);

    if (Test-path ("$seaUnitsDataOverridePath\$ShipClassID\" + $ShipClassID +"_data.txt") ){
     
        $ShipDataRaw = Get-Content -Path ("$seaUnitsDataOverridePath\$ShipClassID\" + $ShipClassID +"_data.txt") 
    }
    else{
        if (Test-path ("$seaUnitsDataDefaultPath\$ShipClassID\" + $ShipClassID +"_data.txt") ){
            $ShipDataRaw = Get-Content -Path ("$seaUnitsDataDefaultPath\$ShipClassID\" + $ShipClassID +"_data.txt") 
        }
        else {
            Write-host "$ShipclassID full data not available" 
           
        }
    }
    $ShipData = $ShipDataRaw[0] |ConvertFrom-Json
    $ShipType = $ShipData.unitSubtypeString
    if ($ShipData.cargo){
        $ShipPassenger = $ShipData.cargo[0]
        $ShipSupplies = $ShipData.cargo[1]
        $ShipEngineering = $ShipData.cargo[2]
        $ShipFuel = $ShipData.cargo[3]
    }
    else{

        $ShipPassenger = $null
        $ShipSupplies = $null
        $ShipEngineering = $null
        $ShipFuel =$null
    }

    $CampaignShipClass = $CampaignShipClasses | Where-Object unitID -eq $ShipClassID
    $Allied = ($CampaignShipClass.Nation -in $Nations0)

    # Getting the list of homeports that can create this ship class.
    $ShipClassHomeports = @{}
    foreach ($homeport in $HomePorts.locationID){
        if ($Allied){
            if (!($homeport -in $CampaignShipClass.notAtLocation )) {
                $ShipClassHomeports.add($homeport ,"Yes") | Out-Null;
             }
             else{
                $ShipClassHomeports.add($homeport ,$null) | Out-Null;
            }

        }
        else{
             $ShipClassHomeports.add($homeport ,$null) | Out-Null;
        }

    }

     $UnavailableInstances =  new-object System.Collections.arraylist
     if ($CampaignSaveData.unavailableShipClasses.indexOf($ShipClassID) -gt -1){
         
         $unavailableInstance = 0
         foreach ($UnavailableShipClass in $CampaignSaveData.UnavailableShipClasses){
            if ($unavailableShipClass -eq $ShipClassID){
                $unavailableInstances.add($CampaignSaveData.unavailableshipInstances[$unavailableInstance])  | Out-Null;
            }
            $unavailableInstance ++;
        }
        
     }

     
     $PlayerOwnedInstances =  new-object System.Collections.arraylist
     if ($CampaignSaveData.PlayerOwnedShipClasses.indexOf($ShipClassID) -gt -1){
         
         $PlayerOwnedInstance = 0
         foreach ($PlayerOwnedShipClass in $CampaignSaveData.PlayerOwnedShipClasses){
            if ($PlayerOwnedShipClass -eq $ShipClassID){
               $PlayerOwnedInstances.add($CampaignSaveData.PlayerOwnedshipInstances[$PlayerOwnedInstance]) | Out-Null;
            }
            $PlayerOwnedInstance ++;
        }
        
     }
   
   $ShipClassAvailableDate = (get-date -Day $CampaignShipClass.available[0] -Month $CampaignShipClass.available[1] -Year $CampaignShipClass.available[2]).Date
   
   $SunkYieldArray = $CampaignShipClass.sunkyield
   $SunkYield = $SunkYieldArray[0].ToString() + "," + $SunkYieldArray[1].ToString() + "," + $SunkYieldArray[2].ToString() + "," + $SunkYieldArray[3].ToString() + "," + $SunkYieldArray[4].ToString()

  

   $Instance = 0
   foreach($Name in $ShipNamesInClass){
       $NewShip = New-object -typename PSCustomObject
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "ShipClassID" -Value  $ShipClassID
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "ShipClassName" -Value   $ShipClassName
 
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "ShipName" -Value $Name
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "ShipType" -Value  $ShipType

       add-member -InputObject $NewShip -MemberType NoteProperty -Name "Passenger" -Value $ShipPassenger
   
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "Supplies" -Value $ShipSupplies
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "Engineering" -Value $ShipEngineering
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "Fuel" -Value $ShipFuel
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "TotalShipsInClass" -Value $ShipNamesInClass.count

       $ThisShipSunk = $null;
        $sunk = $false;
       $ThisShipSunk = $sunkenships |where-object {($_.ShipClassID -eq $ShipClassID) -and ($_.ShipClassInstance -eq $Instance)}
       if ($ThisShipSunk){
             add-member -InputObject $NewShip -MemberType NoteProperty -Name "SunkDate" -Value $ThisShipSunk.sunkdate
            $sunk =  $true;
       }         
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "Sunk" -Value $sunk

       $unavailable =  (($unavailableInstances -contains $Instance) -or ($ShipClassAvailableDate -gt $CampaignDate)) 
       if ($ShipType -eq "Aircraft_Carrier" -and (!$Allied) -and (!$CampaignSaveData.enemyUsesCarriers)) {$unavailable  = $true; }  
       if ($ShipType -eq "Submarine" -and (!$Allied) -and (!$CampaignSaveData.enemyUsesSubmarines)){ $unavailable  = $true;}                                        

       add-member -InputObject $NewShip -MemberType NoteProperty -Name "Available" -Value (!$unavailable)

        $PlayerOwned =  ($PlayerOwnedInstances -contains $Instance)
        if ($PlayerOwned){
            add-member -InputObject $NewShip -MemberType NoteProperty -Name "Player Owned" -Value "Yes"
            $TaskForceName = $PlayerTaskForces |where-object {($_.ShipClassID -eq $ShipClassID) -and ($_.ShipClassInstance -eq $Instance)} |select-object TFName
            add-member -InputObject $NewShip -MemberType NoteProperty -Name "Task Force" -Value $TaskForceName.TFName;
        }
        else{
            add-member -InputObject $NewShip -MemberType NoteProperty -Name "Player Owned" -Value $null;
            add-member -InputObject $NewShip -MemberType NoteProperty -Name "Task Force" -Value $null;
           
        }
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "Nation" -Value $CampaignShipClass.Nation
       
      
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "Allied" -Value $Allied
       
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "Cost" -Value $CampaignShipClass.Cost
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "DateAvailable" -Value $ShipClassAvailableDate
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "SunkYield" -Value $SunkYield
        
       foreach ($homeport in $HomePorts){
           
            if ($Allied){
                add-member -InputObject $NewShip -MemberType NoteProperty -Name ("HomePort " + $homeport.locationName) -Value $ShipClassHomeports[$homeport.locationID]
            }
            else {
                add-member -InputObject $NewShip -MemberType NoteProperty -Name ("HomePort " + $homeport.locationName) -Value $null;
            }

       }
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "In Formation" -Value $Instance
       
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "ShipInstance" -Value $Instance

       
        $AllShips.Add($NewShip) | Out-Null;
        $Instance++;
    }

    $UnavailableInClass = $unavailableInstances.count
    [int]$AvailableInClass = $ShipNamesInClass.count - $sunkInClass.count - $UnavailableInClass
     #$AllShips |where-object {($_.shipClassID -eq $classID) -and ($_.available -eq $false)}


   # $ThisClass = $AllShips | Where-Object {($_.shipclassID -eq $classID) -and ($_.shipInstance -eq 0)}

    $NewClass = new-object  -typename PSCustomObject
    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "ID" -Value $ShipClassID
    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Class Name" -Value $ShipClassHT[$ShipClassID]
    add-member -InputObject $NewClass -MemberType NoteProperty -Name "Cost" -Value $CampaignShipClass.Cost       
    add-member -InputObject $NewClass -MemberType NoteProperty -Name "Passenger" -Value $ShipPassenger
   
    add-member -InputObject $NewClass -MemberType NoteProperty -Name "Supplies" -Value $ShipSupplies
    add-member -InputObject $NewClass -MemberType NoteProperty -Name "Engineering" -Value $ShipEngineering
    add-member -InputObject $NewClass -MemberType NoteProperty -Name "Fuel" -Value $ShipFuel

    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Sunk" -Value  $sunkInClass
    
    
    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Available" -Value  $AvailableInClass
    
    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Unavailable" -Value  $UnavailableInClass
    if ($Allied){
         add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Player Owned" -Value  $PlayerOwnedInstances.count
    }
    else{
         add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Player Owned" -Value $null;
    }
   

    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Class Total" -Value  $ShipNamesInClass.count

    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Date Available" -Value $ShipClassAvailableDate
    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Type" -Value  $ShipType
    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Nation" -Value  $CampaignShipClass.Nation
    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Allied" -Value   $Allied

    foreach ($homeport in $HomePorts){
           
        if ($Allied){
            add-member -InputObject $NewClass -MemberType NoteProperty -Name ("HomePort " + $homeport.locationName) -Value $ShipClassHomeports[$homeport.locationID]
        }
        else {
            add-member -InputObject $NewClass -MemberType NoteProperty -Name ("HomePort " + $homeport.locationName) -Value $null;
        }

    }

    $ShipClasses.add($NewClass) | Out-Null
     clear-variable -Name ShipPassenger, ShipSupplies,ShipEngineering,ShipFuel,ShipClassFile,
                        ShipClassInfo, ShipClassName, ShipNamesInClass,shipInstancesInClass,
                        ShipClassHomeports, CampaignShipClass, SunkYieldArray, UnavailableInClass,ShipType, Allied

}

$AllShips = $AllShips | Sort-Object Allied, Nation, ShipType, ShipClassName, ShipName 




$ShipClasses = $ShipClasses |sort-object Allied, Shiptype, Nation, ShipClassName
#$shipClasses |ft

#Processing Ship Types
Write-host "Processing Ship Types"  -foregroundcolor Green

$TypeTotals = new-object System.Collections.ArrayList
$ShipTypes = $AllShips | sort-object ShipType | select-object -unique ShipType
foreach ($type in $shiptypes.ShipType){

    $TotalSunkInType = 0
    $TotalUnavailableInType = 0;
    $TotalInType = 0
    $PlayerOwnedInType = 0;

    $ShipsAlliedType = $ShipClasses | where-object {$_.Allied -eq $true -and $_."Type" -eq $type}
    $ShipsEnemyType = $ShipClasses | where-object {$_.Allied -eq $false -and $_."Type" -eq $type}

    foreach ($ShipClass2 in $ShipsAlliedType){
        $TotalSunkInType += [int]$ShipClass2.Sunk
        $TotalUnavailableInType += [int]$ShipClass2.Unavailable
        $TotalInType += [int]$ShipClass2."Class Total"
        $PlayerOwnedInType += $ShipClass2."Player Owned"
        #Write-host $ShipClass2.sunkInClass $ShipClass2.TotalInClass
    }

    $NewType = new-object  -typename PSCustomObject
    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Type" -Value $type
    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Sunk" -Value  $TotalSunkInType

    $TotalCurrentLeftInType = $TotalInType - $TotalSunkIntype - $TotalUnavailableInType
    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Available" -Value  $TotalCurrentLeftInType

    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Unavailable" -Value  $TotalUnavailableInType
    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Player Owned" -Value  $PlayerOwnedInType

    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Total" -Value  $TotalInType

    add-member  -InputObject  $NewType -MemberType NoteProperty -Name "Allied" -Value  $true
    $null = $TypeTotals.add($NewType)

    $TotalSunkInType = 0
     $TotalUnavailableInType = 0;
    $TotalInType = 0

    foreach ($ShipClass3 in $ShipsEnemyType){
        #$ShipClass3.ShipclassID
        $TotalSunkInType += [int]$ShipClass3.sunk
         $TotalUnavailableInType += [int]$ShipClass3.Unavailable
        $TotalInType += [int]$ShipClass3."Class Total"
        # Write-host (" " + $ShipClass3.sunkInClass+ " "+ $ShipClass3.TotalInClass)
    }

    $NewType = new-object  -typename PSCustomObject
    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Type" -Value $type
    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Sunk" -Value  $TotalSunkInType

    $TotalCurrentLeftInType = $TotalInType - $TotalSunkIntype - $TotalUnavailableInType

    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Available" -Value  $TotalCurrentLeftInType 
    
    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Unavailable" -Value  $TotalUnavailableInType
    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Player Owned" -Value $null;
    add-member -InputObject  $NewType -MemberType NoteProperty -Name "Total" -Value  $TotalInType
    add-member  -InputObject  $NewType -MemberType NoteProperty -Name "Allied" -Value  $false
    $TypeTotals.add($NewType) | Out-Null


}
$TypeTotals = $Typetotals |sort-object Allied,Type
$TypeTotals | export-csv C:\Temp\ShipsLostByType.csv -NoTypeInformation

$ShipClasses = $ShipClasses |sort-object Allied, Type, Nation, "Class Name"
$ShipClasses | export-csv C:\Temp\ShipClasses.csv -NoTypeInformation

#$AllShips  = $AllShips | sort-object Allied, Type, ShipClassID 

#$ShipsEnemyType = $ShipClasses | where-object {$_.Allied -eq $false -and $_.ShipType -eq $type} | export-csv C:\Temp\ShipsLostByClass.csv -NoTypeInformation -Append

$AllShips |export-csv C:\Temp\AllShips.csv -NoTypeInformation

#Location Processing
Write-Host "Starting Location processing" -foregroundcolor Green

$locationData = $CampaignSaveData.mapLocationSaveData | convertfrom-json
$locationAL = new-object System.Collections.ArrayList #new-object System.Collections.SortedList

$locationColumns = @("Location ID", "Location Name", "Owned By",
                    "Port Level", "Airfield Level", "Allied Troops", "Enemy Troops", 
                    "Allied Supplies", "Allied Engineering", "Allied Fuel",
                    "Airwing Slot 1", "Airwing Slot 2", "Airwing Slot 3", "Airwing Slot 4",
                    "Airwing Replace Date", "Slot 1 Update", "Slot 2 Update", "Slot 3 Update", "Slot 4 Update")

foreach ($locationDatum in $locationData){
    
    $AirwingReplaceDate = (get-date -Year $locationDatum.airwingUpdateDate0List[0] `
                                    -Month $locationDatum.airwingUpdateDate0List[1] `
                                    -Day  $locationDatum.airwingUpdateDate0List[2] ).Date
                                    

    $NewLocation = New-object -typename PSCustomObject
    [int]$colN = -1
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[0] -Value $locationDatum.locationID 
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[1] -Value $locationDatum.locationName
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[2] -Value $locationDatum.currentFaction


    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[3] -Value  $locationDatum.portLevel
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[4] -Value  $locationDatum.airfieldLevel


    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[5] -Value  $locationDatum.troops[0]
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[6] -Value  $locationDatum.troops[1]

    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[7] -Value  $locationDatum.supplies[0]
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[8] -Value  $locationDatum.engineering[0]
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[9] -Value  $locationDatum.fuel[0]
 
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[10] -Value $AllAir[$locationDatum.aircraft0[0]]
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[11] -Value  $AllAir[$locationDatum.aircraft0[1]]
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[12] -Value $AllAir[ $locationDatum.aircraft0[2]]
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[13] -Value  $AllAir[$locationDatum.aircraft0[3]]

    
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[14] -Value $AirwingReplaceDate
   
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[15] -Value  $AllAir[$locationDatum.airwingAircraft0[0]]
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[16] -Value  $AllAir[$locationDatum.airwingAircraft0[1]]
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[17] -Value  $AllAir[$locationDatum.airwingAircraft0[2]]
    add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[18] -Value $AllAir[ $locationDatum.airwingAircraft0[3]]
  

     
     

    $locationAL.Add($NewLocation) | Out-Null;


}
$locationAL = $locationAL |sort-object "Owned By", "Location Name"


#### Excel spreadsheet
#
#
#
#
#
#

Write-Host "Processing Excel Spreadsheet" -foregroundcolor Green

$ExcelFileName = ($CampaignID +"_on_" + $campaigndate.ToString("yyMMdd") + "_at_" +  $campaigndate.ToString("HHmmss") + ".xlsx")

$ExcelObj = New-Object -comobject Excel.Application
$workbook = $excelObj.Workbooks.Add()





# All Ships worksheet
Write-host "Starting All Ships worksheet" -foregroundcolor Green
$worksheetAllShips = $workbook.Worksheets.Item(1)
$worksheetAllShips.Name = "All Ships"

$ExcelObj.ActiveWindow.SplitRow = 1
$ExcelObj.ActiveWindow.freezePanes = $true;

$worksheetAllShips.rows[1].font.Bold  = $true;
$worksheetAllShips.rows[1].HorizontalAlignment = "3"

$AllShipsColumns = new-object System.Collections.ArrayList
$AllShipsColumnsStart = @("Class", "Name","Type","Cost","Passenger Cap.", "Supplies Cap.","Engineering Cap.",
                     "Fuel Cap.", "Total Ships In Class", "Sunk", "Sunk Date",  "Available", "Player Owned", "Task Force", "Nation",
                     "Allied",  "Date Available", "Sunk Yield","Class ID", "Instance");

$AllShipsMembers = @("ShipClassName","ShipName","ShipType","Cost", "ShipPassenger","ShipCargo","ShipEngineering", "ShipFuel", 
                    "TotalShipsInClass", "Sunk", "sunkDate", "Available", "Player Owned", "Task Force", "Nation", "Allied","DateAvailable", "SunkYield","ShipClassID","ShipInstance")
                                         
foreach ($c in $AllShipsColumnsStart){
   $AllShipsColumns.add($c) | Out-Null;
}
                 
        
foreach ($homeport in $homeports){
     $AllShipsColumns.add("Homeport "+ $homeport.locationName) | Out-Null

}

[int]$i = 1;
foreach ($ShipColumn in $AllShipsColumns){
  
   $worksheetAllShips.Cells(1, $i) = $ShipColumn
   $i++;

}

$DateAVailableColumn = ($AllShipsMembers.indexOf("DateAvailable")+1)

[int]$cellRow = 2;
foreach ($ship in $AllShips){
    [int]$cellColumn = 1;

    if ($ship.sunk){
        $worksheetAllShips.Rows($CellRow).font.colorindex = 2       #white
        $worksheetAllShips.Rows($CellRow).interior.colorindex = 1  #black
    }
    if (!$ship.available){
        $worksheetAllShips.Rows($CellRow).font.colorindex = 2       #white
        $worksheetAllShips.Rows($CellRow).interior.colorindex = 16  #dark gray
           
    }
    if ($ship."Player Owned"){
        $worksheetAllShips.Rows($CellRow).style ="20% - Accent6"
       #$worksheetAllShips.Rows($CellRow).font.colorindex = 2       #white
       # $worksheetAllShips.Rows($CellRow).interior.colorindex = 43  #light green
           
    }

    foreach ($c in $AllShipsMembers){
        $worksheetAllShips.Cells($CellRow,$cellColumn) = $Ship.$c

        $cellColumn++;
    }
     foreach ($h in $Homeports){
        $worksheetAllShips.Cells($CellRow,$cellColumn) = $Ship.("Homeport "+ $h.locationName)
        $cellColumn++;
    }
    #if ($ship.DateAvailable -gt $campaignDate){
    #   $worksheetAllShips.Cells($CellRow,$DateAvailableColumn).font.colorindex = 3   #dark red
    #}
 
  $cellRow++;      
} 

 $worksheetAllShips.columns.autofit() |Out-Null;
 $worksheetAllClasses.AutoFilter


# All Classes Worksheet
Write-host "Starting All Classes worksheet" -foregroundcolor Green

$worksheetAllClasses = $workbook.worksheets.Add()
$worksheetAllClasses.Name = "All Classes"

$ExcelObj.ActiveWindow.SplitRow = 1
$ExcelObj.ActiveWindow.freezePanes = $true;

$worksheetAllClasses.rows[1].font.Bold  = $true;
$worksheetAllClasses.rows[1].HorizontalAlignment = "3"

$AllClassColumns = new-object System.Collections.ArrayList
$ClassColumns = @("ID", "Class Name", "Cost", "Passengers", "Supplies", "Engineering", "Fuel", "Sunk", "Available", "Unavailable", "Player Owned", "Class Total", "Date Available", "Type", "Nation", "Allied") 

foreach ($c in $ClassColumns){
   $AllClassColumns.add($c) | Out-Null;
}
   
foreach ($homeport in $homeports){
     $AllClassColumns.add("Homeport "+ $homeport.locationName) | Out-Null

}


 [int]$i = 1;
foreach ($c in $AllClassColumns){
  
   $worksheetAllClasses.Cells(1, $i) = $c
   
   $i++;

}
$worksheetAllClasses.rows(1).font.Bold = $true;

[int]$cellRow = 2;
foreach ($class in $ShipClasses){
    [int]$cellColumn = 1;
    if ($ship."Player Owned"){
       
        $worksheetAllClasses.Rows($CellRow).style ="20% - Accent6" #light green
           
    }
    if ($class.Available -eq 0){
        if ($class.sunk -gt 0){
                $worksheetAllClasses.Rows($CellRow).font.colorindex = 2       #white
                $worksheetAllClasses.Rows($CellRow).interior.colorindex = 1  #black
        }
        else{

            $worksheetAllClasses.Rows($CellRow).font.colorindex = 2       #white
            $worksheetAllClasses.Rows($CellRow).interior.colorindex = 16  #dark gray
        }
    }
    foreach ($c in $AllClassColumns){
        $worksheetAllClasses.Cells($CellRow,$cellColumn) = $class.$c

        $cellColumn++;
    }

  $cellRow++;      
}
$worksheetAllClasses.columns.autofit() | Out-Null
$worksheetAllClasses.AutoFilter


# All Types Sheet
Write-host "Starting All Types worksheet" -foregroundcolor Green

$worksheetAllTypes = $workbook.worksheets.Add();
$worksheetAllTypes.Name = "All Types" 

$ExcelObj.ActiveWindow.SplitRow = 1
$ExcelObj.ActiveWindow.freezePanes = $true;

$worksheetAllTypes.rows[1].font.Bold  = $true;
$worksheetAllTypes.rows[1].HorizontalAlignment = "3"

$TypeColumns = @("Type", "Sunk", "Available", "Unavailable", "Player Owned", "Total", "Allied") 
 [int]$i = 1;


foreach ($c in $TypeColumns){
  
   $worksheetAllTypes.Cells(1, $i) = $c
  
   $i++;

}

[int]$cellRow = 2;
foreach ($type in $TypeTotals){
    [int]$cellColumn = 1;

    if ($Type.currentAvailableInType -eq 0){
        if ($Type.Sunk -gt 0){
            $worksheetAllTypes.Rows($CellRow).font.colorindex = 2       #white
            $worksheetAllTypes.Rows($CellRow).interior.colorindex = 1  #black
        }
        else{
            $worksheetAllTypes.Rows($CellRow).font.colorindex = 2       #white
            $worksheetAllTypes.Rows($CellRow).interior.colorindex = 16  #dark gray
        }
    
    }
    foreach ($c in $TypeColumns){
        $worksheetAllTypes.Cells($CellRow,$cellColumn) = $Type.$c


        $cellColumn++;
    }

  $cellRow++;      
}
$worksheetAllTypes.columns.autofit()
$worksheetAllTypes.AutoFilter


# Location Sheet
Write-host "Starting Locations worksheet" -foregroundcolor Green

$worksheetLocations = $workbook.worksheets.Add();
$worksheetLocations.Name = "Locations"
$ExcelObj.ActiveWindow.SplitRow = 1
$ExcelObj.ActiveWindow.freezePanes = $true;

  [int]$i = 1;
foreach ($c in $LocationColumns){
  
   $worksheetLocations.Cells(1, $i) = $c
   $worksheetLocations.Cells(1, $i).font.bold = $true
   $i++;

}

[int]$cellRow = 2;
foreach ($location in $LocationAL){
    [int]$cellColumn = 1;
    if ($location."Owned by" -eq 0){
        $worksheetLocations.Rows($CellRow).style ="20% - Accent6"
        
        if ($location."Location ID" -in $homeports.locationID){
            $worksheetLocations.Rows($CellRow).style ="40% - Accent6"
        }
    }    

    foreach ($c in $LocationColumns){
        $worksheetLocations.Cells($CellRow,$cellColumn) = $Location.$c
                
        $cellColumn++;
    }

  $cellRow++;      
}

$worksheetLocations.columns.autofit()
$worksheetLocations.AutoFilter

# Save Info Sheet
Write-host "Starting Save Info worksheet" -foregroundcolor Green

$worksheetSaveInfo = $workbook.worksheets.Add();
$worksheetSaveInfo.Name = "Save Info"

$worksheetSaveInfo.columns(1).Font.bold = $true;
$worksheetSaveInfo.Columns(2).HorizontalAlignment = "4";

$worksheetSaveInfo.Cells(1,1) = "Campaign Name"
$worksheetSaveInfo.Cells(2,1) = "Campaign Date and Time" 
$worksheetSaveInfo.Cells(3,1) = "Campaign Day" 

$worksheetSaveInfo.Cells(5,1) = "Player Ships Sunk"
$worksheetSaveInfo.Cells(6,1) = "Enemy Ships Sunk"
$worksheetSaveInfo.Cells(7,1) = "Player Owned Locations"
$worksheetSaveInfo.Cells(8,1) = "Enemy Owned Locations"
$worksheetSaveInfo.Cells(10,1) = "Game Version" 
$worksheetSaveInfo.Cells(11,1) = "Save File Name" 

$worksheetSaveInfo.Cells(1,2) = $CampaignID
$worksheetSaveInfo.Cells(2,2) = $CampaignDate 
$worksheetSaveInfo.Cells(3,2) = $CampaignDay

$worksheetSaveInfo.Cells(5,2) = ($AllShips | Where-Object {($_.Sunk -eq $true) -and ($_.Allied  -eq $true) }).count
$worksheetSaveInfo.Cells(6,2) = ($AllShips | Where-Object {($_.Sunk -eq $true) -and ($_.Allied -eq $false) }).count
$worksheetSaveInfo.Cells(7,2) = ($locationAL | where-object "Owned By" -eq 0).count
$worksheetSaveInfo.Cells(8,2) = ($locationAL | where-object "Owned By" -eq 1).count

$worksheetSaveInfo.Cells(10,2) = $campaignSaveData.gameVersion
$worksheetSaveInfo.Cells(11,2) = $LastSavedFile.Name 

$worksheetSaveInfo.columns.autofit()

$ExcelObj.visible = $true;

$env:OneDrive

$workbook.SaveAs(($OutputPath + $ExcelFileName)) 
#$excelObj.Quit()

Write-host "Finished" -foregroundcolor Green




