#  This script generates a list of ships and locations in the most recent War On the Sea campaign file, 
#  and then gets further ship and air data from default and override folders, override taking priority.

#  Returns an Excel Spreadsheet with all the data, a csv file to show the full list of ships sunk in the game, a csv file to show ships lost by class,
#  and a csv to show the ships lost by category.

# As of May 21, 2026 this has only been tested against Kiko's new Pacific mod 1942HD running on 1.08h version of War On The Sea

# This script is written to run in PowerShell 5 without any additional modules, but does REQUIRE Microsoft Excel to get the final spreadsheet.


# Location for campaign files to test with.  Set the path to a campaign file you wish to test with.  Otherwise the script
# will get your most recent campaign save file.
# NOTE:  The script uses campaign files to flesh out its info about ships and locations so results may vary wildly if the 
# campaign save isn't for the current live campaign data, especially for mods.
$TestCampaignFilePath = $null #<Path to save file you want to test against> or set to $null to pull last saved campaign file.

#Campaign save data pulled from AppData\LocalLow\Killerfish Games\War on the Sea\save\campaign\<SAVENAME.txt>
$campaignPath =  ($env:LOCALAPPDATA + "\..\LocalLow\Killerfish Games\War on the Sea\save\campaign\")

# Paths to files we'll need from the game or mods.  
$SteamPath = (Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Valve\Steam).InstallPath
$WarOnTheSeaPath = "$SteamPath\steamapps\common\War on the Sea\WarOnTheSea_Data\StreamingAssets"
$seaUnitsOverridePath = "$WarOnTheSeaPath\override\language\english\unit\sea"
$seaUnitsDefaultPath =  "$WarOnTheSeaPath\default\language\english\unit\sea"
$seaUnitsDataDefaultPath ="$WarOnTheSeaPath\default\unit\sea\"
$seaUnitsDataOverridePath  =  "$WarOnTheSeaPath\override\unit\sea\"

$airUnitsOverridePath = "$WarOnTheSeaPath\override\language\english\unit\air"
$airUnitsDefaultPath =  "$WarOnTheSeaPath\default\language\english\unit\air"

# Working out whether to save the final Excel document in user's local folder or a OneDrive folder.  Local takes precedence
$OutputPath = ($env:Userprofile + "\Documents\")
if (!(test-path $OutputPath)){
    $OutputPath = ($env:OneDrive + "\Documents\")
    if (!(test-path $OutputPath)){
        Write-Host "Unable to find path to documents folder for final output. Exiting" -ForegroundColor Red
        Exit-PSSession
    }
}




# Get the list of save files in the campaign folder
$campaignFiles = get-childitem -Path $campaignPath 
$LastSavedFile = ($campaignFiles | where-object VersionInfo -ne $null | Sort-Object LastWriteTime -Descending)[0]

$defaultOnly = $false;

# Get the content of the last saved game file, or the test file if that's been specified.
if ($TestCampaignFilePath){
    $CampaignPathFile =  $TestCampaignFilePath 
}
else{
   $CampaignPathFile =  ($CampaignPath + $LastSavedFile.Name)
}
 $CampaignSaveData = get-content  $CampaignPathFile | ConvertFrom-Json

if ($null -eq $CampaignSaveData){
    Write-Error -Message "The campaign data could not be found at $CampaignPathFile. Ending script."
    Exit-PSSession
}



#Get the campaignID and faction sides.
# Nations0 is always the player's side.
$CampaignID = $CampaignSaveData.campaignID
$Nations0 = $CampaignSaveData.nations0
$Nations1 = $CampaignSaveData.nations1

# Getting the current time in the campaign.
$TimeArray = $CampaignSaveData.currentDatetime
$CampaignDate =  get-Date -year $TimeArray[5] -Month $TimeArray[4] -Day $TimeArray[3] -Hour $TimeArray[2] -Minute $timearray[1] -second $timeArray[0]
$CampaignDateFileFormat = $TimeArray[5]

# Mobile formation data for later use.
$PlayerMobileData = $CampaignSaveData.playerMobileObjectSaveData | convertfrom-json
$EnemyMobileData = $CampaignSaveData.enemyMobileObjectSaveData | convertfrom-json

# identifying if this is an override campaign or an original game campaign.
$campaignSeaUnitsFile = ("$WarOnTheSeaPath\override\campaign\$campaignID" + "\seaUnits.txt")
#$campaignLandLocationsFile =  ("$WarOnTheSeaPath\override\campaign\$campaignID" + "\mapLandLocations.txt")

# if we can't get the campaign's id in the override folder then this is a default game campaign.
if (!(test-path $campaignSeaUnitsFile)){
    $defaultOnly = $true;
    $campaignSeaUnitsFile = ("$WarOnTheSeaPath\default\campaign\$campaignID" + "\seaUnits.txt")
    #$campaignLandLocationsFile =  ("$WarOnTheSeaPath\default\campaign\$campaignID" + "\mapLandLocations.txt")
}

# Getting the land locations data from the save file content for good measure
$CampaignLandLocations = $CampaignSaveData.mapLocationSaveData | convertfrom-json
#Creatng a list of the player's home ports for future use.
$HomePorts =  $CampaignLandLocations| where-object CreateShips |Where-Object currentFaction -eq 0 |Select-Object locationID, locationName | sort-object

# Getting the list of aircraft to get aircraft full names if available.
Write-Host "Processing Air Unit files" -ForegroundColor Green

$AirOverrideFiles = get-childitem -Path $airUnitsOverridePath |  where-object Name -like "*.txt" |select-object Name
$AirDefaultFiles = get-childitem -Path $airUnitsDefaultPath | where-object Name -like "*.txt" |select-object Name
$AllAir = new-Object System.Collections.Hashtable

foreach ($AirFile in $AirDefaultFiles){
 
    $AirData = get-content ("$airUnitsDefaultPath\" + $AirFile.Name)  -Raw | convertFrom-json
    $AirID = $AirFile.Name.substring(0,$AirFile.Name.indexOf(".txt"))
    $AllAir.add($AirID, $AirData.unitName) | out-null;
    Clear-Variable AirData
}

# if this is an original campaign, we should ignore the overrride files.
if (!($defaultOnly)){
    foreach ($AirFile in $AirOverrideFiles){
        $AirData = get-content ("$airUnitsOverridePath\" + $AirFile.Name) -Raw | convertFrom-json
        $AirID = $AirFile.Name.substring(0, $AirFile.Name.indexOf(".txt"))
        if ($AllAir[$AirID]){
            $AllAir.Item($AirID) =  $AirData.unitName
        }
        else{
            $AllAir.add($AirID, $AirData.unitName) | Out-Null
        }
        Clear-Variable AirData
    }
}


# Get the potential list of ships from the campaignSeaUnitsFile
try{
    $CampaignShipClasses = get-content  $campaignSeaUnitsFile | ConvertFrom-Json
}
catch {
     Write-host "Could not find the list of ships from the campaign $campaignID at `n$campaignSeaUnitsFile" -ForegroundColor Red
}
if ($null -eq $CampaignShipClasses){
    Write-host "Could not find the list of ships from the campaign $campaignID at `n$campaignSeaUnitsFile" -ForegroundColor Red
}

# get the list of class files from the Sea units folder in the Override path
Write-Host "Processing sunk ships list from the campaign save file. " -ForegroundColor Green

#this line commented out because it was getting every single sea unit in the override folder and not just those
# specific to the campaign.
#$ShipClassFiles = get-childitem -Path $seaUnitsOverridePath | where-object VersionInfo -ne $null


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
$AllShips = new-object System.Collections.ArrayList;
$AllClasses  = new-object System.Collections.arraylist

foreach ($CampaignShipClass in $CampaignShipClasses){
    $sunkInClass = ($sunkenShips |where-object ShipClassID -EQ $ShipClassID).count
  
    $ShipClassID = $CampaignShipClass.unitID

    $ShipClassFile = "$ShipClassID.txt"
   # $ShipClass
   if (!($defaultOnly)){
        if( Test-path ( "$seaUnitsOverridePath\$ShipClassFile")){
            try{
                $ShipClassFileContent = Get-Content -Path "$seaUnitsOverridePath\$ShipClassFile" -Raw
               # $ShipClassFileContent = $ShipClassFileContent.replace('''', '''''')
                $ShipClassInfo = $ShipClassFileContent.substring($ShipClassFileContent.IndexOf('{'), $ShipClassFileContent.LastIndexOf('}')+1) | ConvertFrom-Json
            }
            catch{
                Write-host "Error on JSON conversion for Override file $ShipClassFile"

            }
        }
        else{
            if (Test-path  "$seaUnitsDefaultPath\$ShipClassFile" ){
                try{
                    $ShipClassFileContent = Get-Content -Path "$seaUnitsDefaultPath\$ShipClassFile" -Raw
                    $ShipClassInfo = $ShipClassFileContent.substring($ShipClassFileContent.IndexOf('{'), $ShipClassFileContent.LastIndexOf('}')+1) | ConvertFrom-Json
                }
                catch{
                    Write-host "Error on JSON conversion for Default file $ShipClassFile"
                }   
            
            }
            else{
                Write-host "$ShipClassID unit info not found"
                
            }
        }

    }   
    else{
        if (Test-path  "$seaUnitsDefaultPath\$ShipClassFile" ){
            try{
                $ShipClassFileContent = Get-Content -Path "$seaUnitsDefaultPath\$ShipClassFile" -Raw
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
   $ShipData = $ShipDataRaw[0] | ConvertFrom-Json
   
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

   # $CampaignShipClass = $CampaignShipClasses | Where-Object unitID -eq $ShipClassID
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

       add-member -InputObject $NewShip -MemberType NoteProperty -Name "Passengers" -Value $ShipPassenger
   
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
    
       
       add-member -InputObject $NewShip -MemberType NoteProperty -Name "ShipInstance" -Value $Instance

       
        $AllShips.Add($NewShip) | Out-Null;
        $Instance++;
    }

    $UnavailableInClass = $unavailableInstances.count
    [int]$AvailableInClass = $ShipNamesInClass.count - $sunkInClass - $UnavailableInClass
     #$AllShips |where-object {($_.shipClassID -eq $classID) -and ($_.available -eq $false)}


   # $ThisClass = $AllShips | Where-Object {($_.shipclassID -eq $classID) -and ($_.shipInstance -eq 0)}

    $NewClass = new-object  -typename PSCustomObject
    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "ID" -Value $ShipClassID
    add-member -InputObject  $NewClass -MemberType NoteProperty -Name "Class Name" -Value $ShipClassHT[$ShipClassID]
    add-member -InputObject $NewClass -MemberType NoteProperty -Name "Cost" -Value $CampaignShipClass.Cost       
    add-member -InputObject $NewClass -MemberType NoteProperty -Name "Passengers" -Value $ShipPassenger
   
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

    $AllClasses.add($NewClass) | Out-Null
     clear-variable -Name ShipPassenger, ShipSupplies,ShipEngineering,ShipFuel,ShipClassFile,
                        ShipClassInfo, ShipClassName, ShipNamesInClass,shipInstancesInClass,
                        ShipClassHomeports, CampaignShipClass, SunkYieldArray, UnavailableInClass,ShipType, Allied

}

$AllShips = $AllShips | Sort-Object Allied, Nation, ShipType, ShipClassName, ShipName 




$AllClasses = $AllClasses |sort-object Allied, Shiptype, Nation, ShipClassName
#$AllClasses |ft

#Processing Ship Types
Write-host "Processing Ship Types"  -foregroundcolor Green

$AllTypes = new-object System.Collections.ArrayList
$ShipTypes = $AllShips | sort-object ShipType | select-object -unique ShipType
foreach ($type in $shiptypes.ShipType){

    $TotalSunkInType = 0
    $TotalUnavailableInType = 0;
    $TotalInType = 0
    $PlayerOwnedInType = 0;

    $ShipsAlliedType = $AllClasses | where-object {$_.Allied -eq $true -and $_."Type" -eq $type}
    $ShipsEnemyType = $AllClasses | where-object {$_.Allied -eq $false -and $_."Type" -eq $type}

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
    $null = $AllTypes.add($NewType)

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
    $AllTypes.add($NewType) | Out-Null


}
$AllTypes = $AllTypes |sort-object Allied,Type
$AllTypes | export-csv C:\Temp\ShipsLostByType.csv -NoTypeInformation

$AllClasses = $AllClasses |sort-object Allied, Type, Nation, "Class Name"
$AllClasses | export-csv C:\Temp\ShipClasses.csv -NoTypeInformation

#$AllShips  = $AllShips | sort-object Allied, Type, ShipClassID 

#$ShipsEnemyType = $AllClasses | where-object {$_.Allied -eq $false -and $_.ShipType -eq $type} | export-csv C:\Temp\ShipsLostByClass.csv -NoTypeInformation -Append

$AllShips |export-csv C:\Temp\AllShips.csv -NoTypeInformation

#Location Processing
Write-Host "Starting Location processing" -foregroundcolor Green

$locationData = $CampaignSaveData.mapLocationSaveData | convertfrom-json
$AllLocations = new-object System.Collections.ArrayList #new-object System.Collections.SortedList

$locationColumns = @("Location ID", "Location Name", "Owned By",
                    "Port Level", "Airfield Level", "Allied Troops", "Enemy Troops", 
                    "Allied Supplies", "Allied Engineering", "Allied Fuel",
                    "Airwing Slot 1", "Airwing Slot 2", "Airwing Slot 3", "Airwing Slot 4",
                    "Airwing Replace Date", "Slot 1 Update", "Slot 2 Update", "Slot 3 Update", "Slot 4 Update")


foreach ($locationDatum in $locationData){
    
    if ($null -eq $locationDatum.airwingUpdateDate0List){
         $AirwingReplaceDate = $null;
    }
    else{
        $AirwingReplaceDate = (get-date -Year $locationDatum.airwingUpdateDate0List[0] `
                                    -Month $locationDatum.airwingUpdateDate0List[1] `
                                    -Day  $locationDatum.airwingUpdateDate0List[2] ).Date
    }                           

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
 
    for ([int]$i =0; $i -lt $locationDatum.aircraft0.count; $i++){
        add-member -InputObject $NewLocation -MemberType NoteProperty -Name ("Airwing Slot " + ($i+1)) -Value $AllAir[$locationDatum.aircraft0[$i]]
  
    }
 # add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[11] -Value $AllAir[$locationDatum.aircraft0[1]]
  #  add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[12] -Value $AllAir[$locationDatum.aircraft0[2]]
 #   add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[13] -Value $AllAir[$locationDatum.aircraft0[3]]

    if ($AirwingReplaceDate){
        add-member -InputObject $NewLocation -MemberType NoteProperty -Name "Airwing Replace Date" -Value $AirwingReplaceDate
    }
   # add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[14] -Value $AirwingReplaceDate
   
    for ([int]$j =0; $j -lt $locationDatum.airwingAircraft0.count; $j++){
        add-member -InputObject $NewLocation -MemberType NoteProperty -Name ("Slot " + ($j+1) + " Upgrade") -Value $AllAir[$locationDatum.airwingAircraft0[$j]]

    }
   # add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[15] -Value $AllAir[$locationDatum.airwingAircraft0[0]]
   # add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[16] -Value $AllAir[$locationDatum.airwingAircraft0[1]]
   # add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[17] -Value $AllAir[$locationDatum.airwingAircraft0[2]]
   # add-member -InputObject $NewLocation -MemberType NoteProperty -Name $locationColumns[18] -Value $AllAir[$locationDatum.airwingAircraft0[3]]
  

     
     

    $AllLocations.Add($NewLocation) | Out-Null;


}
<#
$LocationCount = $AllLocations.count
$arrayCheck = @(13..18)
foreach($i in $arrayCheck){
    $NullCount = ($AllLocations | where-object $locationColumns[$i] -eq $null ).count
    if ($NullCount -eq $LocationCount){
        $AllLocations
    }
}#>
$AllLocations = $AllLocations |sort-object "Owned By", "Location Name"

$AllClasses | export-csv C:\Temp\Locations.csv -NoTypeInformation
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

if ($AllShips){
    $ExcelObj.ActiveWindow.SplitRow = 1
    $ExcelObj.ActiveWindow.freezePanes = $true;

    $worksheetAllShips.rows[1].font.Bold  = $true;
    $worksheetAllShips.rows[1].HorizontalAlignment = "3"

    $AllShipsColumns = new-object System.Collections.ArrayList
    $AllShipsColumnsStart = @("Class", "Name","Type","Cost","Passengers", "Supplies","Engineering",
                        "Fuel", "Total Ships In Class", "Sunk", "Sunk Date",  "Available", "Player Owned", "Task Force", "Nation",
                        "Allied",  "Date Available", "Sunk Yield","Class ID", "Instance");

    $AllShipsMembers = @("ShipClassName","ShipName","ShipType","Cost", "Passengers","Supplies","Engineering", "Fuel", 
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
}
else{
    $worksheetAllShips.Cells(1,1)  = "Error:  No campaign ship data found."
    $worksheetAllShips.Cells(1,1).font.colorindex = "3"  # setting color to red
    Write-host "No ship classes found so worksheet not generated." -foregroundcolor Red
}

# All Classes Worksheet
Write-host "Starting All Classes worksheet" -foregroundcolor Green
if ($AllClasses){
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
    foreach ($class in $AllClasses){
        [int]$cellColumn = 1;
        if ($class."Player Owned"){
        
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
        if ($class.Available -eq -1){
            $worksheetAllClasses.Rows($CellRow).font.italic = $true;
        }
        # class should not be available yet.
        if ($class.'Date Available' -gt $CampaignDate){
            $worksheetAllClasses.Rows($CellRow).font.colorindex = 2       #white
            $worksheetAllClasses.Rows($CellRow).interior.colorindex = 16  #dark gray

        }
        foreach ($c in $AllClassColumns){
            $worksheetAllClasses.Cells($CellRow,$cellColumn) = $class.$c

            $cellColumn++;
        }

    $cellRow++;      
    }
    $worksheetAllClasses.columns.autofit() | Out-Null
    $worksheetAllClasses.AutoFilter
}
else{
    Write-host "No ship classes found so worksheet not generated." -foregroundcolor Red
}

# All Types Sheet
Write-host "Starting All Types worksheet" -foregroundcolor Green
if ($AllTypes){
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
    foreach ($type in $AllTypes){
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
    $worksheetAllTypes.columns.autofit() | Out-Null;
    $worksheetAllTypes.AutoFilter
}
else{
    Write-host "No ship types found so worksheet not generated." -foregroundcolor Red
}

# Location Sheet
Write-host "Starting Locations worksheet" -foregroundcolor Green
if ($AllLocations){

    $worksheetLocations = $workbook.worksheets.Add();
    $worksheetLocations.Name = "Locations"
    $worksheetLocations.AutoFilter
    $ExcelObj.ActiveWindow.SplitRow = 1
    $ExcelObj.ActiveWindow.freezePanes = $true;

    # First row
    [int]$i = 1;
    $LocationColumnNames = $AllLocations[0].psobject.Properties |where-object MemberType -eq "NoteProperty" |select-object Name
    foreach ($c in $LocationColumnNames){
    
    $worksheetLocations.Cells(1, $i) = $c.Name
    $worksheetLocations.Cells(1, $i).font.bold = $true
    $i++;

    }

    [int]$cellRow = 2;
    foreach ($location in $AllLocations){
        [int]$cellColumn = 1;
        if ($location."Owned by" -eq 0){
            $worksheetLocations.Rows($CellRow).style ="20% - Accent6"
            
            if ($location."Location ID" -in $homeports.locationID){
                $worksheetLocations.Rows($CellRow).style ="40% - Accent6"
            }
        }    

        foreach ($c in $LocationColumnNames){
            $worksheetLocations.Cells($CellRow,$cellColumn) = $Location.($c.Name)
                    
            $cellColumn++;
        }

    $cellRow++;      
    }

    $worksheetLocations.columns.autofit() | Out-Null;
}
else{
    Write-host "No Locations found so worksheet not generated." -foregroundcolor Red
}

# Save Info Sheet
Write-host "Starting Save Info worksheet" -foregroundcolor Green

$worksheetSaveInfo = $workbook.worksheets.Add();
$worksheetSaveInfo.Name = "Save Info"

$worksheetSaveInfo.columns(1).Font.bold = $true;


$worksheetSaveInfo.Cells(1,1) = "Campaign Name"
$worksheetSaveInfo.Cells(2,1) = "Campaign Date and Time" 
$worksheetSaveInfo.Cells(3,1) = "Campaign Day" 

$worksheetSaveInfo.Cells(5,1) = "Player Ships Sunk"
$worksheetSaveInfo.Cells(6,1) = "Enemy Ships Sunk"
$worksheetSaveInfo.Cells(7,1) = "Player Owned Locations"
$worksheetSaveInfo.Cells(8,1) = "Enemy Owned Locations"
$worksheetSaveInfo.Cells(10,1) = "Game Version" 
$worksheetSaveInfo.Cells(11,1) = "Campaign File Path" 

$worksheetSaveInfo.Cells(1,2) = $CampaignID
$worksheetSaveInfo.Cells(2,2) = $CampaignDate 
$worksheetSaveInfo.Cells(3,2) = $CampaignDay

$worksheetSaveInfo.Cells(5,2) = ($AllShips | Where-Object {($_.Sunk -eq $true) -and ($_.Allied  -eq $true) }).count
$worksheetSaveInfo.Cells(6,2) = ($AllShips | Where-Object {($_.Sunk -eq $true) -and ($_.Allied -eq $false) }).count
$worksheetSaveInfo.Cells(7,2) = ($AllLocations | where-object "Owned By" -eq 0).count
$worksheetSaveInfo.Cells(8,2) = ($AllLocations | where-object "Owned By" -eq 1).count

$worksheetSaveInfo.Cells(10,2) = $campaignSaveData.gameVersion
$worksheetSaveInfo.Cells(11,2) = $CampaignPathFile

$worksheetSaveInfo.Columns(2).HorizontalAlignment = "2";  #left aligning the column
$worksheetSaveInfo.columns.autofit() | Out-Null;

$ExcelObj.visible = $true;



$workbook.SaveAs(($OutputPath + $ExcelFileName)) 
#$excelObj.Quit()

Write-host "Finished" -foregroundcolor Green




