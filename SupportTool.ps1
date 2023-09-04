# v1.1 DummyFix
#Forces powershell to run as an admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{ Start-Process powershell.exe "-NoProfile -Windowstyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

#Imports Windowsforms and Drawing from system
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#Allows the use of wshell for confirmation popups
$wshell = New-Object -ComObject Wscript.Shell

#Links functions to selected option in the dropdown list, activates on button click
#Outputbox.clear() Erases text output from the outputbox before continuing with the script.
Function selectedscript {

    if ($DropDownBox.Selecteditem -eq "Remove PCEye5 Bundle") {
    $Outputbox.Clear()
    UninstallPCEye5Bundle
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove Tobii Device Drivers For Windows") {
    $Outputbox.Clear()
    UninstallTobiiDeviceDriversForWindows
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove WC&GP Bundle") {
    $Outputbox.Clear()
    UninstallWCGP
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove PCEye Package") {
    $Outputbox.Clear()
    UninstallPCeyePackage
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove Communicator") {
    $Outputbox.Clear()
    UninstallCommunicator
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove Compass") {
    $Outputbox.Clear()
    UninstallCompass
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove TGIS only") {
    $Outputbox.Clear()
    UninstallTGIS
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove TGIS profile calibrations") {
    $Outputbox.Clear()
    TGISProfilesremove
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove all users C5") {
    $Outputbox.Clear()
    DeleteC5User
    }
    elseif ($DropDownBox.Selecteditem -eq "Reset TETC") {
    $Outputbox.Clear()
    ResetTETC
    }
    else {
    $Outputbox.AppendText( "" )
    $OutputBox.AppendText( "No option selected. `r`n" )
    Return
    }
}

#A1 Uninstalls PCEye5 Bundle
Function UninstallPCEye5Bundle {

    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove all software included in PCeye5Bundle.`r`nAre you sure you want to continue?`r`n",0,"Caution",48+4)
    if($answer1 -eq 6){$Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif($answer1 -ne 6){$Outputbox.Appendtext( "Action canceled: Remove PCEye5Bundle`r`n" )
    Return
    }

    $Outputbox.Appendtext( "Please wait.`r`n" )
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object {($_.Displayname -Match "Tobii Dynavox Computer Control") -or
    ($_.Displayname -Match "Dynavox Computer Control Updater Service") -or
    ($_.Displayname -Match "Tobii Dynavox Update Notifier") -or
    ($_.Displayname -Match "Tobii Dynavox Eye Tracking") -or
    ($_.Displayname -Eq "Tobii Device Drivers For Windows (PCEye5)") } | Select-Object Displayname, UninstallString
    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
    $Uninstname = $ver.Displayname
    $uninst = $ver.UninstallString -replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
    $uninst = $uninst.Trim()
    $Outputbox.Appendtext( "Removing - "+ "$Uninstname`r`n" )
    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
    }
    $Outputbox.appendtext( "Done! `r`n" )

    $answer3 = $wshell.Popup("Do you want to remove related folders and registry keys? (This includes usersettings and profiles)",0,"Caution",48+4)
    if($answer3 -eq 6){

    $Outputbox.appendtext( "Looking for related folders...`r`n" )
    #Removes WC related folders
    $paths = (
    "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist",
    "$ENV:ProgramData\Tobii Dynavox\EyeAssist",
    "$ENV:ProgramData\HelloDMFT",
    "C:\Program Files (x86)\Tobii Dynavox\Eye Assist",
    "C:\Program Files (x86)\Tobii Dynavox\Eye Tracking Settings")

    foreach($path in $paths)
    {
    if (Test-Path $path)
    {
    $Outputbox.appendtext( "Removing - "+"$path`r`n" )
    Remove-Item $path -Recurse -Force -ErrorAction Ignore
    }
    }
    $Keys = ("HKCU:\Software\Tobii\EyeAssist")

    foreach($Key in $Keys)
    {
    if (test-path $Key)
    {
    $Outputbox.appendtext( "Removing - "+"$Key`r`n" )
    Remove-item $Key -Recurse -ErrorAction Ignore
    }
    }

    $Outputbox.Appendtext( "Done!`r`n" )
    }

    elseif($answer3 -ne 6){$Outputbox.Appendtext( "Action canceled: Remove WC&GP folders and registry keys.`r`n" )
    }
    $Outputbox.Appendtext( "Finished!`r`n" )
}

#A2 Uninstalls ALL Tobii Device Drivers For Windows Bundle
Function UninstallTobiiDeviceDriversForWindows {

    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove all software included in Tobii Device Drivers For Windows Bundles.`r`nAre you sure you want to continue?`r`n",0,"Caution",48+4)
    if($answer1 -eq 6){$Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif($answer1 -ne 6){$Outputbox.Appendtext( "Action canceled: Remove Tobii Device Drivers For Windows`r`n" )
    Return
    }

    $Outputbox.Appendtext( "Please wait.`r`n" )
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object {($_.Displayname -Match "Tobii Device Drivers For Windows")} | Select-Object Displayname, UninstallString
    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
    $Uninstname = $ver.Displayname
    $uninst = $ver.UninstallString
    $Outputbox.Appendtext( "Removing - "+ "$Uninstname`r`n" )

    $uninst = $ver.UninstallString -replace "msiexec.exe","" -Replace "/I","" -Replace "/X","" -replace "/uninstall",""
    $uninst = $uninst.Trim()
    if ($uninst -notcontains 'ProgramData') {
    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
    } else {
    cmd /c $uninst /uninstall /quiet
    }
    }
    $Outputbox.appendtext( "Done!`r`n" )
    $Outputbox.Appendtext( "Finished!`r`n" )
}

#A3 Uninstalls WC Bundle
Function UninstallWCGP {

    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove all software included in Windows Control & Gaze Point Bundles.`r`nAre you sure you want to continue?`r`n",0,"Caution",48+4)
    if($answer1 -eq 6){$Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif($answer1 -ne 6){$Outputbox.Appendtext( "Action canceled: Remove WC&GP`r`n" )
    Return
    }

    #If second answer equals yes or no
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?",0,"Caution",48+4)
    if($answer2 -eq 6){CopyLicenses}

    elseif($answer2 -ne 6){$Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" )
    }

    $Outputbox.Appendtext( "Please wait.`r`n" )
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object {($_.Displayname -Match "Windows Control") -or
    ($_.Displayname -Match "Virtual Remote") -or
    ($_.Displayname -Match "Update Notifier") -or
    ($_.Displayname -Match "Tobii Eye Tracking") -or
    ($_.Displayname -Match "GazeSelection") -or
    ($_.Displayname -Match "Tobii Dynavox Gaze Point") -or
    ($_.Displayname -Match "Tobii Dynavox Gaze Point Configuration Guide")} | Select-Object Displayname, UninstallString
    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
    $Uninstname = $ver.Displayname
    $uninst = $ver.UninstallString
    $Outputbox.Appendtext( "Removing - "+ "$Uninstname`r`n" )
    & cmd /c $uninst /quiet /norestart
    }
    $Outputbox.appendtext( "Done!`r`n" )


    $answer3 = $wshell.Popup("Do you want to remove related folders and registry keys? (This includes usersettings and profiles)",0,"Caution",48+4)
    if($answer3 -eq 6){

    $Outputbox.appendtext( "Looking for related folders...`r`n" )
    #Removes WC related folders
    $paths = ( "$Env:USERPROFILE\AppData\Roaming\Tobii\Tobii Interaction\",
    "$Env:USERPROFILE\AppData\Roaming\Tobii\Tobii Interaction Statistics\",
    "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist",
    "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Gaze Selection",
    "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Windows Control Bundle",
    "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Gaze Point Bundle",
    "$Env:USERPROFILE\AppData\Local\Tobii\Tobii Interaction\",
    "$ENV:ProgramData\Tobii Dynavox\Tobii Interaction\ScreenPlanes\",
    "$ENV:ProgramData\TetServer",
    "$ENV:ProgramData\Tobii Dynavox\Windows Control Configuration Guide\",
    "C:\Program Files (x86)\Tobii Dynavox\Windows Control Configuration Guide",
    "C:\Program Files (x86)\Tobii Dynavox\Update Notifier",
    "C:\Program Files (x86)\Tobii\Service\Plugins",
    "$ENV:ProgramData\Tobii\Statistics\",
    "$ENV:ProgramData\Tobii\Tobii Interaction\",
    "$ENV:ProgramData\Tobii\Tobii Stream Engine\",
    "$ENV:ProgramData\Tobii Dynavox\Update Notifier\",
    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Update Notifier\",
    "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control\",
    "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Point\" )

    foreach($path in $paths)
    {
    if (Test-Path $path)
    {
    $Outputbox.appendtext( "Removing - "+"$path`r`n" )
    Remove-Item $path -Recurse -Force -ErrorAction Ignore
    }
    }
    $Outputbox.Appendtext( "Done!`r`n" )
    $Outputbox.appendtext( "Looking for related Registry keys...`r`n" )

    #Deleting registry keys related to WC
    $Keys = ( "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeX",
    "HKLM:\Software\WOW6432Node\Tobii\EyeXConfig\",
    "HKLM:\Software\Wow6432Node\Tobii\TobiiUpdater\",
    "HKLM:\Software\Wow6432Node\Tobii\Update Notifier\" )

    foreach($Key in $Keys)
    {
    if (test-path $Key)
    {
    $Outputbox.appendtext( "Removing - "+"$Key`r`n" )
    Remove-item $Key -Recurse -ErrorAction Ignore
    }
    }

    }

    elseif($answer3 -ne 6){$Outputbox.Appendtext( "Action canceled: Remove WC&GP folders and registry keys.`r`n" )
    }
    $Outputbox.Appendtext( "Finished!`r`n" )
}

#A4
Function UninstallPCEyePackage {
    #Implement functionality. (PCEye package & TGIS on i-series, start with PCEye package
    $answer1 = $wshell.Popup("This will remove all software included in PCEye Package`r`nAre you sure you want to continue?`r`n",0,"Caution",48+4)
    if($answer1 -eq 6){$Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif($answer1 -ne 6){$Outputbox.Appendtext( "Action canceled: Remove PCEye Package`r`n" )
    Return
    }

    #If second answer equals yes or no
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?",0,"Caution",48+4)
    if($answer2 -eq 6){CopyLicenses}

    elseif($answer2 -ne 6){$Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" )
    }

    $Outputbox.Appendtext( "Please wait.`r`n" )


    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object {($_.Displayname -Match "Tobii Dynavox Gaze Interaction Software") -or
    ($_.Displayname -Match "Tobii Dynavox PCEye Update Notifier") -or
    ($_.Displayname -Match "Tobii Dynavox Gaze Selection Language Packs") -or
    ($_.Displayname -Match "Tobii IS3 Eye Tracker Driver") -or
    ($_.Displayname -Match "Tobii IS4 Eye Tracker Driver") -or
    ($_.Displayname -Match "Tobii Eye Tracker Browser") -or
    ($_.Displayname -Match "Tobii Dynavox PCEye Configuration Guide") -or
    ($_.Displayname -Match "Tobii Dynavox Gaze HID")} | Select-Object Displayname, UninstallString

    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
    $Uninstname = $ver.Displayname
    $uninst = $ver.UninstallString
    $Outputbox.Appendtext( "Removing - "+ "$Uninstname`r`n" )
    & cmd /c $uninst /quiet /norestart
    }

    $UninstallService = Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -match "Tobii Service"}

    ForEach ($Software in $UninstallService) {
    $Uninstname2 = $Software.Name
    $Outputbox.Appendtext( "Removing - "+ "$Uninstname2`r`n")
    $Software.Uninstall()
    }

    $Outputbox.appendtext( "Done!`r`n" )


    $answer3 = $wshell.popup("Do you want to remove related folders and registrykeys?", 0, "Caution", 48+4)
    if($answer3 -eq 6){

    $paths = ( "$ENV:ProgramData\Tobii Dynavox\PCEye Configuration Guide",
    "$ENV:AppData\Tobii Dynavox\PCEye Configuration Guide",
    "$ENV:AppData\Tobii Dynavox\PCEye Update Notifier\",
    "$ENV:ProgramData\Tobii Dynavox\Gaze Interaction\Server",
    "$ENV:ProgramData\Tobii Dynavox\PCEye Update Notifier",
    "$ENV:ProgramData\Tobii\Statistics\",
    "$ENV:ProgramData\Tobii\Tobii Interaction",
    "$ENV:ProgramData\Tobii\Tobii Stream Engine\odin",
    "$ENV:ProgramData\TetServer",
    "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction" )

    $Outputbox.Appendtext( "Looking for folders to remove...`r`n" )
    foreach($path in $paths)
    {
    if (Test-Path $path)
    {
    $Outputbox.appendtext( "Removing - "+"$path`r`n" )
    Remove-Item $path -Recurse -Force -ErrorAction Ignore
    }
    }

    $Key = "HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation"

    if (test-path $Key)
    {
    $Outputbox.appendtext( "Removing - "+"$Key`r`n" )
    Remove-Item $Key -Force -ErrorAction ignore
    }
    }

    elseif($answer3 -ne 6){$Outputbox.appendtext( "Action canceled: Remove keys & folders`r`n" )}

    $Outputbox.appendtext( "Finished!`r`n" )
}

#A5 Uninstalls all Tobii related software
Function UninstallCommunicator {
    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will uninstall Communicator. Are you sure you want to continue?`r`n",0,"Caution",48+4)
    if($answer1 -eq 6){$Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif($answer1 -ne 6){$Outputbox.Appendtext( "Action canceled: Remove Communicator`r`n" )
    Return
    }

    #If second answer equals yes or no - if "Yes" then it will call the function CopyLicenses and then continue.
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?",0,"Caution",48+4)
    if($answer2 -eq 6){CopyLicenses}

    elseif($answer2 -ne 6){$Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" )}

    $Outputbox.Appendtext( "Please wait.`r`n" )

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object {$_.Displayname -match "Tobii Dynavox Communicator"} | Select-Object Publisher, Displayname, UninstallString

    $Outputbox.appendtext( "Starting uninstallation...`r`n" )

    ForEach ($ver in $TobiiVer) {
    $Uninstname = $ver.Displayname
    $Outputbox.Appendtext( "Removing - "+ "$Uninstname`r`n" )
    $uninst = $ver.UninstallString
    & cmd /c $uninst /quiet /norestart
    }
    $Outputbox.appendtext( "Done!`r`n" )

    $answer3 = $wshell.Popup("Do you want to remove related Registrykeys & folders?",0,"Caution",48+4)
    if($answer3 -eq 6){$Outputbox.appendtext( "Looking for Communicator keys & folders...`r`n" )

    $paths = ( "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Communicator",
    "$ENV:ProgramData\Tobii Dynavox\Communicator" )

    foreach($path in $paths)
    {
    if (Test-Path $path)
    {
    $Outputbox.AppendText( "Removing - "+ "$path`r`n")
    Remove-Item $path -Recurse -Force -ErrorAction Ignore
    }
    }


    $Keys = ("HKLM:\SOFTWARE\WOW6432Node\Tobii\MyTobii\MPA\VS Communicator 4",
    "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4",
    "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5" )

    foreach($Key in $Keys)
    {
    if (test-path $Key)
    {
    $Outputbox.appendtext( "Removing - "+"$Key`r`n" )
    Remove-item $Key -Recurse -ErrorAction Ignore
    }
    }

    }

    elseif($answer3 -ne 6){$Outputbox.Appendtext( "Action canceled: Remove Communicator keys & folders.`r`n" )}
    $Outputbox.Appendtext( "Finished!`r`n" )
}

#A6 Uninstalls only Compass
Function UninstallCompass {
    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will uninstall Compass. Are you sure you want to continue?`r`n",0,"Caution",48+4)
    if($answer1 -eq 6){$Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )
    }
    elseif($answer1 -ne 6){$Outputbox.Appendtext( "Action canceled: Remove Compass`r`n" )
    Return
    }

    $Outputbox.Appendtext( "Please wait.`r`n" )


    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object {($_.Displayname -Match "Tobii Dynavox Compass")} | Select-Object Displayname, UninstallString

    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
    $Uninstname = $ver.Displayname
    $Outputbox.Appendtext( "Removing - "+ "$Uninstname" )
    $uninst = $ver.UninstallString
    & cmd /c $uninst /quiet /norestart
    }

    $Outputbox.Appendtext( "Done!`r`n" )


    $answer2 = $wshell.Popup("Do you want to remove related folders?",0,"Caution",48+4)
    if($answer2 -eq 6){$Outputbox.appendtext( "Looking for related folders...`r`n" )

    $Keys = ( "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Compass" )

    foreach($Key in $Keys)
    {
    if (test-path $Key)
    {
    $Outputbox.appendtext( "Removing - "+"$Key`r`n" )
    Remove-item $Key -Recurse -ErrorAction Ignore
    }
    }
    }
    elseif($answer2 -ne 6){$Outputbox.appendtext( "Action canceled: Remove folders`r`n" )
    }

    $Outputbox.appendtext( "Finished!`r`n" )
}

#A7
Function UninstallTGIS {
    $answer1 = $wshell.Popup("This will ONLY remove Tobii Gaze Interaction Software.`r`nAre you sure you want to continue?`r`n",0,"Caution",48+4)
    if($answer1 -eq 6){$Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )


    #If second answer equals yes or no
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?",0,"Caution",48+4)
    if($answer2 -eq 6){CopyLicenses}
    elseif($answer2 -ne 6){$Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" )
    }
    $Outputbox.Appendtext( "Please wait.`r`n" )
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object {($_.Displayname -Match "Tobii Dynavox Gaze Interaction Software")} | Select-Object Displayname, UninstallString
    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
    $Uninstname = $ver.Displayname
    $uninst = $ver.UninstallString
    $Outputbox.Appendtext( "Removing - "+ "$Uninstname`r`n" )
    & cmd /c $uninst /quiet /norestart
    }

    $Outputbox.appendtext( "Done!`r`n" )


    $paths = ("$env:ProgramData\Tobii Dynavox\Gaze Interaction\",
    "$ENV:ProgramData\Tobii Dynavox\Gaze Selection\Word Prediction\Language Packs\")

    $answer3 = $wshell.Popup("Do you want to remove related folders?",0,"Caution",48+4)

    if($answer3 -eq 6){$Outputbox.Appendtext( "Looking for related folders...`r`n" )
    foreach($path in $paths)
    {
    if (Test-Path $path)
    {
    $Outputbox.appendtext( "Removing - "+"$path`r`n" )
    Remove-Item $path -Recurse -Force -ErrorAction Ignore
    }
    }
    }
    elseif($answer3 -ne 6) {$Outputbox.appendtext( "Action canceled: Remove folders`r`n" )
    }

    $Outputbox.appendtext( "Finished!`r`n" )


    }
    elseif($answer1 -ne 6){$Outputbox.Appendtext( "Action canceled: Remove TGIS`r`n" )
    Return
    }
}

#A8 Function for the option "Remove TGIS calibration profiles #Tobii service is stopped
Function TGISProfilesremove {

    $answer1 = $wshell.Popup("This will remove ONLY calibrations for every profile, it will NOT remove the actual profiles. The Gaze Interaction software will close and tobii service will restart.`r`nContinue?",0,"Caution",48+4)
    if($answer1 -eq 6){

    $outputbox.appendtext( "Shutting down TGIS software...`r`n" )
    $Outputbox.appendtext( "Stopping processes:`r`n" )
    $Processkill = get-process "Tobii.Service", "TobiiEyeControlOptions", "TobiiEyeControlServer", "Notifier" | Stop-process -force -Passthru -erroraction ignore | Select Processname |
    Format-table -Hidetableheaders | Out-string
    $Outputbox.Appendtext($Processkill)

    $outputbox.appendtext( "Looking for calibration profiles...`r`n" )
    $paths = ( "$ENV:ProgramData\Tobii Dynavox\Gaze Interaction\Server\Calibration\*" )

    foreach($path in $paths)
    {
    if (Test-Path $path)
    {
    remove-Item $path -Recurse -Force -ErrorAction Ignore
    $Outputbox.appendtext("Calibrations found! - Removing...`r`n" )
    }
    else {
    $Outputbox.Appendtext( "No calibration profiles were found!`r`n" )
    }
    }
    $outputbox.Appendtext( "Attempting to start Tobii Service... `r`n" )
    try {
    Start-Service -Name "Tobii Service" -ErrorAction Stop
    sleep 1
    $Outputbox.Appendtext( "Tobii Service started! `r`n")
    }
    Catch {
    $Outputbox.Appendtext( "Tobii Service failed to start!`r`n" )
    }
    }
    elseif($answer1 -ne 6){$outputBox.appendtext( "Action canceled: Remove calibration profiles." )
    }
    $outputbox.appendtext( "Finished!`r`n" )
}

#A9
Function DeleteC5User {
    $outputBox.clear()
    $outputBox.appendtext( "Deleting C5 users.`r`n" )
    $paths = ( "$env:USERPROFILE\Documents\Communicator 5",
    "$env:USERPROFILE\AppData\Local\VirtualStore\Program Files (x86)\Tobii Dynavox\Communicator 5",
    "$env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Communicator",
    "$env:ProgramData\Tobii Dynavox\Communicator")
    foreach($path in $paths)
    {
    if (Test-Path $path)
    {
    $Outputbox.appendtext( "Removing - "+"$path`r`n" )
    Remove-Item $path -Recurse -Force -ErrorAction Ignore
    }
    }
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#A10 Resets and restart TETC Configuration
Function ResetTETC {
    $outputBox.clear()

    #Question if you want to start do this action
    $answer1 = $wshell.Popup("NOTE: This is an option for Windows Control!`r`nThis will close TETC, remove all calibration profiles and saved screenplanes to reset it to a clean state.`r`nContinue?",0,"Caution",48+4)
    if($answer1 -eq 6){$outputBox.AppendText( "Starting...`r`n" )

    $Processkill = get-process "Tobii.Service",
    "Tobii.EyeX Controller Core",
    "Tobii.EyeX.Engine",
    "Tobii.EyeX.Interaction",
    "Tobii.EyeX.Tray" | Stop-process -force -Passthru -erroraction ignore | Select Processname |
    Format-table -Hidetableheaders | Out-string
    $Outputbox.Appendtext( "Stopping processes:`r`n" )
    $Outputbox.Appendtext($Processkill)
    }
    elseif($answer1 -ne 6){$Outputbox.Appendtext( "Action canceled: Reset TETC.`r`n" )
    return
    }

    $outputBox.AppendText( "Attempting to delete TETC configuration files... `r`n" )
    $keys = ( "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\" )
    $Keys2 = ( "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\*" )

    Foreach($Key in $Keys)
    {
    if (test-path $keys)
    {
    $outputBox.appendtext( "Configuration files found! - Removing...`r`n" )
    Remove-itemProperty $Keys -Name "DefaultEyeTracker" -ErrorAction Ignore
    Remove-item $Keys2 -Recurse -Force -ErrorAction Ignore
    }
    else
    {
    $outputBox.Appendtext("No TETC configuration files were found!`r`n")
    }
    }

    try {
    $Outputbox.Appendtext( "Attempting to start Tobii Service...`r`n" )
    Start-Service -Name "Tobii Service" -ErrorAction Stop
    $Outputbox.Appendtext( "Done!`r`n")
    }
    Catch {
    $Outputbox.Appendtext( "Tobii Service failed to start!`r`n" )
    }

    try {
    $OutputBox.AppendText( "Attempting to start TETC...`r`n" )
    Start-process "C:\Program Files (x86)\Tobii\Tobii EyeX Interaction\Tobii.EyeX.Tray.exe" -ErrorAction Stop
    }
    Catch {
    $outputBox.AppendText( "TETC failed to start!`r`n" )
    }
    $Outputbox.Appendtext( "Finished!`r`n" )
}

#B1 Function listapps - outputs all installed apps with the publisher Tobii
Function Listapps {
    $Outputbox.clear()
    $Outputbox.Appendtext( "Listing installed Tobii software... (If empty, no software found) `r`n" )
    $Listapps = Get-ChildItem -Recurse -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\,
    HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\,
    HKLM:\Software\WOW6432Node\Tobii\ |
    Get-ItemProperty | Where-Object {$_.Publisher -like '*Tobii*'} | Select Displayname, Displayversion | format-table -HideTableHeaders | out-string
    $Listwindowsapp = Get-AppxPackage | Where-Object {($_.Publisher -like '*Tobii*') -or
    ($_.Name -like '*Snap*')} | Select name | format-table -HideTableHeaders | out-string
    $outputBox.AppendText( "TOBII INSTALLED SOFTWARE:`n" )
    $Outputbox.Appendtext($Listapps)
    $outputBox.AppendText( "TOBII WINDOWS STORE APPS:`n" )
    $Outputbox.Appendtext($Listwindowsapp)
}

#B2 Lists currently active tobii processes & services
Function GetProcesses {
    $outputBox.clear()
    $outputBox.appendtext( "Listing active Tobii processes. (If empty - no processes were found) `r`n" )
    $GetProcess = get-process "GazeSelection", "*Tobii*" | Select Processname | Format-table -hidetableheaders | Out-string
    $GetServices = Get-Service -Name '*Tobii*'| Select Name, Status | Format-table -hidetableheaders | Out-string
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("ACTIVE PROCESSES:")
    $outputBox.Appendtext($GetProcess)
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("ACTIVE Services:")
    $outputBox.Appendtext("$GetServices `r`n")
}

#B3 Stops all currently active tobii processes
Function RestartProcesses {
    $outputBox.clear()

    #Question if you want to start do this action
    $answer1 = $wshell.Popup("This will close any tobii software that is currently open, please ensure that you have saved any changes in your applications.`r`nContinue?",0,"Caution",48+4)
    if($answer1 -eq 6){
    $Outputbox.Appendtext( "Stopping Tobii Service... `r`n")
    $StopServices = Get-Service -Name '*Tobii*'| Stop-Service -force -Passthru -erroraction ignore | Select Name, Status | Format-table -hidetableheaders | Out-string
    $Outputbox.Appendtext($StopServices)
    Start-Sleep -s 3
    $Outputbox.Appendtext( "Stopping the following processes shown below... `r`n")
    $Outputbox.Appendtext( "PROCESSES:" )
    $Processkill = get-process "GazeSelection" , "*TobiiDynavox*", "*Tobii.EyeX*", "Notifier" | Stop-process -force -Passthru -erroraction ignore | Select Processname | Format-table -Hidetableheaders | Out-string
    $outputBox.Appendtext($Processkill)
    }
    elseif($answer1 -ne 6){
        $Outputbox.Appendtext( "Action canceled: Restart Tobii Services.`r`n" )
        return
    }
    #start all processes and services
    $Outputbox.Appendtext( "Attempting to start Tobii Service... `r`n" )
    Start-Sleep -s 3
    try {
    $StartServices = Start-Service -Name '*Tobii*' -ErrorAction Stop

    $Outputbox.Appendtext( "$StartServicesr`n")
    }
    Catch {
    $Outputbox.Appendtext( "Tobii Service failed to start!`r`n" )
    }
    $Outputbox.Appendtext( "Attempting to start Eyeassist... `r`n" )
    Start-Sleep -s 3
    try {
    $StartProcesses = Start-process "C:\Program Files (x86)\Tobii Dynavox\Eye Assist\TobiiDynavox.EyeAssist.Engine.exe"
    $Outputbox.Appendtext( "$StartProcesses`r`n" )
    }
    Catch {
    $outputBox.Appendtext( "EyeAssist failed to start!`r`n" )
    }
    $outputBox.Appendtext( "Done!`r`n" )
    }

#B4
Function DeleteServices {
    $outputBox.clear()
    $outputBox.appendtext( "Deleting All Tobii IS5 Services.`r`n" )
    $DeleteServices = Get-Service -Name '*TobiiIS*' , '*TobiiG*' | Stop-Service -Force -passthru -ErrorAction ignore
    foreach ($deleteService in $DeleteServices) {
        $outputbox.appendtext($DeleteService)        sc.exe delete $DeleteService    }
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done! `r`n")
}

#B5
Function IS5PID {
$outputBox.clear()
$outputBox.appendtext( "Checking IS5 PID...`r`n" )
$test = $null
$test = Get-CimInstance Win32_PnPSignedDriver | where Description -Like "*WinUSB Device*" | select DeviceID

if (!$test) {
$outputbox.appendtext("the tracker is not connected")
} else {

$outputbox.appendtext($test)
}
$outputbox.appendtext("`r`n")
}

#B6
Function InstallPDK {
$outputBox.clear()
$outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
Start-Process -FilePath C:\Users\aes\Desktop\script\BeforeUninstall.bat
$outputbox.appendtext("`r`n")
$outputbox.appendtext("ACTIVE PROCESSES: `r`n")
}

#B7
Function ListDrivers {
$outputBox.clear()
$outputBox.appendtext( "listing all drivers in c:/tobii.txt.`r`n" )
pnputil /enum-drivers >c:\tobii.txt
Get-WindowsDriver -Online -All | Where-Object {$_.ProviderName -eq "Tobii AB"} |
ForEach-Object {
$outputbox.appendtext( $_.Driver + "`r`n");
$outputbox.appendtext($_.OriginalFileName + "`r`n");
$outputbox.appendtext($_.DeviceID + "`r`n");
$outputbox.appendtext("`r`n")
}
$outputbox.appendtext("`r`n")
$outputbox.appendtext("Done `r`n")
}

#B8
Function RemoveDrivers {
$outputBox.appendtext( "removing Tobii Drivers...`r`n" )
$TobiiVer = Get-WindowsDriver -Online -All | Where-Object {$_.ProviderName -eq "Tobii AB"} | Select-Object Driver
ForEach ($ver in $TobiiVer) {
$driver = $ver.Driver
pnputil /delete-driver $driver /force /uninstall
}
$outputbox.appendtext("`r`n")
$outputbox.appendtext("Done `r`n")
}

#B9
Function resetBOOT {
IS5PID
$outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
cd 'C:\Users\aes\Documents\Ammar\Tools backup\Windows\Windows\'
.\CastorUsbCli.exe --reset boot
$outputbox.appendtext("`r`n")
$outputbox.appendtext("Done `r`n")
IS5PID
}

#B10
Function ETfw {
$outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
$outputbox.appendtext("`r`n")
$outputbox.appendtext("Done `r`n")
}

#B11
Function FWUpgrade {
$outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
$outputbox.appendtext("`r`n")
$outputbox.appendtext("Done `r`n")
}

#B12
Function BeforeUninstallGG {
$outputBox.clear()
$outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )



$Outputbox.appendtext( "Done! `r`n" )
$outputbox.appendtext("`r`n")
$outputbox.appendtext("ACTIVE PROCESSES: `r`n")
}

#B13
Function ETConnection {
$outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
$outputbox.appendtext("`r`n")
$outputbox.appendtext("Done `r`n")
}

#B14
Function HBTool {
$outputBox.clear()
$outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
Start-Process -FilePath C:\Users\aes\Desktop\script\BeforeUninstall.bat
$outputbox.appendtext("`r`n")
$outputbox.appendtext("ACTIVE PROCESSES: `r`n")
}

#B15
Function BackupGazeInteraction {

$path = ( "C:\ProgramData\Tobii Dynavox\Old Gaze Interaction" )

$outputbox.Appendtext( "Attempting to backup folder...`r`n" )
if (Test-path $path)
{
$outputBox.appendtext( "Backup folder already exist in: C:\ProgramData\Tobii Dynavox\Old Gaze Interaction, please move it to another location or remove it before trying to backup again." )
}
else {
try {
Copy-item "C:\ProgramData\Tobii Dynavox\Gaze Interaction\" "C:\ProgramData\Tobii Dynavox\Old Gaze Interaction\" -recurse -Erroraction Stop
$outputBox.appendtext( "Copying Gaze Interaction folder to 'Old Gaze Interaction' and placing it in C:\ProgramData\Tobii Dynavox\`r`n" )
$outputBox.appendtext( "Finished!`r`n" )
}
Catch {
$outputBox.appendtext( "Failed - No Gaze Interaction folder could be found!`r`n" )
}
}
}

#B16 Copy licenses function. If any path to $Licensepaths exists, it will make a folder "Tobii Licenses", copy the licensefolders to the new folder(Does not contain the keys.xml, it is only the folder)
Function Copylicenses {

$licensepaths = ( "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control",
"C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction",
"C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5",
"C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4",
"C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Viewer" )

$outputBox.appendtext( "Looking for licenses to copy...`r`n" )
ForEach ($Path in $licensepaths) {
if (test-path $path) {
md "C:\Tobii Licenses" -erroraction ignore
copy-item $path "C:\Tobii Licenses" -erroraction ignore
$outputBox.appendtext( "" )
}
elseif((test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\*") -eq $False) {

$outputBox.appendtext( "No licenses found.`r`n" )
Return
}
}
$outputBox.AppendText( "Copying licenses to C:\Tobii Licenses...`r`n" )

#Retrieves the content from keys.xml
if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control\*") {
$GetcontentWC = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control\keys.xml"
$Outputbox.appendtext( "-- Window Control license copied.`r`n" )
}

if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction\*") {
$GetcontentTGIS = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction\keys.xml"
$Outputbox.appendtext( "-- Gaze Interaction license copied.`r`n" )
}

if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5\*") {
$GetcontentTC5 = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5\keys.xml"
$Outputbox.appendtext( "-- Communicator 5 license copied.`r`n" )
}

if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4\*") {
$GetcontentTC4 = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4\keys.xml"
$Outputbox.appendtext( "-- Communicator 4 license copied.`r`n" )
} #Add compass to the list.

#Filters the content to only get the string between the activationkey words
$LicenseWC = [regex]::Matches($getcontentWC,'(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)',"singleline").Value.trim()
$LicenseTGIS = [regex]::Matches($getcontentTGIS,'(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)',"singleline").Value.trim()
$LicenseTC5 = [regex]::Matches($getcontentTC5,'(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)',"singleline").Value.trim()
$LicenseTC4 = [regex]::Matches($getcontentTC4,'(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)',"singleline").Value.trim()

#Creates txt files for licenses
$LicenseWC | Out-file "C:\Tobii Licenses\Windows Control\Windows Control License.txt" -erroraction ignore
$LicenseTGIS | Out-file "C:\Tobii Licenses\Gaze Interaction\Gaze Interaction License.txt" -erroraction ignore
$LicenseTC5 | Out-file "C:\Tobii Licenses\Communicator 5\Communicator 5 License.txt" -erroraction ignore
$LicenseTC4 | Out-file "C:\Tobii Licenses\Communicator 4\Communicator 4 License.txt" -erroraction ignore

$outputBox.AppendText( "Done!`r`n" )
Return
}

#B17
Function TimeIssueFinder {
$outputBox.clear()
$outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
Start-Process -FilePath C:\Users\aes\Desktop\script\BeforeUninstall.bat
$outputbox.appendtext("`r`n")
$outputbox.appendtext("ACTIVE PROCESSES: `r`n")
}

#B18
Function BIOECv {
$outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
$outputbox.appendtext("`r`n")
$outputbox.appendtext("Done `r`n")
}

#B19
Function HWvsSW {
$outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
$outputbox.appendtext("`r`n")
$outputbox.appendtext("Done `r`n")
}

#B20
Function Diagnostic {
$outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
$outputbox.appendtext("`r`n")
$outputbox.appendtext("Done `r`n")
}

#B21
Function InternalSE {
$outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
$outputbox.appendtext("`r`n")
$outputbox.appendtext("Done `r`n")
}

#B22
Function ETSamples {
$outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
$outputbox.appendtext("`r`n")
$outputbox.appendtext("Done `r`n")
}


#Windows forms
$Optionlist=@("Remove PCEye5 Bundle","Remove Tobii Device Drivers For Windows","Remove WC&GP Bundle","Remove PCEye Package","Remove Communicator","Remove Compass","Remove TGIS only","Remove TGIS profile calibrations","Remove all users C5", "Reset TETC")
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(600,550)
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False

#Informationtext above the dropdown list.
$DropDownLabel = new-object System.Windows.Forms.Label
$DropDownLabel.Location = new-object System.Drawing.Size(10,10)
$DropDownLabel.size = new-object System.Drawing.Size(160,20)
$DropDownLabel.Text = "Select an option"
$Form.Controls.Add($DropDownLabel)

#Dropdown list with options
$DropDownBox = New-Object System.Windows.Forms.ComboBox
$DropDownBox.Location = New-Object System.Drawing.Size(10,30)
$DropDownBox.Size = New-Object System.Drawing.Size(220,20)
$DropDownBox.DropDownHeight = 230
$Form.Controls.Add($DropDownBox)

#For each arrayitem in optionlist, add it to $dropdownbox items.
foreach ($option in $optionlist) {
$DropDownBox.Items.Add($option)
}

#Outputbox
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Size(10,150)
$outputBox.Size = New-Object System.Drawing.Size(400,340)
$outputBox.MultiLine = $True
$outputBox.ScrollBars = "Vertical"
$Form.Controls.Add($outputBox)
$outputBox.font = New-Object System.Drawing.Font ("Consolas" , 8,[System.Drawing.FontStyle]::Regular)

#Button "Start"
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(10,60)
$Button.Size = New-Object System.Drawing.Size(110,50)
$Button.Text = "Start"
$Button.Font = New-Object System.Drawing.Font ("" , 12,[System.Drawing.FontStyle]::Regular)
$Form.Controls.Add($Button)
$Button.Add_Click{selectedscript}

#B1 Button "List Tobii Software"
$Button1 = New-Object System.Windows.Forms.Button
$Button1.Location = New-Object System.Drawing.Size(250,0)
$Button1.Size = New-Object System.Drawing.Size(160,30)
$Button1.Text = "List Tobii Software"
$Button1.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button1)
$Button1.Add_Click{ListApps}

#B2 Button "List active Tobii processes"
$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Size(250,30)
$Button3.Size = New-Object System.Drawing.Size(160,30)
$Button3.Text = "Active Process Service"
$Button3.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button3)
$Button3.Add_Click{GetProcesses}

#B3 Restart Services
$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Size(250,60)
$Button3.Size = New-Object System.Drawing.Size(160,30)
$Button3.Text = "Restart Services"
$Button3.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button3)
$Button3.Add_Click{RestartProcesses}

#B4 Button "Delete services"
$Button10 = New-Object System.Windows.Forms.Button
$Button10.Location = New-Object System.Drawing.Size(250, 90)
$Button10.Size = New-Object System.Drawing.Size(160,30)
$Button10.Text = "-Delete services"
$Button10.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button10)
$Button10.Add_Click{DeleteServices}

#B5 Button "Check IS5 PID"
$Button5 = New-Object System.Windows.Forms.Button
$Button5.Location = New-Object System.Drawing.Size(250, 120)
$Button5.Size = New-Object System.Drawing.Size(160,30)
$Button5.Text = "-List IS5 PID"
$Button5.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button5)
$Button5.Add_Click{IS5PID}

#B6 Button "Install PDK"
$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(410,0)
$Button2.Size = New-Object System.Drawing.Size(160,30)
$Button2.Text = "-Install PDK"
$Button2.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button2)
$Button2.Add_Click{InstallPDK}

#B7 Button "List Tobii Drivers"
$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(410, 30)
$Button2.Size = New-Object System.Drawing.Size(160,30)
$Button2.Text = "-List Tobii Drivers"
$Button2.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button2)
$Button2.Add_Click{ListDriver}

#B8 Button "Remove Drivers"
$Button6 = New-Object System.Windows.Forms.Button
$Button6.Location = New-Object System.Drawing.Size(410,60)
$Button6.Size = New-Object System.Drawing.Size(160,30)
$Button6.Text = "-Remove Drivers"
$Button6.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button6)
$Button6.Add_Click{RemoveDrivers}

#B9 Button "Reset IS5 to bootloader"
$Button7 = New-Object System.Windows.Forms.Button
$Button7.Location = New-Object System.Drawing.Size(410, 90)
$Button7.Size = New-Object System.Drawing.Size(160,30)
$Button7.Text = "Reset ET BOOT"
$Button7.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button7)
$Button7.Add_Click{resetBOOT}

#B10 Button "ET fw"
$Button7 = New-Object System.Windows.Forms.Button
$Button7.Location = New-Object System.Drawing.Size(410, 120)
$Button7.Size = New-Object System.Drawing.Size(160,30)
$Button7.Text = "-ET fw"
$Button7.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button7)
$Button7.Add_Click{ETfw}

#B11 Button "FW Upgrade"
$Button6 = New-Object System.Windows.Forms.Button
$Button6.Location = New-Object System.Drawing.Size(410,150)
$Button6.Size = New-Object System.Drawing.Size(160,30)
$Button6.Text = "-FW Upgrade"
$Button6.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button6)
$Button6.Add_Click{FWUpgrade}

#B12 Button "Before Uninstall GG"
$Button6 = New-Object System.Windows.Forms.Button
$Button6.Location = New-Object System.Drawing.Size(410,180)
$Button6.Size = New-Object System.Drawing.Size(160,30)
$Button6.Text = "-Before Uninstall GG"
$Button6.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button6)
$Button6.Add_Click{BeforeUninstallGG}

#B13 Button "Check ET connection through Service"
$Button9 = New-Object System.Windows.Forms.Button
$Button9.Location = New-Object System.Drawing.Size(410, 210)
$Button9.Size = New-Object System.Drawing.Size(160,30)
$Button9.Text = "-Check ET connection"
$Button9.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button9)
$Button9.Add_Click{}

#B14 Button "Run GG HB_tool"
$Button12 = New-Object System.Windows.Forms.Button
$Button12.Location = New-Object System.Drawing.Size(410, 240)
$Button12.Size = New-Object System.Drawing.Size(160,30)
$Button12.Text = "-Run HB tool"
$Button12.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button12)
$Button12.Add_Click{HBTool}

#B15 Button "Backup Gaze Interaction"
$Button12 = New-Object System.Windows.Forms.Button
$Button12.Location = New-Object System.Drawing.Size(410, 270)
$Button12.Size = New-Object System.Drawing.Size(160,30)
$Button12.Text = "-Backup Gaze Int"
$Button12.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button12)
$Button12.Add_Click{BackupGazeInteraction}

#B16 Button "Copy licenses"
$Button8 = New-Object System.Windows.Forms.Button
$Button8.Location = New-Object System.Drawing.Size(410, 300)
$Button8.Size = New-Object System.Drawing.Size(160,30)
$Button8.Text = "-Copy licenses"
$Button8.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button8)
$Button8.Add_Click{Copylicenses}

#B17 Button "I-Series time issue finder"
$Button13 = New-Object System.Windows.Forms.Button
$Button13.Location = New-Object System.Drawing.Size(410, 330)
$Button13.Size = New-Object System.Drawing.Size(160,30)
$Button13.Text = "-I-12 time issue finder"
$Button13.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button13)
$Button13.Add_Click{TimeIssueFinder}

#B18 Button "BIOECv"
$Button8 = New-Object System.Windows.Forms.Button
$Button8.Location = New-Object System.Drawing.Size(410, 360)
$Button8.Size = New-Object System.Drawing.Size(160,30)
$Button8.Text = "-BIO EC v"
$Button8.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button8)
$Button8.Add_Click{BIOECv}

#B19 Button "HWvsSW"
$Button8 = New-Object System.Windows.Forms.Button
$Button8.Location = New-Object System.Drawing.Size(410, 390)
$Button8.Size = New-Object System.Drawing.Size(160,30)
$Button8.Text = "-HW vs SW"
$Button8.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button8)
$Button8.Add_Click{HWvsSW}

#B20 Button "Diagnostic"
$Button14 = New-Object System.Windows.Forms.Button
$Button14.Location = New-Object System.Drawing.Size(410, 420)
$Button14.Size = New-Object System.Drawing.Size(160,30)
$Button14.Text = "-Diagnostic"
$Button14.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button14)
$Button14.Add_Click{Diagnostic}

#B21 Button "InternalSE"
$Button15 = New-Object System.Windows.Forms.Button
$Button15.Location = New-Object System.Drawing.Size(410, 450)
$Button15.Size = New-Object System.Drawing.Size(160,30)
$Button15.Text = "-Internal SE"
$Button15.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button15)
$Button15.Add_Click{InternalSE}

#B22 Button "ETSamples"
$Button16 = New-Object System.Windows.Forms.Button
$Button16.Location = New-Object System.Drawing.Size(410, 480)
$Button16.Size = New-Object System.Drawing.Size(160,30)
$Button16.Text = "-ET Samples"
$Button16.Font = New-Object System.Drawing.Font ("" , 8,[System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button16)
$Button16.Add_Click{ETSamples}

#Form name + activate form.
$Form.Text = "Support Tool 1.1"
$Form.Add_Shown({$Form.Activate()})
$Form.ShowDialog()