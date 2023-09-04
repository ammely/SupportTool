#v1.4 DummyFix
#Forces powershell to run as an admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{ Start-Process powershell.exe "-NoProfile -Windowstyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

#Imports Windowsforms and Drawing from system
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#Allows the use of wshell for confirmation popups
$wshell = New-Object -ComObject Wscript.Shell
$PSScriptRoot
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
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
    elseif ($DropDownBox.Selecteditem -eq "Backup Gaze Interaction") {
        $Outputbox.Clear()
        BackupGazeInteraction
    }
    elseif ($DropDownBox.Selecteditem -eq "Copy License") {
        $Outputbox.Clear()
        Copylicenses
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
    $answer1 = $wshell.Popup("This will remove all software included in PCeye5Bundle.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove PCEye5Bundle`r`n" )
        Return
    }
	
	$Outputbox.Appendtext("Backup Calibration profiles. To restore profiles, go to %temp% and double click on Eula & EyeXConfig`r`n" )
    $RegPath1 = 'HKEY_USERS\S-1-5-21-2271707575-1334560000-3059665169-12978\SOFTWARE\Tobii\EULA'
    $RegPath2 = 'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Tobii\EyeXConfig'
    if (Test-Path -Path $RegPath1) {
    	Invoke-Command  {reg export $RegPath1 "$ENV:USERPROFILE\AppData\Local\Temp\EULA.reg" }
    }
    if (Test-Path -Path $RegPath1) {
    	Invoke-Command  {reg export $RegPath2 "$ENV:USERPROFILE\AppData\Local\Temp\EyeXConfig.reg" }
    }

	$GetProcess = stop-process -Name "*TobiiDynavox*"
    $Outputbox.appendtext("Stopping $GetProcess `r`n" )
    $Outputbox.Appendtext("Please wait.`r`n" )
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { 
        ($_.Displayname -Match "Tobii Dynavox Computer Control") -or
        ($_.Displayname -Match "Dynavox Computer Control Updater Service") -or
        ($_.Displayname -Match "Tobii Dynavox Update Notifier") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking") -or
        ($_.Displayname -Eq "Tobii Device Drivers For Windows (PCEye5)") -or
        ($_.Displayname -Eq "Tobii Experience Software For Windows (PCEye5)") } | Select-Object Displayname, UninstallString
    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString -replace "msiexec.exe", "" -Replace "/I", "" -Replace "/X", ""
        $uninst = $uninst.Trim()
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
    }

    $outputBox.appendtext( "Deleting All Tobii IS5 Services.`r`n" )
    $DeleteServices = Get-Service -Name '*TobiiIS*' , '*TobiiG*' | Stop-Service -Force -passthru -ErrorAction ignore
    foreach ($Service in $DeleteServices) {
        $outputbox.appendtext($Service)
        sc.exe delete $Service
    }

    $outputBox.appendtext( "Removing Tobii Drivers...`r`n" )
    $TobiiVer = Get-WindowsDriver -Online -All | Where-Object { $_.ProviderName -eq "Tobii AB" } | Select-Object Driver
    ForEach ($ver in $TobiiVer) {
        pnputil /delete-driver $ver.Driver /force /uninstall
    }

    $Outputbox.appendtext( "Looking for related folders...`r`n" )
    #Removes WC related folders
    $paths = (
        "C:\Program Files (x86)\Tobii Dynavox\Eye Tracking Settings",	
        "C:\Program Files (x86)\Tobii Dynavox\Eye Assist",
        "C:\Program Files\Tobii\Tobii EyeX",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\App Switcher",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control Bundle",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Update Notifier\",
        "$ENV:ProgramData\Tobii Dynavox\EyeAssist",
        "$ENV:ProgramData\Tobii Dynavox\Computer Control",
        "$ENV:ProgramData\HelloDMFT" )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }
    $Keys = (
        "HKCU:\Software\Tobii\EyeAssist",
        "HKCU:\Software\Tobii\Update Notifier",
        "HKCU:\Software\Tobii Dynavox\Computer Control",
        "HKLM:\SOFTWARE\WOW6432Node\Tobii Dynavox\Computer Control Updater Service" )

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }

    $Outputbox.Appendtext( "Finished!`r`n" )
}

#A2 Uninstalls ALL Tobii Device Drivers For Windows Bundle
Function UninstallTobiiDeviceDriversForWindows {

    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove all software included in Tobii Device Drivers For Windows Bundles.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove Tobii Device Drivers For Windows`r`n" )
        Return
    }
    $Outputbox.Appendtext( "Please wait.`r`n" )

   	$Outputbox.Appendtext("Backup Calibration profiles. To restore profiles, go to %temp% and double click on Eula & EyeXConfig`r`n" )
    $RegPath1 = 'HKEY_USERS\S-1-5-21-2271707575-1334560000-3059665169-12978\SOFTWARE\Tobii\EULA'
    $RegPath2 = 'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Tobii\EyeXConfig'
    if (Test-Path -Path $RegPath1) {
    	Invoke-Command  {reg export $RegPath1 "$ENV:USERPROFILE\AppData\Local\Temp\EULA.reg" }
    }
    if (Test-Path -Path $RegPath1) {
    	Invoke-Command  {reg export $RegPath2 "$ENV:USERPROFILE\AppData\Local\Temp\EyeXConfig.reg" }
    }

    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
	try{ 
        $erroractionpreference = "Stop"
		$Firmware = .\FWUpgrade32.exe --auto --info-only 
    }
    catch [System.Management.Automation.RemoteException] {
		$outputbox.appendtext("No Eye Tracker Connected`r`n")
    }
    if ($Firmware -match "IS5_Gibbon_Gaze") { 
        $outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "BeforeUninstall.bat" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        Set-Location $fpath
        $Installer = cmd /c "BeforeUninstall.bat"
        $Outputbox.appendtext($Installer)
        $outputbox.appendtext("`r`n")
    } 

    $GetProcess = stop-process -Name "*TobiiDynavox*" -Force
    $Outputbox.appendtext("Stopping $GetProcess `r`n" )
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Device Drivers For Windows") -or
        ($_.Displayname -Match "Tobii Experience Software") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking Driver") -or
        ($_.Displayname -Match "Tobii Eye Tracking For Windows") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking") } | Select-Object Displayname, UninstallString
    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        $uninst = $ver.UninstallString -replace "msiexec.exe", "" -Replace "/I", "" -Replace "/X", "" -replace "/uninstall", ""
        $uninst = $uninst.Trim()
        if ($uninst -match "ProgramData") {
            cmd /c $uninst /uninstall /quiet
        }
        else {
            start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
        }
    }

    $outputBox.appendtext( "Deleting All Tobii IS5 Services.`r`n" )
    $DeleteServices = Get-Service -Name '*TobiiIS*' , '*TobiiG*' | Stop-Service -Force -passthru -ErrorAction ignore
    foreach ($Service in $DeleteServices) {
        $outputbox.appendtext($Service)
        sc.exe delete $Service
    }

    $outputBox.appendtext( "Removing Tobii Drivers...`r`n" )
    $TobiiVer = Get-WindowsDriver -Online -All | Where-Object { $_.ProviderName -eq "Tobii AB" } | Select-Object Driver
    ForEach ($ver in $TobiiVer) {
        pnputil /delete-driver $ver.Driver /force /uninstall
    }
    $outputbox.appendtext("`r`n")
    $Outputbox.appendtext( "Done!`r`n" )
}

#A3 Uninstalls WC Bundle
Function UninstallWCGP {

    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove all software included in Windows Control & Gaze Point Bundles.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove WC&GP`r`n" )
        Return
    }

    #If second answer equals yes or no
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?", 0, "Caution", 48 + 4)
    if ($answer2 -eq 6) { CopyLicenses }

    elseif ($answer2 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" )
    }

	$Outputbox.Appendtext("Backup Calibration profiles. To restore profiles, go to %temp% and double click on Eula & EyeXConfig`r`n" )
    $RegPath1 = 'HKEY_USERS\S-1-5-21-2271707575-1334560000-3059665169-12978\SOFTWARE\Tobii\EULA'
    $RegPath2 = 'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Tobii\EyeXConfig'
    if (Test-Path -Path $RegPath1) {
    	Invoke-Command  {reg export $RegPath1 "$ENV:USERPROFILE\AppData\Local\Temp\EULA.reg" }
    }
    if (Test-Path -Path $RegPath1) {
    	Invoke-Command  {reg export $RegPath2 "$ENV:USERPROFILE\AppData\Local\Temp\EyeXConfig.reg" }
    }
    $Outputbox.Appendtext( "Please wait.`r`n" )
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Windows Control") -or
        ($_.Displayname -Match "Virtual Remote") -or
        ($_.Displayname -Match "Update Notifier") -or
        ($_.Displayname -Match "Tobii Eye Tracking") -or
        ($_.Displayname -Match "GazeSelection") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze Point") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze Point Configuration Guide") } | Select-Object Displayname, UninstallString
    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        & cmd /c $uninst /quiet /norestart
    }

    $Outputbox.appendtext( "Looking for related folders...`r`n" )
    #Removes WC related folders
    $paths = ( 
        "$Env:USERPROFILE\AppData\Roaming\Tobii\Tobii Interaction\",
        "$Env:USERPROFILE\AppData\Roaming\Tobii\Tobii Interaction Statistics\",
        "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist",
        "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Gaze Selection",
        "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Windows Control Bundle",
        "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Gaze Point Bundle",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Update Notifier\",
        "$Env:USERPROFILE\AppData\Local\Tobii\Tobii Interaction\",
        "C:\Program Files (x86)\Tobii Dynavox\Windows Control Configuration Guide",
        "C:\Program Files (x86)\Tobii Dynavox\Gaze Point Configuration Guide",
        "C:\Program Files (x86)\Tobii Dynavox\Update Notifier",
        "C:\Program Files (x86)\Tobii\Service\Plugins",
        "$ENV:ProgramData\Tobii Dynavox\Tobii Interaction\ScreenPlanes\",
        "$ENV:ProgramData\Tobii Dynavox\Update Notifier\",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control\",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Point\",
        "$ENV:ProgramData\Tobii Dynavox\Windows Control Configuration Guide\",
        "$ENV:ProgramData\Tobii\Statistics\",
        "$ENV:ProgramData\Tobii\Tobii Interaction\",
        "$ENV:ProgramData\Tobii\Tobii Stream Engine\",
        "$ENV:ProgramData\TetServer" )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }
	
    #Deleting registry keys related to WC
    $Outputbox.appendtext( "Looking for related Registry keys...`r`n" )
    $Keys = ( 
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeX",
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\",
        "HKLM:\SOFTWARE\Wow6432Node\Tobii\TobiiUpdater\",
        "HKLM:\SOFTWARE\Wow6432Node\Tobii\Update Notifier\",
        "HKLM:\SOFTWARE\Wow6432Node\Tobii\EyeXOverview",
        "HKCU:\Software\Tobii\ExternalNotifications",
        "HKCU:\Software\Tobii\Eye Control Suite",
        "HKCU:\Software\Tobii\EyeX",
        "HKCU:\Software\Tobii\Statistics",
        "HKCU:\Software\Tobii\Vouchers"
    )

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }

    $Outputbox.Appendtext( "Finished!`r`n" )
}

#A4
Function UninstallPCEyePackage {
    #Implement functionality. (PCEye package & TGIS on i-series, start with PCEye package
    $answer1 = $wshell.Popup("This will remove all software included in PCEye Package`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove PCEye Package`r`n" )
        Return
    }

    #If second answer equals yes or no
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?", 0, "Caution", 48 + 4)
    if ($answer2 -eq 6) { CopyLicenses }

    elseif ($answer2 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" )
    }

    $Outputbox.Appendtext( "Please wait.`r`n" )

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Dynavox Gaze Interaction Software") -or
        ($_.Displayname -Match "Tobii Dynavox PCEye Update Notifier") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze Selection Language Packs") -or
        ($_.Displayname -Match "Tobii IS3 Eye Tracker Driver") -or
        ($_.Displayname -Match "Tobii IS4 Eye Tracker Driver") -or
        ($_.Displayname -Match "Tobii Eye Tracker Browser") -or
        ($_.Displayname -Match "Tobii Dynavox PCEye Configuration Guide") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze HID") } | Select-Object Displayname, UninstallString

    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        & cmd /c $uninst /quiet /norestart
    }

    $UninstallService = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -match "Tobii Service" }

    ForEach ($Software in $UninstallService) {
        $Uninstname2 = $Software.Name
        $Outputbox.Appendtext( "Removing - " + "$Uninstname2`r`n")
        $Software.Uninstall()
    }

    $paths = ( 
        "$ENV:AppData\Tobii Dynavox\PCEye Configuration Guide",
        "$ENV:AppData\Tobii Dynavox\PCEye Update Notifier\",
        "$ENV:ProgramData\Tobii Dynavox\PCEye Configuration Guide",
        "$ENV:ProgramData\Tobii Dynavox\Gaze Interaction\Server",
        "$ENV:ProgramData\Tobii Dynavox\PCEye Update Notifier",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction",
        "$ENV:ProgramData\Tobii Dynavox\Tobii Interaction",
        "$ENV:ProgramData\Tobii Dynavox\Gaze Selection",
        "$ENV:ProgramData\Tobii\Statistics\",
        "$ENV:ProgramData\Tobii\Tobii Interaction",
        "$ENV:ProgramData\Tobii\Tobii Stream Engine\odin",
        "$ENV:ProgramData\TetServer"
    )

    $Outputbox.Appendtext( "Looking for folders to remove...`r`n" )
    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }

    $Key = (
		"HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation",
		"HKCU:\SOFTWARE\Tobii\PCEye\Update Notifier",
		"HKCU:\SOFTWARE\Tobii\PCEye")
		
		
    if (test-path $Key) {
        $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
        Remove-Item $Key -Force -ErrorAction ignore
    }

    $Outputbox.appendtext( "Finished!`r`n" )
}

#A5 Uninstalls all Tobii related software
Function UninstallCommunicator {
    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will uninstall Communicator. Are you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove Communicator`r`n" )
        Return
    }

    #If second answer equals yes or no - if "Yes" then it will call the function CopyLicenses and then continue.
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?", 0, "Caution", 48 + 4)
    if ($answer2 -eq 6) { CopyLicenses }

    elseif ($answer2 -ne 6) { $Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" ) }

    $Outputbox.Appendtext( "Please wait.`r`n" )

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { $_.Displayname -match "Tobii Dynavox Communicator" } | Select-Object Publisher, Displayname, UninstallString

    $Outputbox.appendtext( "Starting uninstallation...`r`n" )

    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        $uninst = $ver.UninstallString
        & cmd /c $uninst /quiet /norestart
    }
    $Outputbox.appendtext( "Done!`r`n" )

    $Outputbox.appendtext( "Looking for Communicator keys & folders...`r`n" )

    $paths = ( "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Communicator",
        "$ENV:ProgramData\Tobii Dynavox\Communicator" )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.AppendText( "Removing - " + "$path`r`n")
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }

    $Keys = ("HKLM:\SOFTWARE\WOW6432Node\Tobii\MyTobii\MPA\VS Communicator 4",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5" )

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }

    $Outputbox.Appendtext( "Finished!`r`n" )
}

#A6 Uninstalls only Compass
Function UninstallCompass {
    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will uninstall Compass. Are you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove Compass`r`n" )
        Return
    }

    $Outputbox.Appendtext( "Please wait.`r`n" )


    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Dynavox Compass") } | Select-Object Displayname, UninstallString

    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $Outputbox.Appendtext( "Removing - " + "$Uninstname" )
        $uninst = $ver.UninstallString
        & cmd /c $uninst /quiet /norestart
    }

    $answer2 = $wshell.Popup("Do you want to remove related folders?", 0, "Caution", 48 + 4)
    if ($answer2 -eq 6) {
        $Outputbox.appendtext( "Looking for related folders...`r`n" )

        $Keys = ( "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Compass" )

        foreach ($Key in $Keys) {
            if (test-path $Key) {
                $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
                Remove-item $Key -Recurse -ErrorAction Ignore
            }
        }
    }
    elseif ($answer2 -ne 6) {
        $Outputbox.appendtext( "Action canceled: Remove folders`r`n" )
    }

    $Outputbox.appendtext( "Finished!`r`n" )
}

#A7
Function UninstallTGIS {
    $answer1 = $wshell.Popup("This will ONLY remove Tobii Gaze Interaction Software.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove PCEye5Bundle`r`n" )
        Return
    }

    #If second answer equals yes or no
    $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?", 0, "Caution", 48 + 4)
    if ($answer2 -eq 6) { CopyLicenses }
    elseif ($answer2 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" )
    }
    $Outputbox.Appendtext( "Please wait.`r`n" )
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Dynavox Gaze Interaction Software") } | Select-Object Displayname, UninstallString
    $Outputbox.appendtext( "Starting uninstallation...`r`n" )
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        & cmd /c $uninst /quiet /norestart
    }

    $paths = ("$env:ProgramData\Tobii Dynavox\Gaze Interaction\",
        "$ENV:ProgramData\Tobii Dynavox\Gaze Selection\Word Prediction\Language Packs\")

    $Outputbox.Appendtext( "Looking for related folders...`r`n" )
    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }

    $Outputbox.appendtext( "Finished!`r`n" )
}

#A8 Function for the option "Remove TGIS calibration profiles #Tobii service is stopped
Function TGISProfilesremove {

    $answer1 = $wshell.Popup("This will remove ONLY calibrations for every profile, it will NOT remove the actual profiles. The Gaze Interaction software will close and tobii service will restart.`r`nContinue?", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.appendtext( "Shutting down TGIS software...`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $outputBox.appendtext( "Action canceled: Remove calibration profiles." )
    }	
    $Outputbox.appendtext( "Stopping processes:`r`n" )
    $Processkill = get-process "Tobii.Service", "TobiiEyeControlOptions", "TobiiEyeControlServer", "Notifier" | Stop-process -force -Passthru -erroraction ignore | Select Processname |
    Format-table -Hidetableheaders | Out-string
    $Outputbox.Appendtext($Processkill)

    $outputbox.appendtext( "Looking for calibration profiles...`r`n" )
    $paths = ( "$ENV:ProgramData\Tobii Dynavox\Gaze Interaction\Server\Calibration\*" )

    foreach ($path in $paths) {
        if (Test-Path $path) {
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

    $outputbox.appendtext( "Finished!`r`n" )
}

#A9
Function DeleteC5User {
    $outputBox.clear()
    $outputBox.appendtext( "Deleting C5 users.`r`n" )
    $paths = ( 
        "$env:USERPROFILE\Documents\Communicator 5",
        "$env:USERPROFILE\AppData\Local\VirtualStore\Program Files (x86)\Tobii Dynavox\Communicator 5",
        "$env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Communicator",
        "$env:ProgramData\Tobii Dynavox\Communicator")
    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
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
    $answer1 = $wshell.Popup("NOTE: This is an option for Windows Control!`r`nThis will close TETC, remove all calibration profiles and saved screenplanes to reset it to a clean state.`r`nContinue?", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $outputBox.AppendText( "Starting...`r`n" )

    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Reset TETC.`r`n" )
        return
    }

    $Processkill = get-process "Tobii.Service",	"Tobii.EyeX Controller Core", "Tobii.EyeX.Engine", "Tobii.EyeX.Interaction", "Tobii.EyeX.Tray" | Stop-process -force -Passthru -erroraction ignore | Select Processname | Format-table -Hidetableheaders | Out-string
    $Outputbox.Appendtext( "Stopping processes:`r`n" )
    $Outputbox.Appendtext($Processkill)
    $outputBox.AppendText( "Attempting to delete TETC configuration files... `r`n" )
    $keys = ( "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\" )
    $Keys2 = ( "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\*" )

    Foreach ($Key in $Keys) {
        if (test-path $keys) {
            $outputBox.appendtext( "Configuration files found! - Removing...`r`n" )
            Remove-itemProperty $Keys -Name "DefaultEyeTracker" -ErrorAction Ignore
            Remove-item $Keys2 -Recurse -Force -ErrorAction Ignore
        }
        else {
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

#A11
Function BackupGazeInteraction {
    $outputBox.clear()
    $path = ( "C:\ProgramData\Tobii Dynavox\Old Gaze Interaction" )

    $outputbox.Appendtext( "Attempting to backup folder...`r`n" )
    if (Test-path $path) {
        $outputBox.appendtext( "Backup folder already exist in: C:\ProgramData\Tobii Dynavox\Old Gaze Interaction, please move it to another location or remove it before trying to backup again." )
    }
    else {
        try {
            Copy-item "C:\ProgramData\Tobii Dynavox\Gaze Interaction\" "C:\ProgramData\Tobii Dynavox\Old Gaze Interaction\" -Recurse -Erroraction Stop
            $outputBox.appendtext( "Copying Gaze Interaction folder to 'Old Gaze Interaction' and placing it in C:\ProgramData\Tobii Dynavox\`r`n" )
            $outputBox.appendtext( "Finished!`r`n" )
        }
        Catch {
            $outputBox.appendtext( "Failed - No Gaze Interaction folder could be found!`r`n" )
        }
    }
}

#A12 Copy licenses function. If any path to $Licensepaths exists, it will make a folder "Tobii Licenses", copy the licensefolders to the new folder(Does not contain the keys.xml, it is only the folder)
Function Copylicenses {
    $outputBox.clear()
    $licensepaths = ( "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control",
        "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction",
        "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5",
        "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4",
        "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Viewer" )

    $outputBox.appendtext( "Looking for licenses to copy...`r`n" )
    ForEach ($Path in $licensepaths) {
        if (test-path $path) {
            mkdir "C:\Tobii Licenses" -erroraction ignore
            copy-item $path "C:\Tobii Licenses" -erroraction ignore
            $outputBox.appendtext( "" )
        }
        elseif ((test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\*") -eq $False) {
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
    $LicenseWC = [regex]::Matches($getcontentWC, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()
    $LicenseTGIS = [regex]::Matches($getcontentTGIS, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()
    $LicenseTC5 = [regex]::Matches($getcontentTC5, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()
    $LicenseTC4 = [regex]::Matches($getcontentTC4, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()

    #Creates txt files for licenses
    $LicenseWC | Out-file "C:\Tobii Licenses\Windows Control\Windows Control License.txt" -erroraction ignore
    $LicenseTGIS | Out-file "C:\Tobii Licenses\Gaze Interaction\Gaze Interaction License.txt" -erroraction ignore
    $LicenseTC5 | Out-file "C:\Tobii Licenses\Communicator 5\Communicator 5 License.txt" -erroraction ignore
    $LicenseTC4 | Out-file "C:\Tobii Licenses\Communicator 4\Communicator 4 License.txt" -erroraction ignore

    $outputBox.AppendText( "Done!`r`n" )
    Return
}

#B1 Function listapps - outputs all installed apps with the publisher Tobii
Function Listapps {
    $Outputbox.clear()
    $Outputbox.Appendtext( "Listing installed Tobii software... (If empty, no software found) `r`n" )
    $Listapps = Get-ChildItem -Recurse -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\,
    HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\,
    HKLM:\Software\WOW6432Node\Tobii\ |
    Get-ItemProperty | Where-Object { $_.Publisher -like '*Tobii*' } | Select Displayname, Displayversion | format-table -HideTableHeaders | out-string
    $Listwindowsapp = Get-AppxPackage | Where-Object { ($_.Publisher -like '*Tobii*') -or
        ($_.Name -like '*Snap*') } | Select name | format-table -HideTableHeaders | out-string
    $outputBox.AppendText( "TOBII INSTALLED SOFTWARE:$Listapps`n" )
    $outputBox.AppendText( "TOBII WINDOWS STORE APPS:$Listwindowsapp" )
}

#B2 Lists currently active tobii processes & services
Function GetProcesses {
    $outputBox.clear()
    $outputBox.appendtext( "Listing active Tobii processes. (If empty - no processes were found) `r`n" )
    $GetProcess = get-process "*GazeSelection*", "*Tobii*" | Select Processname | Format-table -hidetableheaders | Out-string
    $GetServices = Get-Service -Name '*Tobii*' | Select Name, Status | Format-table -hidetableheaders | Out-string
    if ($GetProcess) {
        $outputbox.appendtext("ACTIVE PROCESSES:$GetProcess")
        $Outputbox.Appendtext( "`r`n" )
    }
    else {
        $outputbox.appendtext("NO ACTIVE PROCESSES")
        $Outputbox.Appendtext( "`r`n" )
    }
    if ($GetServices) {
        $outputbox.appendtext("ACTIVE Services:$GetServices")
        $Outputbox.Appendtext( "`r`n" )
    }
    else {
        $outputbox.appendtext("NO ACTIVE Services")
        $Outputbox.Appendtext( "`r`n" )
    }
}

#B3
Function IS5PID {
    $outputBox.clear()
    $outputBox.appendtext( "Checking IS5 PID...`r`n" )
    $getdeviceid = $null
    $getdeviceid = gwmi Win32_USBControllerDevice | % { [wmi]($_.Dependent) } | Where-Object DeviceID -Like "*Tobii*" | Select-object DeviceID
	$outputbox.appendtext("$getdeviceid `r`n")								   
    $getdeviceid2 = Get-CimInstance Win32_PnPSignedDriver | Where-Object Description -Like "*WinUSB Device*" | Select-Object DeviceID
    Start-Sleep -s 5
	$outputbox.appendtext("$getdeviceid2`r`n")								   
    if (!$getdeviceid -or !$getdeviceid2) {
        $outputbox.appendtext("the tracker is not connected")
    }
    # gwmi Win32_USBControllerDevice |%{[wmi]($_.Dependent)} | Sort Manufacturer,Description,DeviceID | Ft -GroupBy Manufacturer Description,Service,DeviceID | out-file c:\VidPid.txt
	$outputbox.appendtext("Done`r`n")
    $outputbox.appendtext("`r`n")
}

#B4
Function ListDrivers {
    $outputBox.clear()
    $outputBox.appendtext( "Listing all drivers in c:/tobii.txt and here...`r`n" )
    pnputil /enum-drivers >c:\tobii.txt
    $TobiiDrivers = Get-WindowsDriver -Online -All | Where-Object { $_.ProviderName -eq "Tobii AB" } | Select-Object Driver , OriginalFileName
    ForEach ($drivers in $TobiiDrivers) {
        $inf = $drivers.Driver 
        $List = $drivers.OriginalFileName
        $List = $List.Replace("C:\Windows\System32\DriverStore\FileRepository\", "")
        $outputbox.appendtext("`r`n$inf : $List `r`n")
    }
    $outputbox.appendtext("`r`nDone")
}

#B5
Function ETfw {
    $outputBox.clear()
    $outputBox.appendtext( "Checking Eye tracker Firmware...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    try{ 
        $erroractionpreference = "Stop"
		$Firmware = .\FWUpgrade32.exe --auto --info-only 
	}
	Catch [System.Management.Automation.RemoteException] {
        $outputbox.appendtext("No Eye Tracker Connected`r`n")
    }
    $outputbox.appendtext($Firmware) 
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#B6
Function InstallPDK {
    $outputBox.clear()
    $serviceName = "TobiiIS5GIBBONGAZE"
    if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
        $outputbox.appendtext("$serviceName Service is already installed`r`n")
    }
    else {
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Installing PDK on I-Series`r`n")
        sc.exe create $serviceName binpath="C:\Windows\System32\DriverStore\FileRepository\is5gibbongaze.inf_amd64_07ff964b2ca8d0e4\platform_runtime_IS5GIBBONGAZE_service.exe" DisplayName= "Tobii Runtime Service" start= auto
        Start-Service -Name $serviceName -ErrorAction stop
    }
    $outputbox.appendtext("Done!")
}

#B7 Stops all currently active tobii processes
Function RestartProcesses {
    $outputBox.clear()
    $Outputbox.Appendtext( "Restart Services...`r`n")
    $StopServices = Get-Service -Name '*Tobii*' | Stop-Service -force -Passthru -erroraction ignore | Select Name, Status | Format-table -hidetableheaders | Out-string
    $Outputbox.Appendtext( "Stopping following Services:$StopServices")

    Start-Sleep -s 3
    $Processkill = get-process "GazeSelection" , "*TobiiDynavox*", "*Tobii.EyeX*", "Notifier" | Stop-process -force -Passthru -erroraction ignore | Select Processname | Format-table -Hidetableheaders | Out-string
    $Outputbox.Appendtext( "Stopping following processes:$Processkill")

    #start all processes and services
    Start-Sleep -s 3
    try {
        $StartServices = Start-Service -Name '*Tobii*' -ErrorAction Stop
        $Outputbox.Appendtext( "Attempting to start following Services:$StartServices `r`n" )
    }
    Catch {
        $Outputbox.Appendtext( "Tobii Service failed to start!`r`n" )
    }
    Start-Sleep -s 3
    try {
        $StartProcesses = Start-process "C:\Program Files (x86)\Tobii Dynavox\Eye Assist\TobiiDynavox.EyeAssist.Engine.exe"
        $Outputbox.Appendtext( "Attempting to start Eyeassist:$StartProcesses `r`n" )
    }
    Catch {
        $outputBox.Appendtext( "EyeAssist failed to start!`r`n" )
    }
    $outputBox.Appendtext( "Done!" )
}

#B8
Function DeleteServices {
    $outputBox.clear()
    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove IS5 services.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove services`r`n" )
        Return
    }
    $DeleteServices = Get-Service -Name '*TobiiIS*' , '*TobiiG*' | Stop-Service -Force -passthru -ErrorAction ignore
    $outputBox.appendtext("Deleting following Services:`r`n$DeleteServices`r`n")
    foreach ($Service in $DeleteServices) {
        sc.exe delete $Service
    }
    $outputbox.appendtext("Done! `r`n")
}

#B9
Function RemoveDrivers {
    $outputBox.clear()
    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove IS5 Drivers.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove driver`r`n" )
        Return
    }
    $outputBox.appendtext( "Removing Tobii Drivers...`r`n" )
    $TobiiVer = Get-WindowsDriver -Online -All | Where-Object { $_.ProviderName -eq "Tobii AB" } | Select-Object Driver
    ForEach ($ver in $TobiiVer) {
        pnputil /delete-driver $ver.Driver /force /uninstall
    }
    $outputbox.appendtext("`r`nDone")
}

#B10
Function resetBOOT {
    $outputBox.clear()
    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will reset ET to bootloader.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: reset Bootloader`r`n" )
        Return
    }
    $outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
    Get-Service -Name '*TobiiIS*' , '*TobiiG*' , '*Tobii Serivce*' | Stop-Service -Force -passthru -ErrorAction ignore
        
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "CastorUsbCli.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    .\CastorUsbCli.exe --reset BOOT
    $test2 = Get-CimInstance Win32_PnPSignedDriver | Where-Object Description -Like "*WinUSB Device*" | Select-Object DeviceID
    $outputbox.appendtext("The reset is done. ET PID is now:$test2")
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Services will be started now!")
    Start-Service -Name '*Tobii*' -ErrorAction Stop
    $outputbox.appendtext("`r`nDone `r`n")
}

#B11
Function FWUpgrade {
    $outputBox.clear()
    $outputBox.appendtext( "Upgrade IS4 ET FW...`r`n" )
	$path = "C:\Program Files (x86)\Tobii\Service"
    if (Test-Path $path) {
		Set-Location -path $path
        $ETInfo = .\FWUpgrade32.exe --auto --info-only
        $outputbox.appendtext("Connected ET is: $ETInfo")
		$outputbox.appendtext("`r`n")
	}

    else {
        $outputbox.appendtext("No Eye Tracker Connected`r`n")
    }
    if ($ETInfo -match "PCE1M") {
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Upgrading PCEye mini FW..")
        $PCEyeMini = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4pceyemini_firmware_2.27.0-4014648.tobiipkg" --no-version-check
        $outputbox.appendtext($PCEyeMini)
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Done `r`n")
    }
    elseif ($ETInfo -match "IS4_Large_102") {
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Upgrading PCEye Plus FW..")
        $PCEyePlus = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4large102_firmware_2.27.0-4014648.tobiipkg" --no-version-check
        $outputbox.appendtext($PCEyePlus)
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Done `r`n")
    }
    elseif ($ETInfo -match "IS4_Large_Peripheral") {
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Upgrading 4C FW..")
        $4C = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4largetobiiperipheral_firmware_2.27.0-4014648.tobiipkg" --no-version-check
        $outputbox.appendtext($4C)
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Done `r`n")
    }
    elseif ($ETInfo -match "IS4_Base_I-series") {
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Upgrading I-Series+ FW..")
        $ISeries = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii Dynavox\Gaze Interaction\Eye Tracker Firmware Releases\I-Series\iseries_firmware_1.2.4.33069.20150521.1055.root.tobiipkg" --no-version-check
        $outputbox.appendtext($ISeries)
        $outputbox.appendtext("`r`n")
        $outputbox.appendtext("Done. Restart ET through Control Center `r`n")
    }
    else {
        $outputbox.appendtext("No ET connected or ET not supported")
    }
}

#B12
Function BeforeUninstallGG {
    $outputBox.clear()
    $outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "BeforeUninstall.bat" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    $Installer = cmd /c "BeforeUninstall.bat"
    $Outputbox.appendtext($Installer)
    $outputbox.appendtext("`r`n")
    $Outputbox.appendtext( "Done! `r`n" )
    $outputbox.appendtext("`r`n")
}

#B13
Function ETConnection {
    $outputBox.clear()
    $outputBox.appendtext( "Running ET connection check...`r`n" )
    $outputBox.appendtext( "Results of output will be stored in C:\Output.txt...`r`n" )
    $outputbox.appendtext("`r`n")
    $a = 1
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'Loop'
    $msg = 'Enter number of loops:'
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    Do {
        Start-sleep -s 1
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        Set-Location $fpath
		try{ 
			$erroractionpreference = "Stop"
			$getinfo = cmd /c "FWUpgrade32.exe" --auto --info-only | out-string
		}
		catch [System.Management.Automation.RemoteException] {
			$outputbox.appendtext("No Eye Tracker Connected`r`n")
		}
        $time = Get-Date -UFormat %H:%M:%S
        Add-content C:\Output.txt $time, $getinfo
        $a
        $outputbox.appendtext($getinfo)
        $outputbox.appendtext("`r`n")
        $a++
    } while ($a -le $b)
    $outputbox.appendtext("Done `r`n")
}

#B14
Function HBTool {
    $outputBox.clear()
    $outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
    Start-Process -FilePath C:\Users\aes\Desktop\script\BeforeUninstall.bat
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("ACTIVE PROCESSES: `r`n")
    $outputbox.appendtext("Done `r`n")
}

#B15
Function EAProfileCreation {
    $outputBox.clear()
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'SMBios tool'
    $msg = "1 to create profile based on default, `r`n 2 create as many profiles and calibrate"
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    
    if ($b -match "1") { 
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $title = 'Profile Name'
        $msg = "Write a profile name"
        $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
        Set-Location 'C:\Program Files (x86)\Tobii Dynavox\Eye Assist'
        .\TobiiDynavox.EyeAssist.Smorgasbord.exe --createprofilewithdefaultcalibration --profile $b
        $outputbox.appendtext("Profile with $b has been created")
        $outputbox.appendtext("`r`n")

    }
    elseif ($b -match "2") {
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $a = 1
        $title = 'Loop'
        $msg = 'Enter number of loops:'
        $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
        Set-Location "C:\Program Files (x86)\Tobii Dynavox\Eye Assist"
        Do {
            Start-sleep -s 1
            $a
            $profile = .\TobiiDynavox.EyeAssist.Smorgasbord.exe --startcreateprofileandcalibrate --profile $a
            $outputbox.appendtext("Creating profile with name: $b")
            Start-sleep -s 10
            $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
            #Set-Location $fpath
			try{ 
				$erroractionpreference = "Stop"
				$getinfo = cmd /c "$fpath\FWUpgrade32.exe" --auto --info-only | out-string
		    }
			catch [System.Management.Automation.RemoteException] {
				 $outputbox.appendtext( "No Eye Tracker Connected`r`n")
			}
            $time = Get-Date -UFormat %H:%M:%S
            Add-content c:\Output.txt $time, $profile, $getinfo
            Start-sleep -s 3
            .\TobiiDynavox.EyeAssist.Engine.exe -x
            Start-sleep -s 3
            .\TobiiDynavox.EyeAssist.Engine.exe
            Start-sleep -s 3
            $a++
        } while ($a -le $b)
    }
    else { $outputbox.appendtext("N/A") }
    $outputbox.appendtext("Done `r`n")
}

#B16
Function RetrieveUnreleased {
    $outputBox.clear()
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'UN Activate Unreleased tool'
    $msg = "Press 1 to set value to True, `r`n 2 to set the value to False, `r`n 3 to remove the key"
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

    if ($b -match "1") { 
        New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Tobii\Update Notifier' -Name "RetrieveUnreleasedVersions" -PropertyType "String" -Value 'True'
        $outputbox.appendtext("Value set to True")
    }
    elseif ($b -match "2") { 
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Tobii\Update Notifier' -Name "RetrieveUnreleasedVersions" -Value 'False'
        $outputbox.appendtext("Value set to False") 
    }
    elseif ($b -match "3") {
        Remove-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Tobii\Update Notifier' -Name "RetrieveUnreleasedVersions"
        $outputbox.appendtext("String has been removed") 
    }
    else { $outputbox.appendtext("N/A") }
    $outputbox.appendtext("Done `r`n")
}

#B17
Function WCF {
    $outputBox.clear()
    $outputBox.appendtext( "Checking WCF Endpoint Blocking Software...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "handle.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    Start-Process cmd "/c  `"handle.exe net.pipe & pause `""
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#B18
Function SMBios {
    $outputBox.clear()
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'SMBios tool'
    $msg = "Press 1 to run getSMBIOSvalues.cmd, `r`n 2 setName.cmd, `r`n 3 setSerialNumber.cmd, `r`n 4 setVendor.cmd"
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "getSMBIOSvalues.cmd" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath

    if ($b -match "1") { 
        $getvaluses = Start-Process -FilePath .\getSMBIOSvalues.cmd
    }
    elseif ($b -match "2") {
        $getvaluses = Start-Process -FilePath .\setName.cmd
    }
    elseif ($b -match "3") { 
        $getvaluses = Start-Process -FilePath .\setSerialNumber.cmd
    }
    elseif ($b -match "4") { 
        $getvaluses = Start-Process -FilePath .\setVendor.cmd
    }
    else { $outputbox.appendtext("N/A") }

    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#B19
Function ETSamples {
    $outputBox.appendtext( "Starting TD region interaction sample...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "TobiiDynavox.RegionInteraction.Sample.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    .\TobiiDynavox.RegionInteraction.Sample.exe
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#B20
Function Diagnostic {
    $outputBox.appendtext( "Run diagnostics application for Interaction...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "Tobii.EyeX.Diagnostics.Application.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    .\Tobii.EyeX.Diagnostics.Application.exe
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#B21
Function SETest {
    $outputBox.appendtext( "running Stream Engine Test...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "tests.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    .\tests.exe
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#B22
Function InternalSE {
    $outputBox.appendtext( "Starting Stream Engine Sample app...`r`n" )
    $fpath = (Get-ChildItem -Path "$PSScriptRoot" -Filter "sample.exe" -Recurse).FullName | Split-Path 
    Set-Location $fpath
    start .\sample.exe
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#B23
function GetFrameworkVersionsAndHandleOperation() {
    $outputBox.clear()
    $installedFrameworks = @()
    if (IsKeyPresent "HKLM:\Software\Microsoft\.NETFramework\Policy\v1.0" "3705") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 1.0`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v1.1.4322" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 1.1`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v2.0.50727" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 2.0`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.0\Setup" "InstallSuccess") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 3.0`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.5" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 3.5`r`n" )}
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 4.0c`r`n" )}
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 4.0`r`n" )}   

    $result = -1
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client" "Install" -or IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Install") {
        # .net 4.0 is installed
        $result = 0
        $version = GetFrameworkValue "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Release"
        
		if ($version -ge 528040 -Or $version -ge 528372 -Or $version -ge 528049) {
            # .net 4.8
            $outputbox.appendtext( "Installed .Net Framework 4.8")
            $result = 10
        }
		elseif ($version -ge 461808 -Or $version -ge 461814) {
            # .net 4.7.2
           $outputbox.appendtext("Installed .Net Framework 4.7.2")
            $result = 9
        }
        elseif ($version -ge 461308 -Or $version -ge 461310) {
            # .net 4.7.1
            $outputbox.appendtext( "Installed .Net Framework 4.7.1")
            $result = 8
        }
        elseif ($version -ge 460798 -Or $version -ge 460805) {
            # .net 4.7
            $outputbox.appendtext( "Installed .Net Framework 4.7")
            $result = 7
        }
        elseif ($version -ge 394802 -Or $version -ge 394806) {
            # .net 4.6.2
            $outputbox.appendtext( "Installed .Net Framework 4.6.2")
            $result = 6
        }
        elseif ($version -ge 394254 -Or $version -ge 394271) {
            # .net 4.6.1
            $outputbox.appendtext( "Installed .Net Framework 4.6.1")
            $result = 5
        }
        elseif ($version -ge 393295 -Or $version -ge 393297) {
            # .net 4.6
            $outputbox.appendtext( "Installed .Net Framework 4.6")
            $result = 4
        }
        elseif ($version -ge 379893) {
            # .net 4.5.2
            $outputbox.appendtext( "Installed .Net Framework 4.5.2")
            $result = 3
        }
        elseif ($version -ge 378675) {
            # .net 4.5.1
            $outputbox.appendtext( "Installed .Net Framework 4.5.1")
            $result = 2
        }
        elseif ($version -ge 378389) {
            # .net 4.5
            $outputbox.appendtext( "Installed .Net Framework 4.5")
            $result = 1
        }   
    }
    else {
        # .net framework 4 family isn't installed
        $result = -1
    }
    
    return $result    
	#$version = GetFramework40FamilyVersion;
    return $installedFrameworks
    
    if ($version -ge 1) { 
    }
    else { }
}

function IsKeyPresent([string]$path, [string]$key) {
    if (!(Test-Path $path)) { return $false }
    if ((Get-ItemProperty $path).$key -eq $null) { return $false }
    return $true
}
function GetFrameworkValue([string]$path, [string]$key) {
    if (!(Test-Path $path)) { return "-1" }
    return (Get-ItemProperty $path).$key  
}

#B24
Function BatteryLog {
    $outputBox.appendtext( "Starting TobiiDynavox.QA.BatteryMonitor.exe...`r`n" )
    $fpath = (Get-ChildItem -Path "$PSScriptRoot" -Filter "TobiiDynavox.QA.BatteryMonitor.exe" -Recurse).FullName | Split-Path 
    Set-Location $fpath
    start-process .\TobiiDynavox.QA.BatteryMonitor.exe
    $outputbox.appendtext("Results will be saves in $fpath\battery_log.csv")
    $outputbox.appendtext("`r`n")
    $outputbox.appendtext("Done `r`n")
}

#Windows forms
$Optionlist = @("Remove PCEye5 Bundle", "Remove Tobii Device Drivers For Windows", "Remove WC&GP Bundle", "Remove PCEye Package", "Remove Communicator", "Remove Compass", "Remove TGIS only", "Remove TGIS profile calibrations", "Remove all users C5", "Reset TETC", "Backup Gaze Interaction", "Copy License")
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(600, 550)
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False

#Informationtext above the dropdown list.
$DropDownLabel = new-object System.Windows.Forms.Label
$DropDownLabel.Location = new-object System.Drawing.Size(10, 10)
$DropDownLabel.size = new-object System.Drawing.Size(160, 20)
$DropDownLabel.Text = "Select an option"
$Form.Controls.Add($DropDownLabel)

#Dropdown list with options
$DropDownBox = New-Object System.Windows.Forms.ComboBox
$DropDownBox.Location = New-Object System.Drawing.Size(10, 30)
$DropDownBox.Size = New-Object System.Drawing.Size(220, 20)
$DropDownBox.DropDownHeight = 230
$Form.Controls.Add($DropDownBox)

#For each arrayitem in optionlist, add it to $dropdownbox items.
foreach ($option in $optionlist) {
    $DropDownBox.Items.Add($option)
}

#Outputbox
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Size(10, 150)
$outputBox.Size = New-Object System.Drawing.Size(400, 340)
$outputBox.MultiLine = $True
$outputBox.ScrollBars = "Vertical"
$Form.Controls.Add($outputBox)
$outputBox.font = New-Object System.Drawing.Font ("Consolas" , 8, [System.Drawing.FontStyle]::Regular)

#Button "Start"
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(10, 60)
$Button.Size = New-Object System.Drawing.Size(110, 50)
$Button.Text = "Start"
$Button.Font = New-Object System.Drawing.Font ("" , 12, [System.Drawing.FontStyle]::Regular)
$Form.Controls.Add($Button)
$Button.Add_Click{ selectedscript }

#B1 Button1 "List Tobii Software"
$Button1 = New-Object System.Windows.Forms.Button
$Button1.Location = New-Object System.Drawing.Size(250, 0)
$Button1.Size = New-Object System.Drawing.Size(160, 30)
$Button1.Text = "Tobii Software"
$Button1.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button1)
$Button1.Add_Click{ ListApps }

#B2 Button2 "List active Tobii processes"
$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(250, 30)
$Button2.Size = New-Object System.Drawing.Size(160, 30)
$Button2.Text = "Active Process+Service"
$Button2.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button2)
$Button2.Add_Click{ GetProcesses }

#B3 Button3 "Check IS5 PID"
$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Size(250, 60)
$Button3.Size = New-Object System.Drawing.Size(160, 30)
$Button3.Text = "List IS5 PID"
$Button3.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button3)
$Button3.Add_Click{ IS5PID }

#B4 Button4 "List Tobii Drivers"
$Button4 = New-Object System.Windows.Forms.Button
$Button4.Location = New-Object System.Drawing.Size(250, 90)
$Button4.Size = New-Object System.Drawing.Size(160, 30)
$Button4.Text = "Tobii Drivers"
$Button4.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button4)
$Button4.Add_Click{ ListDrivers }

#B5 Button5 "ET fw"
$Button5 = New-Object System.Windows.Forms.Button
$Button5.Location = New-Object System.Drawing.Size(250, 120)
$Button5.Size = New-Object System.Drawing.Size(160, 30)
$Button5.Text = "ET firmware"
$Button5.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button5)
$Button5.Add_Click{ ETfw }

#B6 Button6 "Install PDK"
$Button6 = New-Object System.Windows.Forms.Button
$Button6.Location = New-Object System.Drawing.Size(410, 0)
$Button6.Size = New-Object System.Drawing.Size(160, 30)
$Button6.Text = "Install PDK"
$Button6.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button6)
$Button6.Add_Click{ InstallPDK }

#B7 Button7 Restart Services
$Button7 = New-Object System.Windows.Forms.Button
$Button7.Location = New-Object System.Drawing.Size(410, 30)
$Button7.Size = New-Object System.Drawing.Size(160, 30)
$Button7.Text = "Restart Services"
$Button7.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button7)
$Button7.Add_Click{ RestartProcesses }

#B8 Button8 "Delete services"
$Button8 = New-Object System.Windows.Forms.Button
$Button8.Location = New-Object System.Drawing.Size(410, 60)
$Button8.Size = New-Object System.Drawing.Size(160, 30)
$Button8.Text = "Delete Services"
$Button8.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button8)
$Button8.Add_Click{ DeleteServices }

#B9 Button9 "Remove Drivers"
$Button9 = New-Object System.Windows.Forms.Button
$Button9.Location = New-Object System.Drawing.Size(410, 90)
$Button9.Size = New-Object System.Drawing.Size(160, 30)
$Button9.Text = "Delete Drivers"
$Button9.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button9)
$Button9.Add_Click{ RemoveDrivers }

#B10 Button10 "Reset IS5 to bootloader"
$Button10 = New-Object System.Windows.Forms.Button
$Button10.Location = New-Object System.Drawing.Size(410, 120)
$Button10.Size = New-Object System.Drawing.Size(160, 30)
$Button10.Text = "Reset ET BOOT"
$Button10.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button10)
$Button10.Add_Click{ resetBOOT }

#B11 Button11 "FW Upgrade"
$Button11 = New-Object System.Windows.Forms.Button
$Button11.Location = New-Object System.Drawing.Size(410, 150)
$Button11.Size = New-Object System.Drawing.Size(160, 30)
$Button11.Text = "FW Upgrade"
$Button11.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button11)
$Button11.Add_Click{ FWUpgrade }

#B12 Button12 "Before Uninstall GG"
$Button12 = New-Object System.Windows.Forms.Button
$Button12.Location = New-Object System.Drawing.Size(410, 180)
$Button12.Size = New-Object System.Drawing.Size(160, 30)
$Button12.Text = "Before Uninstall GG"
$Button12.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button12)
$Button12.Add_Click{ BeforeUninstallGG }

#B13 Button13 "Check ET connection through Service"
$Button13 = New-Object System.Windows.Forms.Button
$Button13.Location = New-Object System.Drawing.Size(410, 210)
$Button13.Size = New-Object System.Drawing.Size(160, 30)
$Button13.Text = "Check ET connection"
$Button13.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button13)
$Button13.Add_Click{ ETConnection }

#B14 Button14 "Run GG HB_tool"
$Button14 = New-Object System.Windows.Forms.Button
$Button14.Location = New-Object System.Drawing.Size(410, 240)
$Button14.Size = New-Object System.Drawing.Size(160, 30)
$Button14.Text = "HB-tool GG"
$Button14.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button14)
$Button14.Add_Click{ HBTool }

#B15 Button15 "EAProfileCreation"
$Button15 = New-Object System.Windows.Forms.Button
$Button15.Location = New-Object System.Drawing.Size(410, 270)
$Button15.Size = New-Object System.Drawing.Size(160, 30)
$Button15.Text = "EA Profile Creation"
$Button15.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button15)
$Button15.Add_Click{ EAProfileCreation }

#B16 Button16 "RetrieveUnreleased"
$Button16 = New-Object System.Windows.Forms.Button
$Button16.Location = New-Object System.Drawing.Size(410, 300)
$Button16.Size = New-Object System.Drawing.Size(160, 30)
$Button16.Text = "RetrieveUnreleased"
$Button16.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button16)
$Button16.Add_Click{ RetrieveUnreleased }

#B17 Button17 "WCF"
$Button17 = New-Object System.Windows.Forms.Button
$Button17.Location = New-Object System.Drawing.Size(410, 330)
$Button17.Size = New-Object System.Drawing.Size(160, 30)
$Button17.Text = "WCF"
$Button17.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button17)
$Button17.Add_Click{ WCF }

#B18 Button18 "SMBios"
$Button18 = New-Object System.Windows.Forms.Button
$Button18.Location = New-Object System.Drawing.Size(410, 360)
$Button18.Size = New-Object System.Drawing.Size(160, 30)
$Button18.Text = "SMBIOS"
$Button18.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button18)
$Button18.Add_Click{ SMBios }

#B19 Button19 "ETSamples"
$Button19 = New-Object System.Windows.Forms.Button
$Button19.Location = New-Object System.Drawing.Size(410, 390)
$Button19.Size = New-Object System.Drawing.Size(160, 30)
$Button19.Text = "TD RI Samples"
$Button19.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button19)
$Button19.Add_Click{ ETSamples }

#B20 Button20 "Diagnostic"
$Button20 = New-Object System.Windows.Forms.Button
$Button20.Location = New-Object System.Drawing.Size(410, 420)
$Button20.Size = New-Object System.Drawing.Size(160, 30)
$Button20.Text = "RIDiagnostics"
$Button20.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button20)
$Button20.Add_Click{ Diagnostic }

#B21 Button21 "StreamEngineTest"
$Button21 = New-Object System.Windows.Forms.Button
$Button21.Location = New-Object System.Drawing.Size(410, 450)
$Button21.Size = New-Object System.Drawing.Size(160, 30)
$Button21.Text = "SE-Test"
$Button21.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button21)
$Button21.Add_Click{ SETest }

#B22 Button22 "InternalSE"
$Button22 = New-Object System.Windows.Forms.Button
$Button22.Location = New-Object System.Drawing.Size(410, 480)
$Button22.Size = New-Object System.Drawing.Size(160, 30)
$Button22.Text = "Internal SE"
$Button22.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button22)
$Button22.Add_Click{ InternalSE }

#B23 Button23 ".NET version"
$Button23 = New-Object System.Windows.Forms.Button
$Button23.Location = New-Object System.Drawing.Size(10, 110)
$Button23.Size = New-Object System.Drawing.Size(110, 30)
$Button23.Text = ".NET v."
$Button23.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button23)
$Button23.Add_Click{ GetFrameworkVersionsAndHandleOperation }

#B22 Button24 "BatteryLog"
$Button24 = New-Object System.Windows.Forms.Button
$Button24.Location = New-Object System.Drawing.Size(120, 110)
$Button24.Size = New-Object System.Drawing.Size(110, 30)
$Button24.Text = "Battery Log"
$Button24.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button24)
$Button24.Add_Click{ BatteryLog }

#Form name + activate form.
$Form.Text = "Support Tool 1.4"
$Form.Add_Shown( { $Form.Activate() })
$Form.ShowDialog()