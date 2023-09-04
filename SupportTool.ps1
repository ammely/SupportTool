#v1.6.4
#Forces powershell to run as an admin
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{ Start-Process powershell.exe "-NoProfile -Windowstyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

#Imports Windowsforms and Drawing from system
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

#Allows the use of wshell for confirmation popups
$wshell = New-Object -ComObject Wscript.Shell
$PSScriptRoot

#Links functions to selected option in the dropdown list, activates on button click
#Outputbox.clear() Erases text output from the outputbox before continuing with the script.
Function selectedscript {

    if ($DropDownBox.Selecteditem -eq "Remove Progressive Sweet") {
        $Outputbox.Clear()
        UninstallProgressiveSweet
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove PCEye5 Bundle") {
        $Outputbox.Clear()
        UninstallPCEye5Bundle
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove all ET SW") {
        $Outputbox.Clear()
        UninstallTobiiDeviceDriversForWindows
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove WC&GP Bundle") {
        $Outputbox.Clear()
        UninstallWCGP
    }
    elseif ($DropDownBox.Selecteditem -eq "Remove VC++") {
        $Outputbox.Clear()
        VCRedist
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

#A1 Uninstalls Progressive Sweet
Function UninstallProgressiveSweet {
    # https://stackoverflow.com/questions/46310266/accessing-dynamically-created-variables-inside-a-powershell-function
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $form = New-Object System.Windows.Forms.Form
    $flowlayoutpanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonOK = New-Object System.Windows.Forms.Button

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty  | Where-Object { 
        ($_.Displayname -eq "Tobii Dynavox Switcher") -or
        ($_.Displayname -eq "Tobii Dynavox Switcher (Beta)") -or
        ($_.Displayname -eq "Tobii Dynavox Browse") -or
        ($_.Displayname -eq "Tobii Dynavox Browse (Beta)") -or
        ($_.Displayname -eq "Tobii Dynavox Phone (Beta)") -or
        ($_.Displayname -eq "Tobii Dynavox Talk (Beta)") -or
        ($_.Displayname -eq "Tobii Dynavox Control (Beta)") -or
        ($_.Displayname -eq "Tobii Dynavox Control")
    } | Select-Object Displayname, UninstallString
    if ($TobiiVer) {   
        $usernames = @($TobiiVer.Displayname)
        $totalvalues = ($usernames.count)

        $formsize = 85 + (30 * $totalvalues)
        $flowlayoutsize = 10 + (30 * $totalvalues)
        $buttonplacement = 40 + (30 * $totalvalues)
        $script:CheckBoxArray = @()
    
        $form_Load = {
            foreach ($user in $usernames) {
                $DynamicCheckBox = New-object System.Windows.Forms.CheckBox

                $DynamicCheckBox.Margin = '10, 8, 0, 0'
                $DynamicCheckBox.Name = $user
                #changed to make the text look better
                $DynamicCheckBox.Size = '300, 22' 
                $DynamicCheckBox.Text = "" + $user

                $DynamicCheckBox.TextAlign = 'MiddleLeft'
                $flowlayoutpanel.Controls.Add($DynamicCheckBox)
                $script:CheckBoxArray += $DynamicCheckBox
            }       
        }
    
        $form.Controls.Add($flowlayoutpanel)
        $form.Controls.Add($buttonOK)
        $form.AcceptButton = $buttonOK
        $form.AutoScaleDimensions = '8, 17'
        $form.AutoScaleMode = 'Font'
        $form.ClientSize = "500 , $formsize"
        $form.FormBorderStyle = 'FixedDialog'
        $form.Margin = '5, 5, 5, 5'
        $form.MaximizeBox = $False
        $form.MinimizeBox = $False
        $form.Name = 'form1'
        $form.StartPosition = 'CenterScreen'
        $form.Text = 'Progressive Sweet'
        $form.add_Load($($form_Load))
    } 
    else { 
        $OutputBox.AppendText( "Empty. `r`n" )
    }
    $flowlayoutpanel.BorderStyle = 'FixedSingle'
    $flowlayoutpanel.Location = '48, 13'
    $flowlayoutpanel.Margin = '4, 4, 4, 4'
    $flowlayoutpanel.Name = 'flowlayoutpanel1'
    $flowlayoutpanel.AccessibleName = 'flowlayoutpanel1'
    $flowlayoutpanel.Size = "400, $flowlayoutsize"
    $flowlayoutpanel.TabIndex = 1
    
    $buttonOK.Anchor = 'Bottom, Right'
    $buttonOK.DialogResult = 'OK'
    $buttonOK.Location = "383, $buttonplacement"
    $buttonOK.Margin = '4, 4, 4, 4'
    $buttonOK.Name = 'buttonOK'
    $buttonOK.Size = '100, 30'
    $buttonOK.TabIndex = 0
    $buttonOK.Text = '&OK'
    
    $form.ShowDialog()
    foreach ($cbox in $CheckBoxArray) {
        if ($cbox.CheckState -eq "Unchecked") {
            $Outputbox.Appendtext( "No SW were selected`r`n" )
        }
        elseif ($cbox.CheckState -eq "Checked") {
            #If first answer equals yes or no
            $answer1 = $wshell.Popup("This will remove selected software.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
            if ($answer1 -eq 6) {
                $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress..`r`n" )
            }
            elseif ($answer1 -ne 6) {
                $Outputbox.Appendtext( "Action canceled: Remove Progressive sweet`r`n" )
                Return
            }
            $Uninstname = (Compare-Object -DifferenceObject $TobiiVer.displayname -ReferenceObject $cbox.Name -CaseSensitive -ExcludeDifferent -IncludeEqual | Select-Object InputObject).InputObject
            #$Outputbox.Appendtext( "Following apps will be removed $Uninstname`r`n" ) 
            if ($Uninstname -match "Beta") {
                $test = $Uninstname -replace '\(Beta\)', ""
            }

            $newname1 = $test + "Updater Service (Beta)"
            $newname2 = $test + "Launcher (Beta)"
            $newname3 = $Uninstname + " Updater Service"
            $newname4 = $Uninstname + " Launcher"

            $TobiiVer2 = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
            Get-ItemProperty  | Where-Object { 
                ($_.Displayname -eq "$Uninstname") -or 
                ($_.Displayname -eq "$newname1") -or
                ($_.Displayname -eq "$newname2") -or 
                ($_.Displayname -eq "$newname3") -or
                ($_.Displayname -eq "$newname4")
            } | Select-Object Displayname, UninstallString 
            foreach ( $tobiivers in $TobiiVer2) {
                $Displayname = $tobiivers.Displayname
                $Outputbox.Appendtext( "Removing - " + "$Displayname`r`n" )
                $uninst = $tobiivers.UninstallString -replace "msiexec.exe", "" -Replace "/I", "" -Replace "/X", ""

                start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
            }

            $BrowseBetaPath = (
                "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Browse Beta", 
                "$ENV:ProgramData\Tobii Dynavox\Browse Beta", 
                "$ENV:ProgramData\Tobii Dynavox\Pegasus Review",
                "HKCU:\Software\Tobii Dynavox\Browse Updater Service Beta" )
            $BrowsePath = (
                "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Browse", 
                "$ENV:ProgramData\Tobii Dynavox\Browse")
            $TalkBetaPath = (
                "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Talk Beta", 
                "$ENV:ProgramData\Tobii Dynavox\Talk Beta",
                "HKCU:\Software\Tobii Dynavox\Talk Updater Service Beta")
            $TalkPath = (
                "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Talk")
            $SwitcherBetaPath = (
                "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Switcher Beta", 
                "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\App Switcher Review", 
                "$ENV:ProgramData\Tobii Dynavox\Switcher Beta")
            $SwitcherPath = (
                "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Switcher", 
                "$ENV:ProgramData\Tobii Dynavox\Switcher",
                "$ENV:Program Files\Tobii Dynavox\Switcher",
                "HKCU:\Software\Tobii Dynavox\Switcher")
            $CCReviewPath = (
                "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control Review", 
                "$ENV:ProgramData\Tobii Dynavox\Computer Control Review",
                "HKCU:\Software\Tobii Dynavox\Computer Control Review",
                "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Computer Control Updater Service Review"
            )
            $CCPath = (
                "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control", 
                "$ENV:ProgramData\Tobii Dynavox\Computer Control",
                "HKCU:\Software\Tobii Dynavox\Computer Control",
                "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Computer Control",
                "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Computer Control Updater Service"
            )
            $PhoneBetaPath = (
                "$ENV:ProgramData\Tobii Dynavox\Phone Updater Service Beta",
                "$Env:USERPROFILE\AppData\Local\Tobii Dynavox\Phone Beta\",
                "HKCU:\Software\Tobii Dynavox\Phone Updater Service Review")

            if ($Uninstname -eq "Tobii Dynavox Browse (Beta)") {
                if (Test-Path $BrowseBetaPath) {
                    $Outputbox.appendtext( "Removing - " + "$BrowseBetaPath`r`n" )
                    Remove-Item $BrowseBetaPath -Recurse -Force -ErrorAction Ignore
                }
            }
            elseif ($Uninstname -eq "Tobii Dynavox Browse") {
                if (Test-Path $BrowsePath) {
                    $Outputbox.appendtext( "Removing - " + "$BrowsePath`r`n" )
                    Remove-Item $BrowsePath -Recurse -Force -ErrorAction Ignore
                }
            }
            elseif ($Uninstname -eq "Tobii Dynavox Talk (Beta)") {
                if (Test-Path $TalkBetaPath) {
                    $Outputbox.appendtext( "Removing - " + "$TalkBetaPath`r`n" )
                    Remove-Item $TalkBetaPath -Recurse -Force -ErrorAction Ignore
                }
            }
            elseif ($Uninstname -eq "Tobii Dynavox Switcher (Beta)") {
                if (Test-Path $SwitcherBetaPath) {
                    $Outputbox.appendtext( "Removing - " + "$SwitcherBetaPath`r`n" )
                    stop-process -Name "*switcher*" -Force
                    Remove-Item $SwitcherBetaPath -Recurse -Force -ErrorAction Ignore
                }
            }
            elseif ($Uninstname -eq "Tobii Dynavox Switcher") {
                if (Test-Path $SwitcherPath) {
                    $Outputbox.appendtext( "Removing - " + "$SwitcherPath`r`n" )
                    stop-process -Name "*switcher*" -Force
                    Remove-Item $SwitcherPath -Recurse -Force -ErrorAction Ignore
                }
            }
            elseif ($Uninstname -eq "Tobii Dynavox Control (Beta)") {
                if (Test-Path $CCReviewPath) {
                    $Outputbox.appendtext( "Removing - " + "$CCReviewPath`r`n" )
                    Remove-Item $CCReviewPath -Recurse -Force -ErrorAction Ignore
                }
            }
            elseif ($Uninstname -eq "Tobii Dynavox Control") {
                if (Test-Path $CCPath) {
                    $Outputbox.appendtext( "Removing - " + "$CCPath`r`n" )
                    Remove-Item $CCPath -Recurse -Force -ErrorAction Ignore
                }
            }
            elseif ($Uninstname -eq "Tobii Dynavox Phone (Beta)") {
                if (Test-Path $PhoneBetaPath) {
                    $Outputbox.appendtext( "Removing - " + "$PhoneBetaPath`r`n" )
                    Remove-Item $PhoneBetaPath -Recurse -Force -ErrorAction Ignore
                }
            }
        }
    }
    
    Remove-Variable checkbox*

    $Outputbox.Appendtext( "Done!`r`n" )
}

#A2 Uninstalls PCEye5 Bundle
Function UninstallPCEye5Bundle {

    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove all software included in PCEye5 bundle.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress..`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove PCEye5 bundle`r`n" )
        Return
    }
	
    $RegPath = "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig"
    $TempPath = "$ENV:USERPROFILE\AppData\Local\Temp\EyeXConfig.reg"
    if ((Test-Path -Path $RegPath) -and (!(Test-Path -path $TempPath))) {
        $Outputbox.Appendtext("Backup profiles in %temp%\EyeXConfig.reg`r`n")
        Invoke-Command { reg export "HKLM\SOFTWARE\WOW6432Node\Tobii\EyeXConfig" $TempPath }
    }

    $GetProcess = stop-process -Name "*TobiiDynavox*" -Force
    if ($GetProcess) {
        $Outputbox.appendtext("Stopping $GetProcess `r`n" )
    }

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { 
        ($_.Displayname -eq "Tobii Dynavox Control") -or
        ($_.Displayname -eq "Tobii Dynavox Computer Control") -or
        ($_.Displayname -Match "Tobii Dynavox Update Notifier") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking") -or
        ($_.Displayname -Eq "Tobii Device Drivers For Windows (PCEye5)") -or
        ($_.Displayname -Eq "Tobii Experience Software For Windows (PCEye5)") -or
        ($_.Displayname -eq "Tobii Dynavox Control ") -or
        ($_.Displayname -eq "Tobii Dynavox Control Updater Service") -or 
        ($_.Displayname -eq "Tobii Dynavox Switcher") -or 
        ($_.Displayname -eq "Tobii Dynavox Switcher Updater Service")
    } | Select-Object Displayname, UninstallString
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString -replace "msiexec.exe", "" -Replace "/I", "" -Replace "/X", ""
        $uninst = $uninst.Trim()
        $Outputbox.Appendtext( "Uninstalling - " + "$Uninstname`r`n" )
        start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
    }

    $DeleteServices = Get-Service -Name '*TobiiIS*' , '*TobiiG*' | Stop-Service -Force -passthru -ErrorAction ignore
    foreach ($Service in $DeleteServices) {
        $outputbox.appendtext(" Deleating - " + "$Service `r`n" )
        sc.exe delete $Service
    }

    $TobiiVer = Get-WindowsDriver -Online -All | Where-Object { $_.ProviderName -eq "Tobii AB" } | Select-Object Driver
    ForEach ($ver in $TobiiVer) {
        $outputBox.appendtext( "Removing Drivers - " + "$TobiiVer`r`n" )
        pnputil /delete-driver $ver.Driver /force /uninstall
    }
    stop-process -Name "*switcher*" -Force
    #Removes WC related folders
    $paths = (
        "C:\Program Files (x86)\Tobii Dynavox\Eye Tracking Settings",	
        "C:\Program Files (x86)\Tobii Dynavox\Eye Assist",
        "C:\Program Files (x86)\Tobii Dynavox\Update Notifier",
        "C:\Program Files\Tobii\Tobii EyeX",
        "$ENV:Program Files\Tobii Dynavox\Switcher",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\App Switcher",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control Bundle",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Update Notifier",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Eye Tracking",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Switcher", 
        "$ENV:ProgramData\Tobii Dynavox\Switcher",
        "$ENV:ProgramData\Tobii Dynavox\EyeAssist",
        "$ENV:ProgramData\Tobii Dynavox\Computer Control",
        "$ENV:ProgramData\Tobii Dynavox\Update Notifier",
        "$ENV:ProgramData\Tobii\Statistics",
        "$ENV:ProgramData\Tobii\Tobii Interaction",
        "$ENV:ProgramData\Tobii\Tobii Platform Runtime",
        "$ENV:ProgramData\Tobii\EulaHasBeenAccepted.txt",

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
        "HKLM:\SOFTWARE\WOW6432Node\Tobii Dynavox\Computer Control Updater Service",
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation",
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\Update Notifier",
        "HKLM:\SOFTWARE\WOW6432Node\Tobii Dynavox\Computer Control Updater Service Review")

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }
        
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { 
        ($_.Displayname -Match "Tobii Dynavox Computer Control") -or
        ($_.Displayname -Match "Dynavox Computer Control Updater Service") -or
        ($_.Displayname -Match "Tobii Dynavox Update Notifier") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking") -or
        ($_.Displayname -Eq "Tobii Device Drivers For Windows (PCEye5)") -or
        ($_.Displayname -Eq "Tobii Experience Software For Windows (PCEye5)") } | Select-Object Displayname
    if ($TobiiVer) {
        $outputBox.appendtext( "$TobiiVer couldn't be uninstalled. Reboot your device and try again.`r`n" )
    }
    $Outputbox.Appendtext( "Done!`r`n" )
}

#A3 Uninstalls ALL Tobii Device Drivers For Windows Bundle
Function UninstallTobiiDeviceDriversForWindows {

    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove all software included in Tobii Device Drivers.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress..`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove Tobii Device Drivers`r`n" )
        Return
    }
    
   	$RegPath = "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig"
    $TempPath = "$ENV:USERPROFILE\AppData\Local\Temp\EyeXConfig.reg"
    if ((Test-Path -Path $RegPath) -and (!(Test-Path -path $TempPath))) {
        $Outputbox.Appendtext("Backup profiles in %temp%\EyeXConfig.reg`r`n" )
        Invoke-Command { reg export "HKLM\SOFTWARE\WOW6432Node\Tobii\EyeXConfig" $TempPath }
    }

    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        try { 
            $erroractionpreference = "Stop"
            $Firmware = .\FWUpgrade32.exe --auto --info-only 
            $outputbox.appendtext("$Firmware`r`n")
        }
        catch [System.Management.Automation.RemoteException] {
            $outputbox.appendtext("PDK is not installed`r`n")
        }
    }
    else { 
        $outputbox.appendtext("File FWUpgrade32.exe is missing!`r`n" )
    }

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Device Drivers For Windows") } | Select-Object Displayname, DisplayVersion, UninstallString

    if ($Firmware -match "IS5_Gibbon_Gaze" -and $TobiiVer.DisplayVersion -eq "4.49.0.4000" ) { 
        $outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "BeforeUninstall.bat" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        if ($fpath.count -gt 0) {
            Set-Location $fpath
            $Installer = cmd /c "BeforeUninstall.bat"
            $Outputbox.appendtext("$Installer`r`n")
        }
        else { 
            $outputbox.appendtext("File BeforeUninstall.bat is missing!`r`n" )
        }
        $Outputbox.appendtext( "Done!`r`n" )
    } 
    else { $outputbox.appendtext( "No need to run the script`r`n") }

    $GetProcess = stop-process -Name "*TobiiDynavox*" -Force
    if ($GetProcess) {
        $Outputbox.appendtext("Stopping $GetProcess `r`n" )
    }

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { 
        ($_.Displayname -Match "Tobii Device Drivers For Windows") -or
        ($_.Displayname -Match "Tobii Experience Software") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking Driver") -or
        ($_.Displayname -Match "Tobii Eye Tracking For Windows") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking") -or
        ($_.Displayname -Match "Tobii Eye Tracking") } | Select-Object Displayname, UninstallString
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString
        $Outputbox.Appendtext( "Uninstalling - " + "$Uninstname`r`n" )
        $uninst = $ver.UninstallString -replace "msiexec.exe", "" -Replace "/I", "" -Replace "/X", "" -replace "/uninstall", ""
        $uninst = $uninst.Trim()
        if ($uninst -match "ProgramData") {
            try {
                cmd /c $uninst /uninstall /quiet
            }
            catch { 
                Write-Output "not"
            }
        }
        else {
            start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
        }
    }
    
    if (Get-AppxPackage *TobiiAB.TobiiEyeTrackingPortal*) {
        $outputBox.appendtext( "Removing Tobii Experience software.`r`n" )
        Get-AppxPackage *TobiiAB.TobiiEyeTrackingPortal* | Remove-AppxPackage
    }

    $DeleteServices = Get-Service -Name '*TobiiIS*' , '*TobiiG*' | Stop-Service -Force -passthru -ErrorAction ignore
    foreach ($Service in $DeleteServices) {
        $outputbox.appendtext(" Deleating - " + "$Service `r`n" )
        sc.exe delete $Service
    }
        
    $TobiiVer = Get-WindowsDriver -Online -All | Where-Object { $_.ProviderName -eq "Tobii AB" } | Select-Object Driver
    ForEach ($ver in $TobiiVer) {
        $outputBox.appendtext( "Removing Drivers - " + "$TobiiVer`r`n" )
        pnputil /delete-driver $ver.Driver /force /uninstall
    }

    #Removes Tobii related folders
    $paths = ( 
        "C:\Program Files\Tobii\Tobii EyeX",
        "$ENV:ProgramData\TetServer",
        "$ENV:ProgramData\Tobii\HelloDMFT",
        "$ENV:ProgramData\Tobii\Statistics",
        "$ENV:ProgramData\Tobii\Tobii Interaction",
        "$ENV:ProgramData\Tobii\Tobii Stream Engine",
        "$ENV:ProgramData\Tobii\Statistics",
        "$ENV:ProgramData\Tobii\Tobii Interaction",
        "$ENV:ProgramData\Tobii\Tobii Platform Runtime",
        "$ENV:ProgramData\Tobii\EulaHasBeenAccepted.txt",
        "$Env:USERPROFILE\AppData\Local\Tobii_AB\"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }
    #Deleting registry keys related to WC
    $Keys = ( 
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeX",
        "HKCU:\Software\Tobii\EyeAssist",
        "HKCU:\Software\Tobii\EyeX",
        "HKCU:\Software\Tobii\Vouchers",
        "HKCU:\Software\Tobii\GameHub"
    )

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }
    
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { 
        ($_.Displayname -Match "Tobii Device Drivers For Windows") -or
        ($_.Displayname -Match "Tobii Experience Software") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking Driver") -or
        ($_.Displayname -Match "Tobii Eye Tracking For Windows") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking") -or
        ($_.Displayname -Match "Tobii Eye Tracking") } | Select-Object Displayname
    $TobiiVer = $TobiiVer.DisplayName
    if ($TobiiVer) {
        $outputBox.appendtext( "$TobiiVer couldn't be uninstalled. Reboot your device and try again.`r`n" )
    }
    $Outputbox.appendtext( "Done!`r`n" )
}

#A4 Uninstalls WC Bundle
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


    $RegPath = "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig"
    $TempPath = "$ENV:USERPROFILE\AppData\Local\Temp\EyeXConfig.reg"
    if ((Test-Path -Path $RegPath) -and (!(Test-Path -path $TempPath))) {
       	$Outputbox.Appendtext("Backup profiles in %temp%\EyeXConfig.reg`r`n" )
        Invoke-Command { reg export "HKLM\SOFTWARE\WOW6432Node\Tobii\EyeXConfig" $TempPath }
    }


    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Windows Control") -or
        ($_.Displayname -Match "Virtual Remote") -or
        ($_.Displayname -Match "Update Notifier") -or
        ($_.Displayname -Match "Tobii Eye Tracking") -or
        ($_.Displayname -Match "GazeSelection") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze Point") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze Point Configuration Guide") } | Select-Object Displayname, UninstallString

    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        & cmd /c $uninst /quiet /norestart
    }

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
        "$ENV:ProgramData\Tobii Dynavox\Gaze Selection",
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

    $Outputbox.Appendtext( "Done!`r`n" )
}

#A5 Uninstalls VC++ redist
Function VCRedist {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $form = New-Object System.Windows.Forms.Form
    $flowlayoutpanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonOK = New-Object System.Windows.Forms.Button


    $x = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\ , HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | 
    Get-ItemProperty  | Where-Object { 
        ($_.Displayname -like "Microsoft Visual C++ 2005 Redistributable*") -or
        ($_.Displayname -like "Microsoft Visual C++ 2008 Redistributable *") -or
        ($_.Displayname -like "Microsoft Visual C++ 2010 * Redistributable *") -or
        ($_.Displayname -like "Microsoft Visual C++ 2012 Redistributable *") -or
        ($_.Displayname -like "Microsoft Visual C++ 2013 Redistributable *") -or
        ($_.Displayname -like "Microsoft Visual C++ 2015* Redistributable *") -or
        ($_.Displayname -like "Microsoft Visual C++ 2017 Redistributable *")
    } | Select-Object Displayname, UninstallString  


    $uninst = $x.UninstallString    

    $usernames = @($x.Displayname) | Sort-Object -Unique
    $totalvalues = ($usernames.count)

    $formsize = 85 + (30 * $totalvalues)
    $flowlayoutsize = 10 + (30 * $totalvalues)
    $buttonplacement = 40 + (30 * $totalvalues)
    $script:CheckBoxArray = @()
    
    $form_Load = {
        foreach ($user in $usernames) {
            $DynamicCheckBox = New-object System.Windows.Forms.CheckBox

            $DynamicCheckBox.Margin = '10, 8, 0, 0'
            $DynamicCheckBox.Name = $user
            #changed to make the text look better
            $DynamicCheckBox.Size = '400, 22' 
            $DynamicCheckBox.Text = "" + $user

            $DynamicCheckBox.TextAlign = 'MiddleLeft'
            $flowlayoutpanel.Controls.Add($DynamicCheckBox)
            $script:CheckBoxArray += $DynamicCheckBox
        }       
    }
    
    $form.Controls.Add($flowlayoutpanel)
    $form.Controls.Add($buttonOK)
    $form.AcceptButton = $buttonOK
    $form.AutoScaleDimensions = '8, 17'
    $form.AutoScaleMode = 'Font'
    $form.ClientSize = "600 , $formsize"
    $form.FormBorderStyle = 'FixedDialog'
    $form.Margin = '5, 5, 5, 5'
    $form.MaximizeBox = $False
    $form.MinimizeBox = $False
    $form.Name = 'form1'
    $form.StartPosition = 'CenterScreen'
    $form.Text = 'VC++'
    $form.add_Load($($form_Load))

    $flowlayoutpanel.BorderStyle = 'FixedSingle'
    $flowlayoutpanel.Location = '48, 13'
    $flowlayoutpanel.Margin = '4, 4, 4, 4'
    $flowlayoutpanel.Name = 'flowlayoutpanel1'
    $flowlayoutpanel.AccessibleName = 'flowlayoutpanel1'
    $flowlayoutpanel.Size = "500, $flowlayoutsize"
    $flowlayoutpanel.TabIndex = 1
    
    $buttonOK.Anchor = 'Bottom, Right'
    $buttonOK.DialogResult = 'OK'
    $buttonOK.Location = "383, $buttonplacement"
    $buttonOK.Margin = '4, 4, 4, 4'
    $buttonOK.Name = 'buttonOK'
    $buttonOK.Size = '100, 30'
    $buttonOK.TabIndex = 0
    $buttonOK.Text = '&OK'

    $form.ShowDialog()
    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove selected software.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress..`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove VC++`r`n" )
        Return
    }
    foreach ($cbox in $CheckBoxArray) {
        if ($cbox.CheckState -eq "Unchecked") {
           
        }
        elseif ($cbox.CheckState -eq "Checked") {
           
            $remove = $cbox.Name
            $Uninstname = (Compare-Object -DifferenceObject $x.displayname -ReferenceObject $cbox.Name -CaseSensitive -ExcludeDifferent -IncludeEqual | Select-Object InputObject).InputObject
            $tobiivers = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\ , HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { ($_.Displayname -eq "$Uninstname") } | Select-Object Displayname, UninstallString
            $uninst = $tobiivers.UninstallString
            $Outputbox.appendtext( "Removing - " + "$remove `r`n" )
            
            cmd /c $uninst "/quiet" "/norestart"
        }
    }
    Remove-Variable checkbox*
    $Outputbox.Appendtext( "Done!`r`n" )
}

#A6 Uninstall PCEye Package
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

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Dynavox Gaze Interaction Software") -or
        ($_.Displayname -Match "Tobii Dynavox PCEye Update Notifier") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze Selection Language Packs") -or
        ($_.Displayname -Match "Tobii IS3 Eye Tracker Driver") -or
        ($_.Displayname -Match "Tobii IS4 Eye Tracker Driver") -or
        ($_.Displayname -Match "Tobii Eye Tracker Browser") -or
        ($_.Displayname -Match "Tobii Dynavox PCEye Configuration Guide") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze HID") } | Select-Object Displayname, UninstallString

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
        "$ENV:ProgramData\TetServer",
        "$ENV:ProgramData\Tobii Dynavox\Gaze Interaction"
    )

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
		
		
    foreach ($key in $Key) {
        if (test-path $key) {
            $Outputbox.appendtext( "Removing - " + "$key`r`n" )
            Remove-Item $key -Force -ErrorAction ignore
        }
    }

    $Outputbox.appendtext( "Done!`r`n" )
}

#A7 Uninstall Communicator
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

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { $_.Displayname -match "Tobii Dynavox Communicator" } | Select-Object Publisher, Displayname, UninstallString

    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        $uninst = $ver.UninstallString
        & cmd /c $uninst /quiet /norestart
    }

    $paths = ( "$Env:USERPROFILE\AppData\Roaming\Tobii Dynavox\Communicator",
        "$ENV:ProgramData\Tobii Dynavox\Communicator" )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.AppendText( "Removing - " + "$path`r`n")
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }

    $Keys = (
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\MyTobii\MPA\VS Communicator 4",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4",
        "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5" )

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }

    $Outputbox.Appendtext( "Done!`r`n" )
}

#A8 Uninstalls only Compass
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

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Dynavox Compass") } | Select-Object Displayname, UninstallString

    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        $uninst = $ver.UninstallString
        & cmd /c $uninst /quiet /norestart
    }

    $Keys = ( "$ENV:ProgramData\Tobii Dynavox\Tobii.Licensing\Compass" )
    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }

    $Outputbox.appendtext( "Done!`r`n" )
}

#A9 Uninstall TGIS
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

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Dynavox Gaze Interaction Software") } | Select-Object Displayname, UninstallString
    ForEach ($ver in $TobiiVer) {
        $Uninstname = $ver.Displayname
        $uninst = $ver.UninstallString
        $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
        & cmd /c $uninst /quiet /norestart
    }

    $paths = (
        "$env:ProgramData\Tobii Dynavox\Gaze Interaction\",
        "$ENV:ProgramData\Tobii Dynavox\Gaze Selection\Word Prediction\Language Packs\")

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }

    $Outputbox.appendtext( "Done!`r`n" )
}

#A10 Function for the option "Remove TGIS calibration profiles #Tobii service is stopped
Function TGISProfilesremove {

    $answer1 = $wshell.Popup("This will remove ONLY calibrations for every profile, it will NOT remove the actual profiles. The Gaze Interaction software will close and tobii service will restart.`r`nContinue?", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.appendtext( "Shutting down TGIS software...`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $outputBox.appendtext( "Action canceled: Remove calibration profiles." )
    }	

    $Processkills = get-process "Tobii.Service", "TobiiEyeControlOptions", "TobiiEyeControlServer", "Notifier" | Stop-process -force -Passthru -erroraction ignore | Select-Object Processname |
    Format-table -Hidetableheaders | Out-string
    foreach ($Processkill in $Processkills) {
        if ($Processkill) {
            $Outputbox.Appendtext( "Stopping: " + "$Processkill`r`n" )
        }
    }
    
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
    try {
        Start-Service -Name "Tobii Service" -ErrorAction Stop
        Start-Sleep 1
        $Outputbox.Appendtext( "Tobii Service started! `r`n")
    }
    Catch {
        $Outputbox.Appendtext( "Tobii Service failed to start!`r`n" )
    }

    $outputbox.appendtext( "Done!`r`n" )
}

#A11
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
    $outputbox.appendtext("Done! `r`n")
}

#A12 Resets and restart TETC Configuration
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

    try {
        $Processkills = get-process  -Name '*Tobii.Service*', '*Tobii.EyeX.Engine*', '*Tobii.EyeX.Interaction*', '*Tobii.EyeX.Tray*' | Stop-process -force -Passthru -erroraction ignore | Select-Object Processname | Format-table -Hidetableheaders | Out-string
    }
    catch { 
        $OutputBox.AppendText( "No processes were found`r`n" )
    }
    foreach ($Processkill in $Processkills) {
        if ($Processkill) {
            $Outputbox.Appendtext( "Stopping: " + "$Processkill`r`n" )
        }
    }
 
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

#A13
Function BackupGazeInteraction {
    $outputBox.clear()
    $path = ( "C:\ProgramData\Tobii Dynavox\Old Gaze Interaction" )

    $outputbox.Appendtext( "Attempting to backup folder...`r`n" )
    if (Test-path $path) {
        $outputBox.appendtext( "Backup folder already exist in: C:\ProgramData\Tobii Dynavox\Old Gaze Interaction, please move it to another location or remove it before trying to backup again.`r`n" )
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

#A14 Copy licenses function. If any path to $Licensepaths exists, it will make a folder "Tobii Licenses", copy the licensefolders to the new folder(Does not contain the keys.xml, it is only the folder)
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
    #Filters the content to only get the string between the activationkey words
    #Creates txt files for licenses
    if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control\*") {
        $GetcontentWC = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Windows Control\keys.xml"
        $Outputbox.appendtext( "-- Window Control license copied.`r`n" )
        $LicenseWC = [regex]::Matches($getcontentWC, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()
        $LicenseWC | Out-file "C:\Tobii Licenses\Windows Control\Windows Control License.txt" -erroraction ignore
    }

    if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction\*") {
        $GetcontentTGIS = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Gaze Interaction\keys.xml"
        $Outputbox.appendtext( "-- Gaze Interaction license copied.`r`n" )
        $LicenseTGIS = [regex]::Matches($getcontentTGIS, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()
        $LicenseTGIS | Out-file "C:\Tobii Licenses\Gaze Interaction\Gaze Interaction License.txt" -erroraction ignore
    }

    if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5\*") {
        $GetcontentTC5 = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 5\keys.xml"
        $Outputbox.appendtext( "-- Communicator 5 license copied.`r`n" )
        $LicenseTC5 = [regex]::Matches($getcontentTC5, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()
        $LicenseTC5 | Out-file "C:\Tobii Licenses\Communicator 5\Communicator 5 License.txt" -erroraction ignore
    }

    if (test-path "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4\*") {
        $GetcontentTC4 = get-content "C:\ProgramData\Tobii Dynavox\Tobii.Licensing\Communicator 4\keys.xml"
        $Outputbox.appendtext( "-- Communicator 4 license copied.`r`n" )
        $LicenseTC4 = [regex]::Matches($getcontentTC4, '(?<=\<ActivationKey\>).+(?=\</ActivationKey\>)', "singleline").Value.trim()
        $LicenseTC4 | Out-file "C:\Tobii Licenses\Communicator 4\Communicator 4 License.txt" -erroraction ignore
    } #Add compass to the list.

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
    Get-ItemProperty | Where-Object { $_.Publisher -like '*Tobii*' } | Select-Object Displayname, Displayversion | Sort-Object Displayname | format-table -HideTableHeaders | out-string
    $Listwindowsapp = Get-AppxPackage | Where-Object { ($_.Publisher -like '*Tobii*') -or
        ($_.Name -like '*Snap*') } | Select-Object name | format-table -HideTableHeaders | out-string
    $outputBox.AppendText( "TOBII INSTALLED SOFTWARE:$Listapps`r`n" )
    $outputBox.AppendText( "TOBII WINDOWS STORE APPS:$Listwindowsapp`r`n" )
}

#B2 Lists currently active tobii processes & services
Function GetProcesses {
    $outputBox.clear()
    $outputBox.appendtext( "Listing active Tobii processes. (If empty - no processes were found) `r`n" )
    $GetProcess = get-process "*GazeSelection*", "*Tobii*" | Select-Object Processname | Format-table -hidetableheaders | Out-string
    $GetServices = Get-Service -Name '*Tobii*' | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
    if ($GetProcess) {
        $outputbox.appendtext("ACTIVE PROCESSES:$GetProcess`r`n")
    }
    else {
        $outputbox.appendtext("NO ACTIVE PROCESSES`r`n")
    }
    if ($GetServices) {
        $outputbox.appendtext("ACTIVE Services:$GetServices`r`n")
    }
    else {
        $outputbox.appendtext("NO ACTIVE Services`r`n")
    }
}

#B3
Function TTSWversions {
    $outputBox.clear()
    $Listapps = Get-ChildItem -Recurse -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty | Where-Object { 
        $_.Displayname -like '*Tobii Experience Software*' -or 
        $_.Displayname -like '*Tobii Device Drivers*' -or 
        $_.Displayname -like '*Tobii Eye Tracking For Windows*' 
    } | Select-Object Displayname, Displayversion | Sort-Object Displayname | format-table -HideTableHeaders   | out-string
    
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        try { 
            $erroractionpreference = "Stop"
            $Firmware = .\FWUpgrade32.exe --auto --info-only 
        }
        Catch [System.Management.Automation.RemoteException] {
            $outputbox.appendtext("No Eye Tracker Connected`r`n")
        }
    }
    else { 
        $outputbox.appendtext("File FWUpgrade32.exe is missing!`r`n" )
    }
    $testpath = "C:\Program Files\Tobii\Tobii EyeX"
    if (Test-path $testpath) { 
        Set-Location "C:\Program Files\Tobii\Tobii EyeX"
        $Components = Get-childitem * -include platform_runtime_IS5GIBBONGAZE_service.exe, InstallerPackageRemovalTool.exe, Tobii.Configuration.exe, Tobii.EyeX.Engine.exe, Tobii.EyeX.Interaction.exe, Tobii.Service.exe, tobii_stream_engine.dll  | foreach-object { "{0}`t{1}" -f $_.Name, [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion }
    }
    else {
        $outputbox.appendtext("File 1 not exist `r`n")
    }
    if ($Firmware -match "IS5_Gibbon_Gaze") {
        $PDKversion = Get-ChildItem -Path "C:\Program Files\Tobii\Tobii EyeX" -Recurse -file -include "platform_runtime_IS5GIBBONGAZE_service.exe" | foreach-object { "{0}`t{1}" -f $_.Name, [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion }

    }
    elseif ($Firmware -match "IS5_Large_PC_Eye_5") {
        $PDKversion = Get-ChildItem -Path "C:\Program Files\Tobii\Tobii EyeX" -Recurse -file -include "platform_runtime_IS5LARGEPCEYE5_service.exe" | foreach-object { "{0}`t{1}" -f $_.Name, [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion }
    }
    else {
        $outputbox.appendtext("File 2 not exist `r`n")
    }

    $outputbox.appendtext("$Listapps`r`n")
    foreach ($component in $Components) {
        $outputbox.appendtext("$component`r`n")
    }

    foreach ($pdkversions in $PDKversion) {
        $outputbox.appendtext("$pdkversions`r`n")
    }

    $ETVandModel = $Firmware | Select-String -Pattern "Firmware version", "Model"
    foreach ($NewETVandModel in $ETVandModel) {
        $outputbox.appendtext("$NewETVandModel`r`n")
    }
    $ETSN = $Firmware | Select-String -Pattern "tobii-ttp"
    $ETSN = "$ETSN"
    $NewETSN = $ETSN -replace "Automatically selected eye tracker", ""
    $outputbox.appendtext("Eye Tracker S/N: $NewETSN `r`n")

    $driver = Get-WmiObject Win32_PnPSignedDriver | Where-Object { ($_.DeviceName -match "Tobii Hello") -or ($_.DeviceName -match "Tobii Eye Tracker") } | Select-Object DeviceName, DriverVersion
    foreach ($Drivers in $driver) {
        $drivername = $Drivers.DeviceName
        $driverversion = $Drivers.DriverVersion
        $outputbox.appendtext("$drivername  $driverversion`r`n")
    }
    $outputbox.appendtext("Done! `r`n")
}
#B4
Function IS5PID {
    $outputBox.clear()
    $outputBox.appendtext( "Checking IS5 PID...`r`n" )
    $getdeviceid = $null
    $getdeviceid = Get-WmiObject Win32_USBControllerDevice | % { [wmi]($_.Dependent) } | Where-Object DeviceID -Like "*Tobii*" | Select-object DeviceID
    $outputbox.appendtext("$getdeviceid `r`n")								   
    $getdeviceid2 = Get-CimInstance Win32_PnPSignedDriver | Where-Object Description -Like "*WinUSB Device*" | Select-Object DeviceID
    Start-Sleep -s 5
    $outputbox.appendtext("$getdeviceid2`r`n")								   
    if (!$getdeviceid -or !$getdeviceid2) {
        $outputbox.appendtext("the tracker is not connected`r`n")
    }
    # gwmi Win32_USBControllerDevice |%{[wmi]($_.Dependent)} | Sort Manufacturer,Description,DeviceID | Ft -GroupBy Manufacturer Description,Service,DeviceID | out-file c:\VidPid.txt
    $outputbox.appendtext("Done!`r`n")
}

#B5
Function ListDrivers {
    $outputBox.clear()
    $outputBox.appendtext( "Listing all drivers in c:/tobii.txt and here...`r`n" )
    pnputil /enum-drivers >c:\tobii.txt
    $TobiiDrivers = Get-WindowsDriver -Online -All | Where-Object { $_.ProviderName -eq "Tobii AB" } | Select-Object Driver , OriginalFileName
    ForEach ($drivers in $TobiiDrivers) {
        $inf = $drivers.Driver 
        $List = $drivers.OriginalFileName
        $List = $List.Replace("C:\Windows\System32\DriverStore\FileRepository\", "")
        $outputbox.appendtext("$inf : $List `r`n")
    }
    $outputbox.appendtext("Done!`r`n")
}

#B6
Function ETfw {
    $outputBox.clear()
    $outputBox.appendtext( "Checking Eye tracker Firmware...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        try { 
            $erroractionpreference = "Stop"
            $Firmware = .\FWUpgrade32.exe --auto --info-only 
        }
        Catch [System.Management.Automation.RemoteException] {
            $outputbox.appendtext("No Eye Tracker Connected`r`n")
        }
        $outputbox.appendtext("$Firmware`r`n") 
    }
    else { 
        $outputbox.appendtext("File FWUpgrade32.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B7
function GetFrameworkVersionsAndHandleOperation() {
    $outputBox.clear()
    $installedFrameworks = @()
    if (IsKeyPresent "HKLM:\Software\Microsoft\.NETFramework\Policy\v1.0" "3705") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 1.0`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v1.1.4322" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 1.1`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v2.0.50727" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 2.0`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.0\Setup" "InstallSuccess") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 3.0`r`n") }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.5" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 3.5`r`n" ) }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 4.0c`r`n" ) }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Install") { $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 4.0`r`n" ) }   

    $result = -1
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client" "Install" -or IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Install") {
        # .net 4.0 is installed
        $result = 0
        $version = GetFrameworkValue "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Release"
        
        if ($version -ge 528040 -Or $version -ge 528372 -Or $version -ge 528049) {
            # .net 4.8
            $outputbox.appendtext( "Installed .Net Framework 4.8`r`n")
            $result = 10
        }
        elseif ($version -ge 461808 -Or $version -ge 461814) {
            # .net 4.7.2
            $outputbox.appendtext("Installed .Net Framework 4.7.2`r`n")
            $result = 9
        }
        elseif ($version -ge 461308 -Or $version -ge 461310) {
            # .net 4.7.1
            $outputbox.appendtext( "Installed .Net Framework 4.7.1`r`n")
            $result = 8
        }
        elseif ($version -ge 460798 -Or $version -ge 460805) {
            # .net 4.7
            $outputbox.appendtext( "Installed .Net Framework 4.7`r`n")
            $result = 7
        }
        elseif ($version -ge 394802 -Or $version -ge 394806) {
            # .net 4.6.2
            $outputbox.appendtext( "Installed .Net Framework 4.6.2`r`n")
            $result = 6
        }
        elseif ($version -ge 394254 -Or $version -ge 394271) {
            # .net 4.6.1
            $outputbox.appendtext( "Installed .Net Framework 4.6.1`r`n")
            $result = 5
        }
        elseif ($version -ge 393295 -Or $version -ge 393297) {
            # .net 4.6
            $outputbox.appendtext( "Installed .Net Framework 4.6`r`n")
            $result = 4
        }
        elseif ($version -ge 379893) {
            # .net 4.5.2
            $outputbox.appendtext( "Installed .Net Framework 4.5.2`r`n")
            $result = 3
        }
        elseif ($version -ge 378675) {
            # .net 4.5.1
            $outputbox.appendtext( "Installed .Net Framework 4.5.1`r`n")
            $result = 2
        }
        elseif ($version -ge 378389) {
            # .net 4.5
            $outputbox.appendtext( "Installed .Net Framework 4.5`r`n")
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
    if ($null -eq (Get-ItemProperty $path).$key) { return $false }
    #if ((Get-ItemProperty $path).$key -eq $null) { return $false }
    return $true
}
function GetFrameworkValue([string]$path, [string]$key) {
    if (!(Test-Path $path)) { return "-1" }
    return (Get-ItemProperty $path).$key  
}

#B8
Function BatteryLog {
    $outputBox.clear()
    $outputBox.appendtext( "Starting TobiiDynavox.QA.BatteryMonitor.exe...`r`n" )
    $fpath = (Get-ChildItem -Path "$PSScriptRoot" -Filter "TobiiDynavox.QA.BatteryMonitor.exe" -Recurse).FullName | Split-Path 
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        start-process .\TobiiDynavox.QA.BatteryMonitor.exe
        $outputbox.appendtext("Results will be saves in $fpath\battery_log.csv`r`n")
    }
    else { 
        $outputbox.appendtext("File BatteryMonitor.exe is missing!`r`n" )
    }	
    $outputbox.appendtext("Done! `r`n")
}

#B9
Function TrackStatus {
    $outputBox.clear()
    $outputBox.appendtext( "Showing EA Track Status...`r`n" )
    $testpath = "C:\Program Files (x86)\Tobii Dynavox\Eye Assist"
    if (!(Test-path $testpath)) {
        $outputbox.appendtext("EA may not been installed. Make sure that EA is installed and try again.`r`n")
    }
    else {
        Set-Location $testpath
        $value = Get-Process | Where-Object { $_.MainWindowTitle -like "track status" } | Select-Object MainWindowTitle
        if ($value) {
            .\TobiiDynavox.EyeAssist.Smorgasbord.exe --hidetrackstatus
        }
        elseif (!($value)) {
            .\TobiiDynavox.EyeAssist.Smorgasbord.exe --showtrackstatus
        }
    }
    $outputbox.appendtext("Done! `r`n")
}



#B10 LogCollector
Function LogCollector {
    $outputBox.clear()
    $outputbox.appendtext("Start `r`n")
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "LogCollectorTool_V0.8.ps1" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        .\LogCollectorTool_V0.8.ps1
    }
    else { 
        $outputbox.appendtext("File LogCollectorTool_V0.8.ps1 is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B11
Function WCF {
    $outputBox.clear()
    $outputBox.appendtext( "Checking WCF Endpoint Blocking Software...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "handle.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        Start-Process cmd "/c  `"handle.exe net.pipe & pause `""
    }
    else { 
        $outputbox.appendtext("File handle.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B12
Function BeforeUninstallGG {
    $outputBox.clear()
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        try { 
            $erroractionpreference = "Stop"
            $Firmware = .\FWUpgrade32.exe --auto --info-only 
        }
        catch [System.Management.Automation.RemoteException] {
            $outputbox.appendtext("PDK is not installed`r`n")
        }
    }
    else { 
        $outputbox.appendtext("File FWUpgrade32.exe is missing!`r`n" )
    }
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Device Drivers For Windows") } | Select-Object Displayname, DisplayVersion, UninstallString

    if ($Firmware -match "IS5_Gibbon_Gaze" -and $TobiiVer.DisplayVersion -eq "4.49.0.4000" ) { 
        $outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "BeforeUninstall.bat" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        if ($fpath.count -gt 0) {
            $Outputbox.Appendtext("Files found!`r`n" )
            Set-Location $fpath
            $Installer = cmd /c "BeforeUninstall.bat"
            $Outputbox.appendtext("$Installer`r`n")
        }
        else { 
            $outputbox.appendtext("File BeforeUninstall.bat is missing!`r`n" )
        }
        $Outputbox.appendtext( "Done! `r`n" )
    } 
    else { $outputbox.appendtext( "No need to run the script`r`n") }
}

#B13 DriverSetup in GG
Function DriverSetup {
    $outputBox.clear()
    $outputbox.appendtext("Start `r`n")
    $LogPath = "$ENV:USERPROFILE\AppData\Local\Temp"

    $ErrorPath = "$LogPath\ErrorLogs"
    if (!(Test-Path "$ErrorPath")) {
        $outputbox.appendtext( "Creating ErrorLogs folder.. `r`n")
        New-Item -Path "$ErrorPath" -ItemType Directory   
    }
    if (!(Test-Path "$ErrorPath\InstallerError.txt") -or !(Test-Path "$ErrorPath\InstallerError2.txt") -or !(Test-Path "$ErrorPath\InstallerError3.txt")) {
        $ErrorFile = New-Item -Path $ErrorPath -Name "InstallerError.txt" -ItemType "file"
        $ErrorFile2 = New-Item -Path $ErrorPath -Name "InstallerError2.txt" -ItemType "file"
        $ErrorFile3 = New-Item -Path $ErrorPath -Name "InstallerError3.txt" -ItemType "file"
        $outputbox.appendtext( "creating file `r`n")
    }
    else {
        Clear-Content -Path "$ErrorPath\InstallerError.txt"
        Clear-Content -Path "$ErrorPath\InstallerError2.txt"
        Clear-Content -Path "$ErrorPath\InstallerError3.txt"
        $outputbox.appendtext( "cleaing `r`n")
    }
    Set-Location $LogPath
    $Installercontent = Get-ChildItem "tobii*.log" -Recurse -File | Sort-Object name -desc | Select-Object -expand Fullname
    foreach ($NewInstallercontent in $Installercontent) {
        New-Item -Path $ErrorPath -Name "temp.txt" -ItemType "file"
        Get-Content -Path "$NewInstallercontent" -Raw | ForEach-Object -Process { $_ -replace "- `r`n", '- ' } | Add-Content -Path "$ErrorPath\temp.txt"
        $string = "Executing\s+op\:\s+CustomActionSchedule\(Action\=DisconnectDevices,ActionType\=3073,Source\=BinaryData,Target\=WixQuietExec,CustomActionData\="
        $content9 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "$string" -AllMatches | ForEach-Object -Process { $_ -replace ".*CustomActionData=" -replace "-inf.*" } | ForEach-Object -Process { $_ -replace ("`"", "") }
        add-Content "$ErrorPath\InstallerError.txt" -value $content9, "`n"
        Remove-Item "$ErrorPath\temp.txt"
    }
	
    (Get-Content "$ErrorPath\InstallerError.txt") | Where-Object { $_.trim() -ne "" } | set-content "$ErrorPath\InstallerError2.txt"
    $tester4 = (Get-Content "$ErrorPath\InstallerError2.txt")
    if ($null -eq $tester4) {
        $outputbox.appendtext("File is Empty `r`n")
    }
    else {
        
        foreach ($line in $tester4) {
            $array = $line.split("\")
            $path = [string]::Join("\", $array[0..($array.length - 2)]) 
            Add-Content -Path "$ErrorPath\InstallerError3.txt" -Value $path
        }
    }
    #Remove-Item "$ErrorPath\InstallerError.txt"

    if ($Null -eq (Get-Content "$ErrorPath\InstallerError3.txt")) {
        $OutputBox.AppendText( "File is empty, no need to copy driver setup`r`n")
    }
    else {
        $test4 = Get-Content -Path "$ErrorPath\InstallerError3.txt"

        foreach ($tests4 in $test4) {
            $OutputBox.AppendText("Copy DriverSetup to specific path`r`n")
            New-Item -ItemType Directory -Force -Path $tests4
            $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "DriverSetup.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
            #if($fpath.count -gt 0) {
            #$Outputbox.Appendtext("Files found!`r`n" )
            Set-Location $fpath
            Copy-Item -Path ("DriverSetup.exe") -Destination $tests4
            #    } else { 
            #$outputbox.appendtext("File DriverSetup.exe is missing!`r`n" )
            #}
        }


    }
    $outputbox.appendtext("Done! `r`n")
}

#B14 uninstall_gibbon
Function uninstall_gibbon {
    $outputBox.clear()
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        try { 
            $erroractionpreference = "Stop"
            $Firmware = .\FWUpgrade32.exe --auto --info-only 
        }
        catch [System.Management.Automation.RemoteException] {
            $outputbox.appendtext("PDK is not installed`r`n")
        }
    }
    else { 
        $outputbox.appendtext("File FWUpgrade32.exe is missing!`r`n" )
    }
    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty | Where-Object { ($_.Displayname -Match "Tobii Device Drivers For Windows") } | Select-Object Displayname, DisplayVersion, UninstallString

    if ($Firmware -match "IS5_Gibbon_Gaze" -and $TobiiVer.DisplayVersion -eq "4.49.0.4000" ) { 
        $outputBox.clear()
        $outputbox.appendtext("Start `r`n")
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "uninstall_gibbon.ps1" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        if ($fpath.count -gt 0) {
            $Outputbox.Appendtext("Files found!`r`n" )
            Set-Location $fpath
            .\uninstall_gibbon.ps1
        }
        else { 
            $outputbox.appendtext("File uninstall_gibbon.ps1 is missing!`r`n" )
        }
        $outputbox.appendtext("Done! `r`n")
    } 
    else { $outputbox.appendtext( "No need to run the script`r`n") }
}

#B15 SetDebugLogging
Function SetDebugLogging {
    #Instructions
    #Start PowerShell as admin
    #Run "Set-ExecutionPolicy Unrestricted"
    #Run the attached script.
    #Tobii Service is then restarted and new log levels are set. If you are experiencing issues, please restart Tobii Service.
    #https://confluence.tobii.intra/pages/viewpage.action?spaceKey=EYEX&title=Changing+log+level
    
    #$outputBox.clear()
    param (    
        [switch]$reset,
        [ValidateSet( 'DEBUG', 'INFO', 'WARNING', 'ERROR', 'FATAL')]
        [string]$MinLevel = 'DEBUG',
        [ValidateSet( 'DEBUG', 'INFO', 'WARNING', 'ERROR', 'FATAL')] 
        [string]$MaxLevel = 'FATAL')

    if ($reset) {
        $MinLevel = 'INFO'
        $MaxLevel = 'ERROR'  
        $OutputBox.AppendText( "'Reset log levels to  + $MinLevel +  and  + $MaxLevel`r`n" )
        
    }
    else {
        $OutputBox.AppendText( "Set log levels to  + $MinLevel +  and  + $MaxLevel`r`n")
    }   
    
    $tobiiInstallPath = "C:\Program Files\Tobii\Tobii EyeX\"
    $configAppConfig = [IO.Path]::Combine($tobiiInstallPath, 'Tobii.Configuration.exe.config')
    $interactionAppConfig = [IO.Path]::Combine($tobiiInstallPath, 'Tobii.EyeX.Interaction.exe.config')
    $EngineAppConfig = [IO.Path]::Combine($tobiiInstallPath, 'Tobii.EyeX.Engine.exe.config')
    $ServiceAppConfig = [IO.Path]::Combine($tobiiInstallPath, 'Tobii.Service.exe.config')

    $PathToConfigFiles = $configAppConfig, $interactionAppConfig, $EngineAppConfig, $ServiceAppConfig

    foreach ($configFilPath in $PathToConfigFiles) {
        $appConfig = New-Object XML
        # load the config file as an xml object
        $appConfig.Load($configFilPath)
        $OutputBox.AppendText( "Updating config file  + $configFilPath`r`n")

        $minLevelNode = $appConfig.SelectSingleNode("//*[@name='LevelMin']")
        if ($minLevelNode.Value -ne $MinLevel) {
            # 'Change: ' + $minLevelNode.name + ' from: ' + $minLevelNode.Value + ' to: ' + $MinLevel 
            $minLevelNode.Value = $MinLevel           
        }
        else {
            $OutputBox.AppendText( "Required min level is already set, skip..`r`n")

        }
    
        $maxLevelNode = $appConfig.SelectSingleNode("//*[@name='LevelMax']")
        if ($maxLevelNode.Value -ne $MaxLevel) {
            # 'Change: ' + $maxLevelNode.name + ' from: ' + $maxLevelNode.Value + ' to: ' + $MaxLevel 
            $maxLevelNode.Value = $MaxLevel           
        }
        else {
            $OutputBox.AppendText( "Required max level is already set, skip..`r`n")
        }
        
        # save the updated config file
        $appConfig.Save($configFilPath)
    }
    $OutputBox.AppendText( "All config files are updated - Done!`r`n")


}

#B16 
Function ModelCheck {
    $outputBox.clear()
    $RegPath1 = Test-Path -Path HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation\
    $RegPath2 = Test-Path -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation\
    $RegPath3 = Test-Path -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\OEMInformation\

    if (Test-Path $RegPath1) {
        $TobiiVer1 = Get-ItemProperty  -Path  HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation\
    }
    if (Test-Path $RegPath2) {
        $TobiiVer2 = Get-ItemProperty  -Path  HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation\
    }
    if (Test-Path $RegPath3) {
        $TobiiVer3 = Get-ItemProperty  -Path  HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\OEMInformation\
    }

 
    if ($TobiiVer1.EyeTrackerModel -eq "PCEye5") {
        #Checked
        $OutputBox.AppendText("Eye tracker is IS514`r`n")
    }
    elseif ($TobiiVer1.ProductType -eq "TDG16") {
        #Checked
        $OutputBox.AppendText("Eye tracker is TDG16`r`n")
    }
    elseif ($TobiiVer1.ProductType -eq "TDH10") {
        #
        $OutputBox.AppendText("Eye tracker is TDH10`r`n")
    }
    elseif ($TobiiVer1.ProductType -eq "TDTW7") {
        #
        $OutputBox.AppendText("Eye tracker is TDTW7`r`n")
    }
    elseif ($TobiiVer1.ProductType -eq "TDG10") {
        #
        $OutputBox.AppendText("Eye tracker is TDG10`r`n")
    }
    elseif ($TobiiVer1.ProductType -eq "I-Series" -and $TobiiVer1.ProductModel -eq "I-12+") {
        #
        $OutputBox.AppendText("Eye tracker is TDI12-xxxxx`r`n")
    }
    elseif ($TobiiVer1.EyeTrackerModel -eq "EM12") {
        #
        $OutputBox.AppendText("Eye tracker is TEM12`r`n")
    }
    elseif ($TobiiVer3.EyeTrackerModel -eq "PCEye2") {
        #
        $OutputBox.AppendText("Eye tracker is PCEGO or PCEye Mini`r`n")
    }
    elseif ($TobiiVer3.EyeTrackerModel -eq "PCEyeExplore") {
        #
        $OutputBox.AppendText("Eye tracker is PCEyeExplore`r`n")
    }
    else {
        $OutputBox.AppendText( "No match," + $TobiiVer1.EyeTrackerModel + "and " + $TobiiVer.ProductType + "`r`n") 

    }
    #$TobiiVer1
    #$TobiiVer2
    #$TobiiVer3
}

#B17 
Function HWInfo {
    $outputBox.clear()
    #Creating folder
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "SupportTool v1.6.4.ps1" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
    }
    else { 
        $outputbox.appendtext("File SupportTool v1.6.4.ps1 is missing!`r`n" )
    }
    $infofolder = "$fpath\infofolder"
    if (!(Test-Path "$infofolder")) {
        $OutputBox.AppendText( "Creating folder..`r`n")
        New-Item -Path "$infofolder" -ItemType Directory  
    }
    else {
        Remove-Item -Path $infofolder\* -Recurse 
        $OutputBox.AppendText( "Folder is already created. Removing files`r`n")
    }

    #Creating files
    if (!(Test-Path "$infofolder\Monitors.txt") -or 
        !(Test-Path "$infofolder\hidDevices.txt") -or 
        !(Test-Path "$infofolder\motherboard.txt") -or 
        !(Test-Path "$infofolder\operatingSystem.txt") -or 
        !(Test-Path "$infofolder\pnpDevices.txt") -or  
        !(Test-Path "$infofolder\USBDeviceTree.txt") -or 
        !(Test-Path "$infofolder\PersistedData.txt") -or 
        !(Test-Path "$infofolder\ETInfo.txt") -or
        !(Test-Path "$infofolder\DeviceInfo.txt")
    ) {
        New-Item -Path $infofolder -Name "Monitors.txt" -ItemType "file"
        New-Item -Path $infofolder -Name "hidDevices.txt" -ItemType "file"
        New-Item -Path $infofolder -Name "motherboard.txt" -ItemType "file"
        New-Item -Path $infofolder -Name "operatingSystem.txt" -ItemType "file"
        New-Item -Path $infofolder -Name "pnpDevices.txt" -ItemType "file"
        New-Item -Path $infofolder -Name "USBDeviceTree.txt" -ItemType "file"
        New-Item -Path $infofolder -Name "PersistedData.txt" -ItemType "file"
        New-Item -Path $infofolder -Name "ETInfo.txt" -ItemType "file"
        New-Item -Path $infofolder -Name "DeviceInfo.txt" -ItemType "file"
        $OutputBox.AppendText( "creating info files`r`n")
    }
    else {
        Clear-Content -Path "$infofolder\Monitors.txt"
        Clear-Content -Path "$infofolder\hidDevices.txt"
        Clear-Content -Path "$infofolder\motherboard.txt"
        Clear-Content -Path "$infofolder\operatingSystem.txt"
        Clear-Content -Path "$infofolder\pnpDevices.txt"
        Clear-Content -Path "$infofolder\USBDeviceTree.txt"
        Clear-Content -Path "$infofolder\PersistedData.txt"
        Clear-Content -Path "$infofolder\ETInfo.txt"
        $OutputBox.AppendText( "cleaing files`r`n")
    }

    $DesktopMonitors = Get-CimInstance -ClassName Win32_DesktopMonitor -Property *
    $hidDevices = Get-WmiObject Win32_PnPSignedDriver | Where-Object devicename -Like "*tobii*" | Select-Object devicename, driverversion
    $motherboard = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -Property Mainboard, AdminPasswordStatus, AutomaticManagedPagefile, AutomaticResetBootOption, AutomaticResetCapability, BootOptionOnLimit, BootOptionOnWatchDog, BootROMSupported, BootStatus, BootupState, Caption, ChassisBootupState, ChassisSKUNumber, CreationClassName, CurrentTimeZone, DaylightInEffect, Description, DNSHostName, Domain, DomainRole, EnableDaylightSavingsTime, FrontPanelResetStatus, HypervisorPresent, InfraredSupported, InitialLoadInfo, InstallDate, KeyboardPasswordStatus, LastLoadInfo, Manufacturer, Model, Name, NameFormat, NetworkServerModeEnabled, NumberOfLogicalProcessors, NumberOfProcessors, OEMLogoBitmap, OEMStringArray, PartOfDomain, PauseAfterReset, PCSystemType, PCSystemTypeEx, PowerManagementCapabilities, PowerManagementSupported, PowerOnPasswordStatus, PowerState, PowerSupplyState, PrimaryOwnerContact, PrimaryOwnerName, ResetCapability, ResetCount, ResetLimit, Roles, Status, SupportContactDescription, SystemFamily, SystemSKUNumber, SystemStartupDelay, SystemStartupOptions, SystemStartupSetting, SystemType, ThermalState, TotalPhysicalMemory, UserName, WakeUpType, Workgroup
    $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -Property 'BootDevice', 'BuildNumber', 'BuildType', 'Caption', 'CodeSet', 'CountryCode', 'CreationClassName', 'CSCreationClassName', 'CSDVersion', 'CSName', 'CurrentTimeZone', 'DataExecutionPrevention_32BitApplications', 'DataExecutionPrevention_Available', 'DataExecutionPrevention_Drivers', 'DataExecutionPrevention_SupportPolicy', 'Debug', 'Description', 'Distributed', 'EncryptionLevel', 'ForegroundApplicationBoost', 'FreePhysicalMemory', 'FreeSpaceInPagingFiles', 'FreeVirtualMemory', 'InstallDate', 'LastBootUpTime', 'LocalDateTime', 'Locale', 'Manufacturer', 'MaxNumberOfProcesses', 'MaxProcessMemorySize', 'MUILanguages', 'Name', 'NumberOfLicensedUsers', 'NumberOfProcesses', 'NumberOfUsers', 'OperatingSystemSKU', 'Organization', 'OSArchitecture', 'OSLanguage', 'OSProductSuite', 'OSType', 'OtherTypeDescription', 'PAEEnabled', 'PlusProductID', 'PlusVersionNumber', 'PortableOperatingSystem', 'Primary', 'ProductType', 'RegisteredUser', 'SerialNumber', 'ServicePackMajorVersion', 'ServicePackMinorVersion', 'SizeStoredInPagingFiles', 'Status', 'SuiteMask', 'SystemDevice', 'SystemDirectory', 'SystemDrive', 'TotalSwapSpaceSize', 'TotalVirtualMemorySize', 'TotalVisibleMemorySize', 'Version', 'WindowsDirectory'
    $pnpDevices = Get-WmiObject Win32_PNPEntity
    $usbControllers = Get-WmiObject Win32_USBHub
    $USBDeviceTree1 = Get-CimInstance -ClassName Win32_USBHub -Property * 
    $USBDeviceTree2 = Get-CimInstance -ClassName Win32_USBControllerDevice
    $Monitor1 = Get-PnpDevice | Where-Object Class -Match "Monitor"
    $Monitor2 = Get-WmiObject WmiMonitorID -Namespace root\wmi
    $Display = Get-WmiObject -Namespace root\wmi -Class WmiMonitorBasicDisplayParams | Select-Object @{ N = "Computer"; E = { $_.__SERVER } }, InstanceName, @{N = "Horizonal"; E = { [System.Math]::Round(($_.MaxHorizontalImageSize) * 10, 2) } }, @{N = "Vertical"; E = { [System.Math]::Round(($_.MaxVerticalImageSize) * 10, 2) } }, @{N = "Size"; E = { [System.Math]::Round(([System.Math]::Sqrt([System.Math]::Pow($_.MaxHorizontalImageSize, 2) + [System.Math]::Pow($_.MaxVerticalImageSize, 2))), 2) } }, @{N = "Ratio"; E = { [System.Math]::Round(($_.MaxHorizontalImageSize) / ($_.MaxVerticalImageSize), 2) } }
    $PersistedData1 = Get-ChildItem -Path Registry::HKEY_CURRENT_USER\SOFTWARE\Tobii -Recurse
    $PersistedData2 = Get-ChildItem -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Tobii -Recurse
    
    Add-Content -path "$infofolder\Monitors.txt" -Value $DesktopMonitors
    Add-Content -path "$infofolder\hidDevices.txt" -Value $hidDevices
    Add-Content -path "$infofolder\motherboard.txt" -Value $motherboard
    Add-Content -path "$infofolder\operatingSystem.txt" -Value $operatingSystem
    Add-Content -path "$infofolder\pnpDevices.txt" -Value $pnpDevices
    Add-Content -path "$infofolder\USBDeviceTree.txt" -Value $usbControllers
    Add-Content -path "$infofolder\USBDeviceTree.txt" -Value $USBDeviceTree1
    Add-Content -path "$infofolder\USBDeviceTree.txt" -Value $USBDeviceTree2
    Add-Content -path "$infofolder\Monitors.txt" -Value $Monitor1
    Add-Content -path "$infofolder\Monitors.txt" -Value $Monitor2
    Add-Content -path "$infofolder\Monitors.txt" -Value $Display
    Add-Content -path "$infofolder\PersistedData.txt" -Value $PersistedData1
    Add-Content -path "$infofolder\PersistedData.txt" -Value $PersistedData2

    Get-Service -Name '*TobiiIS*' | Stop-Service -Force -passthru -ErrorAction ignore
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "CastorUsbCli.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if (test-path $fpath) {
        Set-Location $fpath

        $info = .\CastorUsbCli.exe "--info"
        $status = .\CastorUsbCli.exe "--status"
        $unitinfo = .\CastorUsbCli.exe "--unit-info"
        $flashinfo = .\CastorUsbCli.exe "--flash-info"
        $list = .\CastorUsbCli.exe "--list"
        $properties = .\CastorUsbCli.exe "--properties"
        $platform = .\CastorUsbCli.exe "--platform"
        $reset = .\CastorUsbCli.exe "--reset"
        $execute = .\CastorUsbCli.exe "--execute"
        $readbootheader = .\CastorUsbCli.exe "--read-boot-header"
        $readappheader = .\CastorUsbCli.exe "--read-app-header"
        $showscreenplane = .\CastorUsbCli.exe "--show-screen-plane"

        Get-Service -Name '*TobiiIS*' | Stop-Service -Force -passthru -ErrorAction ignore

        Add-Content -path "$infofolder\ETInfo.txt" -Value $info
        Add-Content -path "$infofolder\ETInfo.txt" -Value $status
        Add-Content -path "$infofolder\ETInfo.txt" -Value $unitinfo
        Add-Content -path "$infofolder\ETInfo.txt" -Value $flashinfo
        Add-Content -path "$infofolder\ETInfo.txt" -Value $list
        Add-Content -path "$infofolder\ETInfo.txt" -Value $properties
        Add-Content -path "$infofolder\ETInfo.txt" -Value $platform
        Add-Content -path "$infofolder\ETInfo.txt" -Value $reset
        Add-Content -path "$infofolder\ETInfo.txt" -Value $execute
        Add-Content -path "$infofolder\ETInfo.txt" -Value $readbootheader
        Add-Content -path "$infofolder\ETInfo.txt" -Value $readappheader
        Add-Content -path "$infofolder\ETInfo.txt" -Value $showscreenplane

        Get-Service -Name '*TobiiIS*' | start-Service  -passthru -ErrorAction ignore
    }
    else {
        $outputbox.appendtext("Not able to run ET info since it missing exe file")

    }
    $outputbox.appendtext("Reading battery info!`r`n")
    $key = 'HKLM:\SOFTWARE\WOW6432Node\Tobii Dynavox\Device'
    #$fpath = Get-ChildItem -Path $PSScriptRoot -Filter "batteryreport.ps1" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    #Set-Location $fpath

    if (Test-Path $key) {
        $SerialNumber = (Get-ItemProperty -Path $key)."Serial Number" 
        Add-Content -path "$infofolder\DeviceInfo.txt" -Value "Device's Serial Number is $SerialNumber"
        $outputbox.appendtext("Device Serial Number is $SerialNumber`r`n")
    
        $OEMImage = (Get-ItemProperty -Path $key)."OEM Image" 
        Add-Content -path "$infofolder\DeviceInfo.txt" -Value "Device's OEM Image is $OEMImage"
        $outputbox.appendtext("Device OEM Image is $OEMImage`r`n")

        $ProductKey = (Get-ItemProperty -Path $key)."Product Key"
        Add-Content -path "$infofolder\DeviceInfo.txt" -Value "Device's Product Key is $ProductKey"
        $outputbox.appendtext("Device Product Key is $ProductKey`r`n")
    }
    else {
        $SerialNumber = (Get-CimInstance -ClassName Win32_bios).SerialNumber
        $Model = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
        Add-Content -path "$infofolder\DeviceInfo.txt" -Value "This device is not TD device"
        Add-Content -path "$infofolder\DeviceInfo.txt" -Value "Device's Serial Number is $SerialNumber"
        Add-Content -path "$infofolder\DeviceInfo.txt" -Value "Device's Model is $Model"
    }

    if ($SerialNumber -match "TD110-") {
        $outputbox.appendtext("Battery report is not support on this device, runt I-110MLK.bat to get the report.`r`n")
    }
    else {
        powercfg /batteryreport /output "$infofolder\$SerialNumber-battery-report.html"
    }

    $DesignedCapacity = (Get-WmiObject -Class BatteryStaticData -Namespace ROOT\WMI).DesignedCapacity / 1000
    Add-Content -path "$infofolder\DeviceInfo.txt" -Value "Battery Designed Capacity is $DesignedCapacity mWh"
    $outputbox.appendtext("Design Capacity is $DesignedCapacity mWh`r`n")

    $FullChargedCapacity = (Get-WmiObject -Class BatteryFullChargedCapacity -Namespace ROOT\WMI).FullChargedCapacity / 1000
    Add-Content -path "$infofolder\DeviceInfo.txt" -Value "Battery Full Charged Capacity is $FullChargedCapacity mWh"
    $outputbox.appendtext("Full Charge Capacity is $FullChargedCapacity mWh`r`n")

    #$BatteryHealth = ($FullChargedCapacity/$DesignedCapacity)
    $BatteryHealth = [Math]::Round($FullChargedCapacity / $DesignedCapacity * 100)
    Add-Content -path "$infofolder\DeviceInfo.txt" -Value "Battery Health is $BatteryHealth %`r`n"
    $outputbox.appendtext("Battery Health is $BatteryHealth %`r`n")

    $outputbox.appendtext("Logs saved in $infofolder `r`nDone!`r`n")
}

#B18
Function InstallPDK {
    $outputBox.clear()
    $serviceName = "TobiiIS5GIBBONGAZE"
    if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
        $outputbox.appendtext("$serviceName Service is already installed`r`n")
    }
    else {
        $outputbox.appendtext("Installing PDK on I-Series`r`n")
        sc.exe create $serviceName binpath="C:\Windows\System32\DriverStore\FileRepository\is5gibbongaze.inf_amd64_07ff964b2ca8d0e4\platform_runtime_IS5GIBBONGAZE_service.exe" DisplayName= "Tobii Runtime Service" start= auto
        Start-Service -Name $serviceName -ErrorAction stop
    }
    $outputbox.appendtext("Done!`r`n")
}

#B19 Stops all currently active tobii processes
Function RestartProcesses {
    $outputBox.clear()
    $Outputbox.Appendtext( "Restart Services...`r`n")
    $StopServices = Get-Service -Name '*Tobii*' | Stop-Service -force -Passthru -erroraction ignore | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
    $Outputbox.Appendtext( "Stopping following Services:$StopServices`r`n")

    Start-Sleep -s 3
    $Processkill = get-process "GazeSelection" , "*TobiiDynavox*", "*Tobii.EyeX*", "Notifier" | Stop-process -force -Passthru -erroraction ignore | Select-Object Processname | Format-table -Hidetableheaders | Out-string
    $Outputbox.Appendtext( "Stopping following processes:$Processkill`r`n")

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
    $outputBox.Appendtext( "Done!`r`n" )
}

#B20
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

#B21
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
    $outputbox.appendtext("Done!`r`n")
}

#B22
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

    $outputbox.appendtext("Pinging ET..`r`n")
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        try { 
            $erroractionpreference = "Stop"
            $Firmware = .\FWUpgrade32.exe --auto --info-only 
            $outputBox.appendtext( "$Firmware`r`n" )
        }
        Catch [System.Management.Automation.RemoteException] {
            $outputbox.appendtext("No eye tracker could be found`r`n")
        }
    }
    else { 
        $outputbox.appendtext("File FWUpgrade32.exe is missing!`r`n" )
    }
    Try {
        $outputBox.appendtext( "Stopping all Tobii Services...`r`n" )
        Get-WMIObject Win32_Service -Filter "Name LIKE '%Tobii Service%' " | Select-Object -ExpandProperty Name | Stop-Service -Force -passthru -ErrorAction ignore
        Get-WMIObject Win32_Service -Filter "Name LIKE '%TobiiIS5LARGEPCEYE5%' " | Select-Object -ExpandProperty Name | Stop-Service -Force -passthru -ErrorAction ignore
        Get-WMIObject Win32_Service -Filter "Name LIKE '%TobiiGeneric%' " | Select-Object -ExpandProperty Name | Stop-Service -Force -passthru -ErrorAction ignore
    }
    Catch [System.Management.Automation.RemoteException] {
        $outputbox.appendtext("No Service Installed`r`n")
    }
    Try {
        $outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "CastorUsbCli.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        if ($fpath.count -gt 0) {
            $Outputbox.Appendtext("Files found!`r`n" )
            Set-Location $fpath
            .\CastorUsbCli.exe --reset BOOT
        }
        else { 
            $outputbox.appendtext("File CastorUsbCli.exe is missing!`r`n" )
        }
        $getPID = Get-WmiObject Win32_PnPSignedDriver | Where-Object devicename -Like "*WinUSB Device*" | Select-Object DeviceID
        if ($getPID) {
            $outputbox.appendtext("The reset is done. ET PID is now:$getPID`r`n")
        }
        else {
            $outputbox.appendtext("Not able to read PID`r`n")
        }
        $outputbox.appendtext("Starting all services now..`r`n")
        Get-WMIObject Win32_Service -Filter "Name LIKE '%Tobii Service%'" | Select-Object -ExpandProperty Name | Start-Service -ErrorAction Stop
        Get-WMIObject Win32_Service -Filter "Name LIKE '%TobiiIS5LARGEPCEYE5%'" | Select-Object -ExpandProperty Name | Start-Service -ErrorAction Stop
        Get-WMIObject Win32_Service -Filter "Name LIKE '%TobiiGeneric%'" | Select-Object -ExpandProperty Name | Start-Service -ErrorAction Stop

    }
    Catch [System.Management.Automation.RemoteException] {
        $outputbox.appendtext("No Eye Tracker Connected`r`n")
    }


    $outputbox.appendtext("Done! `r`n")
}

#B23
Function FWUpgrade {
    $outputBox.clear()
    $outputBox.appendtext( "Upgrade IS4 ET FW...`r`n" )
    $path = "C:\Program Files (x86)\Tobii\Service"
    if (Test-Path $path) {
        Set-Location -path $path
        $ETInfo = .\FWUpgrade32.exe --auto --info-only
        $outputbox.appendtext("Connected ET is: $ETInfo`r`n")
    }

    else {
        $outputbox.appendtext("No Eye Tracker Connected or FW upgrade not supported`r`n")
    }
    if ($ETInfo -match "PCE1M") {
        #PCEye Mini: tobii-ttp://PCE1M-010106010685
        $outputbox.appendtext("Upgrading PCEye mini FW..`r`n")
        $PCEyeMini = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4pceyemini_firmware_2.27.0-4014648.tobiipkg" --no-version-check
        $outputbox.appendtext("$PCEyeMini`r`n")
        $outputbox.appendtext("Done! `r`n")
    }
    elseif ($ETInfo -match "IS4_Large_102") {
        $outputbox.appendtext("Upgrading PCEye Plus FW..`r`n")
        $PCEyePlus = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4large102_firmware_2.27.0-4014648.tobiipkg" --no-version-check
        $outputbox.appendtext("$PCEyePlus`r`n")
        $outputbox.appendtext("Done! `r`n")
    }
    elseif ($ETInfo -match "IS4_Large_Peripheral") {
        $outputbox.appendtext("Upgrading 4C FW..`r`n")
        $4C = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4largetobiiperipheral_firmware_2.27.0-4014648.tobiipkg" --no-version-check
        $outputbox.appendtext("$4C`r`n")
        $outputbox.appendtext("Done`r`n")
    }
    elseif ($ETInfo -match "IS4_Base_I-series") {
        $outputbox.appendtext("Upgrading I-Series+ FW..`r`n")
        $ISeries = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii Dynavox\Gaze Interaction\Eye Tracker Firmware Releases\IS4B1\is4iseriesb_firmware_2.9.0.tobiipkg" --no-version-check
        $outputbox.appendtext("$ISeries`r`n")
        $outputbox.appendtext("Done. Restart ET through Control Center `r`n")
    }
    elseif ($ETInfo -match "tet-tcp") {
        #Tobii Firmware Upgrade Tool Automatically selected eye tracker tet-tcp://172.28.195.1 Failed to open file
        $outputbox.appendtext("ET model is IS20. Use ET Browser to upgrade. Make sure that Bonjure is installed.`r`n")
    }
    else {
        $outputbox.appendtext("No ET connected or ET not supported`r`n")
    }
}

#B24 virtual window_EXPMISC-321
Function VirtWin {
    $outputBox.clear()
    $outputbox.appendtext("Start `r`n")
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "VirtualWindowCoreSDK.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        Start-Process cmd "/c `"VirtualWindowCoreSDK.exe`"" 
    }
    else { 
        $outputbox.appendtext("File VirtualWindowCoreSDK.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B25
Function ETConnection {
    $outputBox.clear()
    $outputBox.appendtext( "Running ET connection check...`r`n" )
    $outputBox.appendtext( "Results of output will be stored in C:\Output.txt...`r`n" )
    $a = 1
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'Loop'
    $msg = 'Enter number of loops:'
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    Do {
        Start-sleep -s 1
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        if ($fpath.count -gt 0) {
            $Outputbox.Appendtext("Files found!`r`n" )
            Set-Location $fpath
            try { 
                $erroractionpreference = "Stop"
                $getinfo = cmd /c "FWUpgrade32.exe" --auto --info-only | out-string
            }
            catch [System.Management.Automation.RemoteException] {
                $outputbox.appendtext("No Eye Tracker Connected`r`n")
            }
        }
        else { 
            $outputbox.appendtext("File FWUpgrade32.exe is missing!`r`n" )
        }

        $time = Get-Date -UFormat %H:%M:%S
        Add-content C:\Output.txt $time, $getinfo
        $a
        $outputbox.appendtext("$getinfo`r`n")
        $a++
    } while ($a -le $b)
    $outputbox.appendtext("Done! `r`n")
}

#B26
Function HBTool {
    $outputBox.clear()
    $outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
    Start-Process -FilePath C:\Users\aes\Desktop\script\BeforeUninstall.bat
    $outputbox.appendtext("ACTIVE PROCESSES: `r`n")
    $outputbox.appendtext("Done! `r`n")
}

#B27
Function EAProfileCreation {
    $outputBox.clear()
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'SMBios tool'
    $msg = "Press`r`n1 to create profile based on default `r`n2 to create as many profiles and calibrate,`r`n3 to remove all created profiles"
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    
    if ($b -match "1") { 
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
        $title = 'Profile Name'
        $msg = "Write a profile name"
        $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
        Set-Location 'C:\Program Files (x86)\Tobii Dynavox\Eye Assist'
        .\TobiiDynavox.EyeAssist.Smorgasbord.exe --createprofilewithdefaultcalibration --profile $b
        $outputbox.appendtext("Profile with $b has been created`r`n")
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
            $NewProfile = .\TobiiDynavox.EyeAssist.Smorgasbord.exe --startcreateprofileandcalibrate --profile $a
            $outputbox.appendtext("`r`nCreating profile with name: $a`r`n")
            Start-sleep -s 10
            $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
            if ($fpath.count -gt 0) {
                $Outputbox.Appendtext("Files found!`r`n" )
                #Set-Location $fpath
                try { 
                    $erroractionpreference = "Stop"
                    $getinfo = cmd /c "$fpath\FWUpgrade32.exe" --auto --info-only | out-string
                }
                catch [System.Management.Automation.RemoteException] {
                    $outputbox.appendtext( "No Eye Tracker Connected`r`n")
                }
            }
            else { 
                $outputbox.appendtext("File FWUpgrade32.exe is missing!`r`n" )
            }

            $time = Get-Date -UFormat %H:%M:%S
            Add-content c:\Output.txt $time, $NewProfile, $getinfo
            Start-sleep -s 3
            .\TobiiDynavox.EyeAssist.Engine.exe -x
            Start-sleep -s 3
            .\TobiiDynavox.EyeAssist.Engine.exe
            Start-sleep -s 3
            $a++
        } while ($a -le $b)
    }
    elseif ($b -match "3") {
        $Keys = ("HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\UserProfiles" )
        Remove-item $Keys -Recurse -ErrorAction Ignore
    }
    else { $outputbox.appendtext("The number you entered is not applicable, try again`r`n") }
    $outputbox.appendtext("Done! `r`n")
}

#B28
Function RetrieveUnreleased {
    $outputBox.clear()
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'UN Activate Unreleased tool'
    $msg = "Press:  `r`n1 to set value to True, `r`n2 to set the value to False, `r`n3 to remove the key"
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    
    $regpath = 'HKLM:\SOFTWARE\WOW6432Node\Tobii\Update Notifier'
    if (!(Test-Path $regpath)) {
        $regpath = 'HKLM:\SOFTWARE\WOW6432Node\Tobii\I-Series\Update Notifier'
    }

    $Check = Get-ItemProperty -Path "$regpath" -Name RetrieveUnreleasedVersions -ErrorAction SilentlyContinue
    if ($b -match "1") {
        if ($Check) {
            Set-ItemProperty -Path "$regpath" -Name "RetrieveUnreleasedVersions" -Value 'True'
        }
        else {
            New-ItemProperty -Path "$regpath" -Name "RetrieveUnreleasedVersions" -PropertyType "String" -Value 'True'
        }
        $outputbox.appendtext("Value set to True`r`n")
    }
    elseif ($b -match "2") {
        if ($Check) { 
            Set-ItemProperty -Path "$regpath" -Name "RetrieveUnreleasedVersions" -Value 'False'
        }
        else {
            New-ItemProperty -Path "$regpath" -Name "RetrieveUnreleasedVersions" -PropertyType "String" -Value 'False'
        }
        $outputbox.appendtext("Value set to False`r`n") 
    }
    elseif ($b -match "3") {
        if ($Check) {
            Remove-ItemProperty -Path "$regpath" -Name "RetrieveUnreleasedVersions"
            $outputbox.appendtext("String has been removed`r`n")
        }
    }
    else { $outputbox.appendtext("N/A`r`n") }
    $outputbox.appendtext("Done!`r`n")
}

#B29 GB2SmbiosTool
Function GB2SmbiosTool {
    $outputBox.clear()
    $outputbox.appendtext("Start `r`n")
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "GB2SmbiosTool.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        Start-Process cmd -Verb runAs "/c `"GB2SmbiosTool.exe`"" 
    }
    else { 
        $outputbox.appendtext("File GB2SmbiosTool.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B30
Function SMBios {
    $outputBox.clear()
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'SMBios tool'
    $msg = "Press 1 to run getSMBIOSvalues.cmd, `r`n 2 setName.cmd, `r`n 3 setSerialNumber.cmd, `r`n 4 setVendor.cmd"
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "getSMBIOSvalues.cmd" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
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
        else { $outputbox.appendtext("N/A`r`n") }
    }
    else { 
        $outputbox.appendtext("File getSMBIOSvalues.cmd is missing!`r`n" )
    }
    $outputbox.appendtext("Done!`r`n")
}

#B31
Function ETSamples {
    $outputBox.clear()
    $outputBox.appendtext( "Starting TD region interaction sample...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "Tdx.EyeTracking.RegionInteraction.EyeAssist.Sample.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        .\Tdx.EyeTracking.RegionInteraction.EyeAssist.Sample.exe
    }
    else { 
        $outputbox.appendtext("File Tdx.EyeTracking.RegionInteraction.EyeAssist.Sample.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B32
Function Diagnostic {
    $outputBox.clear()
    $outputBox.appendtext( "Run diagnostics application for Interaction...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "Tobii.EyeX.Diagnostics.Application.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        start-process cmd "/c `"Tobii.EyeX.Diagnostics.Application.exe`""
    }
    else { 
        $outputbox.appendtext("File Tobii.EyeX.Diagnostics.Application.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B33
Function SETest {
    $outputBox.clear()
    $outputBox.appendtext( "running Stream Engine Test...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "tests.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        .\tests.exe
    }
    else { 
        $outputbox.appendtext("File tests.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B34
Function InternalSE {
    $outputBox.clear()
    $outputBox.appendtext( "Starting Stream Engine Sample app...`r`n" )
    $fpath = (Get-ChildItem -Path "$PSScriptRoot" -Filter "sample.exe" -Recurse).FullName | Split-Path 
    if ($fpath.count -gt 0) {
        $Outputbox.Appendtext("Files found!`r`n" )
        Set-Location $fpath
        Start-Process .\sample.exe
    }
    else { 
        $outputbox.appendtext("File sample.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#Windows forms
$Optionlist = @("Remove Progressive Sweet", "Remove PCEye5 Bundle", "Remove all ET SW", "Remove WC&GP Bundle", "Remove VC++", "Remove PCEye Package", "Remove Communicator", "Remove Compass", "Remove TGIS only", "Remove TGIS profile calibrations", "Remove all users C5", "Reset TETC", "Backup Gaze Interaction", "Copy License")
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
$Button.Location = New-Object System.Drawing.Size(10, 50)
$Button.Size = New-Object System.Drawing.Size(110, 30)
$Button.Text = "Start"
$Button.Font = New-Object System.Drawing.Font ("" , 12, [System.Drawing.FontStyle]::Regular)
$Form.Controls.Add($Button)
$Button.Add_Click{ selectedscript }

#B1 Button1 "List Tobii Software"
$Button1 = New-Object System.Windows.Forms.Button
$Button1.Location = New-Object System.Drawing.Size(410, 0)
$Button1.Size = New-Object System.Drawing.Size(90, 30)
$Button1.Text = "Tobii Softwaredsdf"
$Button1.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button1)
$Button1.Add_Click{ ListApps }

#B2 Button2 "List active Tobii processes"
$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(410, 30)
$Button2.Size = New-Object System.Drawing.Size(90, 30)
$Button2.Text = "Active Process+Service"
$Button2.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button2)
$Button2.Add_Click{ GetProcesses }

#B3 Button3 "TTSWversions"
$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Size(410, 60)
$Button3.Size = New-Object System.Drawing.Size(90, 30)
$Button3.Text = "TT SW versions"
$Button3.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button3)
$Button3.Add_Click{ TTSWversions }

#B4 Button4 "Check IS5 PID"
$Button4 = New-Object System.Windows.Forms.Button
$Button4.Location = New-Object System.Drawing.Size(410, 90)
$Button4.Size = New-Object System.Drawing.Size(90, 30)
$Button4.Text = "List IS5 PID"
$Button4.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button4)
$Button4.Add_Click{ IS5PID }

#B5 Button5 "List Tobii Drivers"
$Button5 = New-Object System.Windows.Forms.Button
$Button5.Location = New-Object System.Drawing.Size(410, 120)
$Button5.Size = New-Object System.Drawing.Size(90, 30)
$Button5.Text = "Tobii Drivers"
$Button5.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button5)
$Button5.Add_Click{ ListDrivers }

#B6 Button6 "ET fw"
$Button6 = New-Object System.Windows.Forms.Button
$Button6.Location = New-Object System.Drawing.Size(410, 150)
$Button6.Size = New-Object System.Drawing.Size(90, 30)
$Button6.Text = "ET firmware"
$Button6.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button6)
$Button6.Add_Click{ ETfw }

#B7 Button7 ".NET version"
$Button7 = New-Object System.Windows.Forms.Button
$Button7.Location = New-Object System.Drawing.Size(410, 180)
$Button7.Size = New-Object System.Drawing.Size(90, 30)
$Button7.Text = ".NET v."
$Button7.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button7)
$Button7.Add_Click{ GetFrameworkVersionsAndHandleOperation }

#B8 Button8 "BatteryLog"
$Button8 = New-Object System.Windows.Forms.Button
$Button8.Location = New-Object System.Drawing.Size(410, 210)
$Button8.Size = New-Object System.Drawing.Size(90, 30)
$Button8.Text = "Battery Log"
$Button8.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button8)
$Button8.Add_Click{ BatteryLog }

#B9 Button9 "Show Track status"
$Button9 = New-Object System.Windows.Forms.Button
$Button9.Location = New-Object System.Drawing.Size(410, 240)
$Button9.Size = New-Object System.Drawing.Size(90, 30)
$Button9.Text = "Show/hide Track S."
$Button9.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button9)
$Button9.Add_Click{ TrackStatus }

#B10 Button10 "LogCollector"
$Button10 = New-Object System.Windows.Forms.Button
$Button10.Location = New-Object System.Drawing.Size(410, 270)
$Button10.Size = New-Object System.Drawing.Size(90, 30)
$Button10.Text = "LogCollector"
$Button10.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button10)
$Button10.Add_Click{ LogCollector }

#B11 Button11 "WCF"
$Button11 = New-Object System.Windows.Forms.Button
$Button11.Location = New-Object System.Drawing.Size(410, 300)
$Button11.Size = New-Object System.Drawing.Size(90, 30)
$Button11.Text = "WCF"
$Button11.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button11)
$Button11.Add_Click{ WCF }

#B12 Button12 "Before Uninstall GG"
$Button12 = New-Object System.Windows.Forms.Button
$Button12.Location = New-Object System.Drawing.Size(410, 330)
$Button12.Size = New-Object System.Drawing.Size(90, 30)
$Button12.Text = "BeforeUnins.G"
$Button12.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button12)
$Button12.Add_Click{ BeforeUninstallGG }

#B13 Button13 "DriverSetup"
$Button13 = New-Object System.Windows.Forms.Button
$Button13.Location = New-Object System.Drawing.Size(410, 360)
$Button13.Size = New-Object System.Drawing.Size(90, 30)
$Button13.Text = "DriverSetupGG"
$Button13.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button13)
$Button13.Add_Click{ DriverSetup }

#B14 Button14 "uninstall_gibbon"
$Button14 = New-Object System.Windows.Forms.Button
$Button14.Location = New-Object System.Drawing.Size(410, 390)
$Button14.Size = New-Object System.Drawing.Size(90, 30)
$Button14.Text = "UninstallGG"
$Button14.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button14)
$Button14.Add_Click{ uninstall_gibbon }

#B15 Button15 "SetDebugLogging"
$Button15 = New-Object System.Windows.Forms.Button
$Button15.Location = New-Object System.Drawing.Size(410, 420)
$Button15.Size = New-Object System.Drawing.Size(90, 30)
$Button15.Text = "SetDebugLog"
$Button15.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button15)
$Button15.Add_Click{ SetDebugLogging }

#B16 Button16 "ModelCheck"
$Button16 = New-Object System.Windows.Forms.Button
$Button16.Location = New-Object System.Drawing.Size(410, 450)
$Button16.Size = New-Object System.Drawing.Size(90, 30)
$Button16.Text = "ETModel Check"
$Button16.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button16)
$Button16.Add_Click{ ModelCheck }

#B17 Button17 "HWInfo"
$Button17 = New-Object System.Windows.Forms.Button
$Button17.Location = New-Object System.Drawing.Size(410, 480)
$Button17.Size = New-Object System.Drawing.Size(90, 30)
$Button17.Text = "HWInfo"
$Button17.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button17)
$Button17.Add_Click{ HWInfo }

#B18 Button18 "Install PDK"
$Button18 = New-Object System.Windows.Forms.Button
$Button18.Location = New-Object System.Drawing.Size(500, 0)
$Button18.Size = New-Object System.Drawing.Size(90, 30)
$Button18.Text = "Install PDK"
$Button18.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button18)
$Button18.Add_Click{ InstallPDK }

#B19 Button19 Restart Services
$Button19 = New-Object System.Windows.Forms.Button
$Button19.Location = New-Object System.Drawing.Size(500, 30)
$Button19.Size = New-Object System.Drawing.Size(90, 30)
$Button19.Text = "RestartServices"
$Button19.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button19)
$Button19.Add_Click{ RestartProcesses }

#B20 Button20 "Delete services"
$Button20 = New-Object System.Windows.Forms.Button
$Button20.Location = New-Object System.Drawing.Size(500, 60)
$Button20.Size = New-Object System.Drawing.Size(90, 30)
$Button20.Text = "Delete Services"
$Button20.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button20)
$Button20.Add_Click{ DeleteServices }

#B21 Button21 "Remove Drivers"
$Button21 = New-Object System.Windows.Forms.Button
$Button21.Location = New-Object System.Drawing.Size(500, 90)
$Button21.Size = New-Object System.Drawing.Size(90, 30)
$Button21.Text = "Delete Drivers"
$Button21.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button21)
$Button21.Add_Click{ RemoveDrivers }

#B22 Button22 "Reset IS5 to bootloader"
$Button22 = New-Object System.Windows.Forms.Button
$Button22.Location = New-Object System.Drawing.Size(500, 120)
$Button22.Size = New-Object System.Drawing.Size(90, 30)
$Button22.Text = "ResetETBOOT"
$Button22.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button22)
$Button22.Add_Click{ resetBOOT }

#B23 Button23 "FW Upgrade"
$Button23 = New-Object System.Windows.Forms.Button
$Button23.Location = New-Object System.Drawing.Size(500, 150)
$Button23.Size = New-Object System.Drawing.Size(90, 30)
$Button23.Text = "FW Upgrade"
$Button23.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button23)
$Button23.Add_Click{ FWUpgrade }

#B24 Button24 "VirtWin"
$Button24 = New-Object System.Windows.Forms.Button
$Button24.Location = New-Object System.Drawing.Size(500, 180)
$Button24.Size = New-Object System.Drawing.Size(90, 30)
$Button24.Text = "VirtWin TT"
$Button24.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button24)
$Button24.Add_Click{ VirtWin }

#B25 Button25 "Check ET connection through Service"
$Button25 = New-Object System.Windows.Forms.Button
$Button25.Location = New-Object System.Drawing.Size(500, 210)
$Button25.Size = New-Object System.Drawing.Size(90, 30)
$Button25.Text = "Check ET con."
$Button25.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button25)
$Button25.Add_Click{ ETConnection }

#B26 Button26 "Run GG HB_tool"
$Button26 = New-Object System.Windows.Forms.Button
$Button26.Location = New-Object System.Drawing.Size(500, 240)
$Button26.Size = New-Object System.Drawing.Size(90, 30)
$Button26.Text = "HB-tool GG"
$Button26.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button26)
$Button26.Add_Click{ HBTool }

#B27 Button27 "EAProfileCreation"
$Button27 = New-Object System.Windows.Forms.Button
$Button27.Location = New-Object System.Drawing.Size(500, 270)
$Button27.Size = New-Object System.Drawing.Size(90, 30)
$Button27.Text = "EA Profile"
$Button27.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button27)
$Button27.Add_Click{ EAProfileCreation }

#B28 Button28 "RetrieveUnreleased"
$Button28 = New-Object System.Windows.Forms.Button
$Button28.Location = New-Object System.Drawing.Size(500, 300)
$Button28.Size = New-Object System.Drawing.Size(90, 30)
$Button28.Text = "RetrieveUN"
$Button28.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button28)
$Button28.Add_Click{ RetrieveUnreleased }

#B29 Button29 "GB2SmbiosTool"
$Button29 = New-Object System.Windows.Forms.Button
$Button29.Location = New-Object System.Drawing.Size(500, 330)
$Button29.Size = New-Object System.Drawing.Size(90, 30)
$Button29.Text = "SmbiosTool3"
$Button29.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button29)
$Button29.Add_Click{ GB2SmbiosTool }

#B30 Button30 "SMBios"
$Button30 = New-Object System.Windows.Forms.Button
$Button30.Location = New-Object System.Drawing.Size(500, 360)
$Button30.Size = New-Object System.Drawing.Size(90, 30)
$Button30.Text = "SMBIOS"
$Button30.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button30)
$Button30.Add_Click{ SMBios }

#B31 Button31 "ETSamples"
$Button31 = New-Object System.Windows.Forms.Button
$Button31.Location = New-Object System.Drawing.Size(500, 390)
$Button31.Size = New-Object System.Drawing.Size(90, 30)
$Button31.Text = "TD RI Samples"
$Button31.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button31)
$Button31.Add_Click{ ETSamples }

#B32 Button32 "Diagnostic"
$Button32 = New-Object System.Windows.Forms.Button
$Button32.Location = New-Object System.Drawing.Size(500, 420)
$Button32.Size = New-Object System.Drawing.Size(90, 30)
$Button32.Text = "RIDiagnostics"
$Button32.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button32)
$Button32.Add_Click{ Diagnostic }

#B33 Button33 "StreamEngineTest"
$Button33 = New-Object System.Windows.Forms.Button
$Button33.Location = New-Object System.Drawing.Size(500, 450)
$Button33.Size = New-Object System.Drawing.Size(90, 30)
$Button33.Text = "SE-Test"
$Button33.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button33)
$Button33.Add_Click{ SETest }

#B34 Button34 "InternalSE"
$Button34 = New-Object System.Windows.Forms.Button
$Button34.Location = New-Object System.Drawing.Size(500, 480)
$Button34.Size = New-Object System.Drawing.Size(90, 30)
$Button34.Text = "Internal SE"
$Button34.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button34)
$Button34.Add_Click{ InternalSE }



#Form name + activate form.
$Form.Text = "Support Tool 1.6.4"
$Form.Add_Shown( { $Form.Activate() })
$Form.ShowDialog()