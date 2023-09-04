#Arthur Ammar Elyas, ammar.elyas@tobiidynavox.com
#File version 
$fileversion = "SupportTool v1.6.11.ps1"

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

    if ($DropDownBox.Selecteditem -eq "Remove Progressive Suite") {
        $Outputbox.Clear()
        UninstallProgressiveSuite
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
    elseif ($DropDownBox.Selecteditem -eq "Remove C5 Emails") {
        $Outputbox.Clear()
        DeleteEmailsC5
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

#A1 Uninstalls Progressive Suite
Function UninstallProgressiveSuite {
    # https://stackoverflow.com/questions/46310266/accessing-dynamically-created-variables-inside-a-powershell-function
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $form = New-Object System.Windows.Forms.Form
    $flowlayoutpanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonOK = New-Object System.Windows.Forms.Button
    $CancelButton = New-Object System.Windows.Forms.Button
  
    $products = "Browse", "Browse (Beta)", "Browse (Development)", "Control", "Control (Beta)", "Control (Development)", "Phone", "Phone (Beta)", "Phone (Development)", "Talk", "Talk (Beta)", "Talk (Development)", "Switcher", "Switcher (Beta)"
    
    foreach ($product in $products) {
        $TobiiVer += @(Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object {
            ($_.Displayname -eq "Tobii Dynavox $product") 
            } | Select-Object Displayname, UninstallString | Sort-Object Displayname  )
    } 
   
    if ($TobiiVer) {   
        $usernames = @($TobiiVer.Displayname)
        $totalvalues = ($usernames.count)

        $formsize = 100 + (30 * $totalvalues)
        $flowlayoutsize = 10 + (30 * $totalvalues)
        $buttonplacement = 40 + (30 * $totalvalues)
        $script:CheckBoxArray = @()
    
        $form_Load = {
            foreach ($user in $usernames) {
                $DynamicCheckBox = New-object System.Windows.Forms.CheckBox

                $DynamicCheckBox.Margin = '10, 8, 0, 0'
                $DynamicCheckBox.Name = $user
                #changed to make the text look better
                $DynamicCheckBox.Size = '330, 22' 
                $DynamicCheckBox.Text = "" + $user

                $DynamicCheckBox.TextAlign = 'MiddleLeft'
                $flowlayoutpanel.Controls.Add($DynamicCheckBox)
                $script:CheckBoxArray += $DynamicCheckBox
            }       
        }

        $form.Controls.Add($flowlayoutpanel)
        $form.Controls.Add($buttonOK)
        $form.Controls.Add($CancelButton)
        $form.AcceptButton = $buttonOK
        $form.CancelButton = $CancelButton
        $form.AutoScaleDimensions = '8, 17'
        $form.AutoScaleMode = 'Font'
        $form.ClientSize = "500 , $formsize"
        $form.FormBorderStyle = 'FixedDialog'
        $form.Margin = '5, 5, 5, 5'
        $form.MaximizeBox = $False
        $form.MinimizeBox = $False
        $form.Name = 'form1'
        $form.StartPosition = 'CenterScreen'
        $form.Text = 'Progressive Suite'
        $form.add_Load($($form_Load))
    } 

    $flowlayoutpanel.BorderStyle = 'FixedSingle'
    $flowlayoutpanel.Location = '48, 13'
    $flowlayoutpanel.Margin = '4, 4, 4, 4'
    $flowlayoutpanel.Name = 'flowlayoutpanel1'
    $flowlayoutpanel.AccessibleName = 'flowlayoutpanel1'
    if ($flowlayoutsize) {
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
    
        $CancelButton.Anchor = 'Bottom, Right'
        $CancelButton.DialogResult = 'Cancel'
        $CancelButton.Location = "283, $buttonplacement"
        $CancelButton.Margin = '4, 4, 4, 4'
        $CancelButton.Name = 'CancelButton'
        $CancelButton.Size = '100, 30'
        $CancelButton.TabIndex = 0
        $CancelButton.Text = '&Cancle'
    }
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {   
        foreach ($cbox in $CheckBoxArray) {
            if ($cbox.CheckState -eq "Checked") {
        
                #If first answer equals yes or no
                $Uninstname = (Compare-Object -DifferenceObject $TobiiVer.displayname -ReferenceObject $cbox.Name -CaseSensitive -ExcludeDifferent -IncludeEqual | Select-Object InputObject).InputObject
                if ($null -ne $Uninstname) {
                
                    #$services = (Get-service -Name "*tdx*").name
                    #foreach ($service in $services) {
                    #    stop-Service $service
                    #}
                    foreach ($Uninstnames in $Uninstname) {
                        $Outputbox.Appendtext( "selected App = $Uninstnames`r`n")
                        if ($Uninstnames -match "Control") {
                            $Uninstnames = $Uninstnames -replace "Control", "Computer Control"
                        }
                        $Uninstnames0 = $Uninstnames -replace ('[(\)]', '')
                        $Uninstnames1 = $Uninstnames -replace ('\s+\(', ' Launcher ') -replace '[(\)]', ''
                        $Uninstnames2 = $Uninstnames -replace ('\s+\(', ' Updater Service ') -replace '[(\)]', ''
                        $Uninstnames3 = $Uninstnames -replace ('\s+\(', ' Launcher (') 
                        $Uninstnames4 = $Uninstnames -replace ('\s+\(', ' Updater Service (')                    
                        $Uninstnames5 = $Uninstnames + ' Launcher'
                        $Uninstnames6 = $Uninstnames + ' Updater Service'
                        if ($Uninstname -match "Beta") {
                            $Uninstnames7 = $Uninstnames -replace "(Beta)", "Review" -replace '[(\)]', ''
                        }
                        $Uninstnames8 = $Uninstnames7 -replace "Review", "Launcher Review"
                        $Uninstnames9 = $Uninstnames7 -replace "Review", "Updater Service Review"
                        $Uninstnames10 = $Uninstnames -replace ' Computer', ''
                        $Uninstnames11 = $Uninstnames4 -replace ' Computer', ''
                        $Uninstnames12 = $Uninstnames6 -replace ' Computer', ''

                        $AllNames = "$Uninstnames", "$Uninstnames0", "$Uninstnames1", "$Uninstnames2", "$Uninstnames3", "$Uninstnames4" , "$Uninstnames5", 
                        "$Uninstnames6", "$Uninstnames7", "$Uninstnames8", "$Uninstnames9" , "$Uninstnames10", "$Uninstnames11", "$Uninstnames12" 
                    
                        $b = $AllNames | Select-Object -Unique
                    
                        foreach ($bs in $b) {
                            $testpath = $bs.replace('Tobii Dynavox ', '')
                            $GCI = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
                        ($_.Displayname -eq "$bs") 
                            } | Select-Object Displayname, UninstallString
                            $GCIDisplay = $GCI.Displayname
                            $UninstallString = $GCI.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                            Start-process "msiexec.exe" -arg "/X $UninstallString /quiet /norestart" -Wait
                            if ($testpath) { 
                                $paths = ( 
                                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\$testpath",
                                    "$ENV:USERPROFILE\AppData\Local\Tobii Dynavox\$testpath",
                                    "$ENV:ProgramData\Tobii Dynavox\$testpath",
                                    "C:\Program Files\Tobii Dynavox\$testpath",
                                    "HKCU:\Software\Tobii Dynavox\$testpath",
                                    "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\$testpath"
                                )
                                foreach ($path in $paths) {
                                    if ((Test-path -path "$path")) {
                                        Remove-item $path -Recurse -ErrorAction Ignore
                                    }                            
                                }
                            } 
                        }
                    }
                
                    #$services = (Get-service -Name "*tdx*").name
                    #foreach ($service in $services) {
                    #    start-Service $service
                    #}
                }
            }
        }
    
        Remove-Variable * -ErrorAction SilentlyContinue
        Remove-Variable checkbox*

        $Outputbox.Appendtext("Done!`r`n")
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        return
    }
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
    $PCEyeSWs = "Tobii Experience Software For Windows (PCEye5)", "Tobii Dynavox Control", "Tobii Dynavox Control", "Tobii Dynavox Control Updater Service", "Tobii Dynavox Eye Tracking", "Tobii Dynavox Switcher", "Tobii Dynavox Switcher Updater Service", "Tobii Dynavox Update Notifier"
    foreach ($PCEyeSW in $PCEyeSWs) {
        $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
        Get-ItemProperty | Where-Object { 
            ($_.Displayname -eq $PCEyeSW )
        } | Select-Object Displayname, UninstallString
        
        ForEach ($ver in $TobiiVer) {
            $Uninstname = $ver.Displayname
            $uninst = $ver.UninstallString -replace "msiexec.exe", "" -Replace "/I", "" -Replace "/X", ""
            $uninst = $uninst.Trim()
            $Outputbox.Appendtext( "Uninstalling - " + "$Uninstname`r`n" )
            start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
        }
    }

    $DeleteServices = Get-Service -Name '*TobiiIS*' , '*TobiiG*' | Stop-Service -Force -passthru -ErrorAction ignore
    foreach ($Service in $DeleteServices) {
        $outputbox.appendtext(" Deleating - " + "$Service `r`n" )
        sc.exe delete $Service
    }

    $DeleteDrivers = Get-WindowsDriver -Online | Where-Object { $_.ProviderName -match "Tobii" } | Select-Object Driver
    ForEach ($Drivers in $DeleteDrivers) {
        $outputBox.appendtext( "Removing Drivers - " + "$Drivers`r`n" )
        pnputil /delete-driver $Drivers.Driver /force /uninstall
    }
    
    $paths = (
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control Bundle",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Overlays",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Shared Predictions",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Switcher",
        "$ENV:USERPROFILE\AppData\Local\Tobii\Installer",
        "$ENV:ProgramData\Tobii Dynavox\Computer Control",
        "$ENV:ProgramData\Tobii Dynavox\Switcher",
        "$ENV:ProgramData\Tobii Dynavox\Update Notifier"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item "$path" -Recurse -Force
        }
    }

    $PDKPath = "$ENV:ProgramData\Tobii\Tobii Platform Runtime\IS5LARGEPCEYE5",
    "$ENV:LocalAppData\Tobii\Installer"
    Foreach ($PDKPath in $PDKPaths) {
        if (Test-path $PDKPath) {
            $test = Get-ChildItem -Path "$PDKPath" -Recurse -af |  foreach-object { $_.FullName }
            Remove-Item "$PDKPath" -Recurse 
        }
    } 
    $Keys = (
        "HKCU:\Software\Tobii\EyeAssist",
        "HKCU:\Software\Tobii Dynavox\Analytics",
        "HKLM:\SOFTWARE\WOW6432Node\Tobii Dynavox\Computer Control Updater Service",
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation"
    )

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }
    if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\OEMInformation\EyeTrackerModel") {
        Remove-ItemProperty -path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\OEMInformation" -Name "EyeTrackerModel"
    }

    if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation") { 
        Get-Item -Path "HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation"
    }

    if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation\EyeTrackerModel") {
        Remove-ItemProperty -path "HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation" -Name "EyeTrackerModel"
    }
    
    $PCEyeSWs = "Tobii Experience Software For Windows (PCEye5)", "Tobii Dynavox Control", "Tobii Dynavox Control", "Tobii Dynavox Control Updater Service", "Tobii Dynavox Eye Tracking", "Tobii Dynavox Switcher", "Tobii Dynavox Switcher Updater Service", "Tobii Dynavox Update Notifier"
    foreach ($PCEyeSW in $PCEyeSWs) {
        $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
        Get-ItemProperty | Where-Object { 
            ($_.Displayname -eq $PCEyeSW )
        } | Select-Object Displayname
    }
    if ($PCEyeSWs) {
        $outputBox.appendtext( "$PCEyeSWs couldn't be uninstalled. Reboot your device and try again.`r`n" )
    }
    $Outputbox.Appendtext( "Done!`r`n" )
}

#A3 Uninstalls ALL Tobii Device Drivers For Windows Bundle
Function UninstallTobiiDeviceDriversForWindows {

    #If first answer equals yes or no
    $answer1 = $wshell.Popup("This will remove all software included in Tobii ET SW.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress..`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled: Remove Tobii ET SW`r`n" )
        Return
    }

    $TempPaths = "$ENV:USERPROFILE\AppData\Local\Temp", "$ENV:USERPROFILE\AppData\Local\Tobii\Installer"
    $ErrorPath = "$ENV:USERPROFILE\AppData\Local\Temp\ErrorLogs"
    $RegPath = "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig"
    $TempPathReg = "$ENV:USERPROFILE\AppData\Local\Temp\EyeXConfig.reg"
	
    if (!(Test-Path "$ErrorPath")) {
        New-Item -Path "$ErrorPath" -ItemType Directory   
    }
    if (!(Test-Path "$ErrorPath\InstallerError.txt") -or !(Test-Path "$ErrorPath\InstallerError2.txt") -or !(Test-Path "$ErrorPath\InstallerError3.txt")) {
        New-Item -Path $ErrorPath -Name "InstallerError.txt" -ItemType "file"
        New-Item -Path $ErrorPath -Name "InstallerError2.txt" -ItemType "file"
        New-Item -Path $ErrorPath -Name "InstallerError3.txt" -ItemType "file"
    }
    else {
        Clear-Content -Path "$ErrorPath\InstallerError.txt"
        Clear-Content -Path "$ErrorPath\InstallerError2.txt"
        Clear-Content -Path "$ErrorPath\InstallerError3.txt"
    }
	
    Foreach ($TempPath in $TempPaths) {
        if (test-path "$TempPaths") {
            Set-Location $TempPath
            $Installercontent = Get-ChildItem "tobii*.log" -Recurse -File | Sort-Object name -desc | Select-Object -expand Fullname
            foreach ($NewInstallercontent in $Installercontent) {
                New-Item -Path $ErrorPath -Name "temp.txt" -ItemType "file"
                Get-Content -Path "$NewInstallercontent" -Raw | ForEach-Object -Process { $_ -replace "- `r`n", '- ' } | Add-Content -Path "$ErrorPath\temp.txt"
                $string = "Executing\s+op\:\s+CustomActionSchedule\(Action\=DisconnectDevices,ActionType\=3073,Source\=BinaryData,Target\=WixQuietExec,CustomActionData\="
                $content = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "$string" -AllMatches | ForEach-Object -Process { $_ -replace ".*CustomActionData=" -replace "-inf.*" } | ForEach-Object -Process { $_ -replace ("`"", "") }
                add-Content "$ErrorPath\InstallerError.txt" -value $content, "`n"
                Remove-Item "$ErrorPath\temp.txt"
            }
        }

        (Get-Content "$ErrorPath\InstallerError.txt") | Where-Object { $_.trim() -ne "" } | set-content "$ErrorPath\InstallerError2.txt"
        $Output = (Get-Content "$ErrorPath\InstallerError2.txt")
        if ($null -ne $Output) {
            foreach ($line in $Output) {
                $array = $line.split("\")
                $path = [string]::Join("\", $array[0..($array.length - 2)]) 
                Add-Content -Path "$ErrorPath\InstallerError3.txt" -Value $path
            }
        }

        if ($Null -ne (Get-Content "$ErrorPath\InstallerError3.txt")) {
            $Content2 = Get-Content -Path "$ErrorPath\InstallerError3.txt"
            $OutputBox.AppendText("Copy DriverSetup to specific path`r`n")
            foreach ($NewContent2 in $Content2) {    
                New-Item -ItemType Directory -Force -Path $NewContent2
                $fpathDriver = Get-ChildItem -Path $PSScriptRoot -Filter "DriverSetup.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
                Set-Location $fpathDriver
                Copy-Item -Path ("DriverSetup.exe") -Destination $NewContent2
            }
        } 
    }

    if ((Test-Path -Path $RegPath) -and (!(Test-Path -path $TempPathReg))) {
        $Outputbox.Appendtext("Backup profiles in %temp%\EyeXConfig.reg`r`n" )
        Invoke-Command { reg export "HKLM\SOFTWARE\WOW6432Node\Tobii\EyeXConfig" $TempPathReg }
    }

    #Getting FW version
    $fpathfw = Get-ChildItem -Path $PSScriptRoot -Filter "Tdx.EyeTrackerInfo.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpathfw.count -gt 0) {
        Set-Location $fpathfw
        $ETModel = .\Tdx.EyeTrackerInfo.exe --model
    }
    else { 
        $outputbox.appendtext("File Tdx.EyeTrackerInfo.exe is missing!`r`n" )
    }
	
    if ($ETModel -match "IS5_Gibbon_Gaze" ) { 
        $outputBox.appendtext( "Running BeforeUninstall.bat script.`r`n" )
        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "BeforeUninstall.bat" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        if ($fpath.count -gt 0) {
            Set-Location $fpath
            $Installer = Start-Process -FilePath "$fpath\BeforeUninstall.bat" -Wait
            $Outputbox.appendtext("$Installer`r`n")
        }
        else { 
            $outputbox.appendtext("File BeforeUninstall.bat is missing!`r`n" )
        }
    } 
	
    $GetProcess = stop-process -Name "*TobiiDynavox*" -Force
    if ($GetProcess) {
        $Outputbox.appendtext("Stopping $GetProcess `r`n" )
    }

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | 
    Get-ItemProperty | Where-Object { 
        ($_.Displayname -Match "Tobii Device Drivers For Windows") -or
        ($_.Displayname -Match "Tobii Experience Software") -or
        ($_.Displayname -Match "Tobii Eye Tracking For Windows") -or
        ($_.Displayname -Match "Tobii Dynavox Eye Tracking Driver") -or
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
        
    $TobiiDriver = Get-WindowsDriver -Online | Where-Object { $_.ProviderName -match "Tobii" } | Select-Object Driver
    ForEach ($NewTobiiDriver in $TobiiDriver) {
        $outputBox.appendtext( "Removing Drivers - " + "$NewTobiiDriver`r`n" )
        pnputil /delete-driver $NewTobiiDriver.Driver /force /uninstall
    }

    #Removes Tobii related folders
    $TobiiFiles = (
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist",
        "$Env:USERPROFILE\AppData\Local\Tobii\Installer",
        "$Env:USERPROFILE\AppData\Local\Tobii_AB\",
        "C:\Program Files\Tobii\Tobii EyeX",
        "$ENV:ProgramData\TetServer",
        "$ENV:ProgramData\Tobii\HelloDMFT",
        "$ENV:ProgramData\Tobii\Statistics",
        "$ENV:ProgramData\Tobii\Tobii Interaction",
        "$ENV:ProgramData\Tobii\Tobii Stream Engine",
        "$ENV:ProgramData\Tobii\Statistics",
        "$ENV:ProgramData\Tobii\Tobii Interaction",
        "$ENV:ProgramData\Tobii\Tobii Platform Runtime",
        "$ENV:ProgramData\Tobii\EulaHasBeenAccepted.txt"
    )
    $runtimepath = "$ENV:ProgramData\Tobii\Tobii Platform Runtime" 
    if (Test-Path $runtimepath) {
        $folder = Get-ChildItem -Path "$ENV:ProgramData\Tobii\Tobii Platform Runtime" -Directory
        $folder = $folder.Name 
        foreach ($folders in $folder) {
            if ( $folders -match "IS5") {
                Get-ChildItem -Path "$ENV:ProgramData\Tobii\Tobii Platform Runtime\$folders" -Recurse -af |  foreach-object { $_.FullName }
                Remove-Item "$ENV:ProgramData\Tobii\Tobii Platform Runtime\$folders" -Recurse 
            }
        }
    }
    foreach ($NewTobiiFiles in $TobiiFiles) {
        if (Test-Path $NewTobiiFiles) {
            $Outputbox.appendtext( "Removing - " + "$NewTobiiFiles`r`n" )
            Remove-Item $NewTobiiFiles -Recurse -Force -ErrorAction Ignore
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
        "$ENV:ProgramData\Tobii Dynavox\Gaze Interaction",
        "C:\Program Files (x86)\Tobii Dynavox\PCEye Update Notifier"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item $path -Recurse -Force -ErrorAction Ignore
        }
    }

    $Keys = (
        "HKCU:\SOFTWARE\Tobii\PCEye\Update Notifier",
        "HKCU:\SOFTWARE\Tobii\PCEye", 
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\PCEye\Update Notifier",
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\PCEye"
    )
				
    foreach ($key in $Keys) {
        if (test-path $key) {
            $Outputbox.appendtext( "Removing - " + "$key`r`n" )
            Remove-Item $key -Force -ErrorAction ignore
        }
    }

    $OEMInfoPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\OEMInformation"
    $EyeTrackerModel = "EyeTrackerModel"
    if ((Get-ItemProperty $OEMInfoPath).PSObject.Properties.Name -contains $EyeTrackerModel) { Remove-ItemProperty -path $OEMInfoPath -Name "EyeTrackerModel" }

    $Outputbox.appendtext( "Done!`r`n" )
}

#A7 Uninstall Communicator
Function UninstallCommunicator {

    Add-Type -AssemblyName System.Windows.Forms    
    Add-Type -AssemblyName System.Drawing

    # Build Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "C5"
    $Form.Size = New-Object System.Drawing.Size(300, 300)
    $Form.StartPosition = "CenterScreen"
    #$Form.Topmost = $True

    # Add Button1
    $Button1 = New-Object System.Windows.Forms.Button
    $Button1.Location = New-Object System.Drawing.Size(75, 50)
    $Button1.Size = New-Object System.Drawing.Size(150, 50)
    $Button1.Text = "Remove only C5"
    $Form.Controls.Add($Button1)
    
    # Add Button2
    $Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Size(75, 120)
    $Button2.Size = New-Object System.Drawing.Size(150, 50)
    $Button2.Text = "Remove C5 Suit"
    $Form.Controls.Add($Button2)
    
    #Add Button event 
    $Button1.Add_Click( {
            #If second answer equals yes or no - if "Yes" then it will call the function CopyLicenses and then continue.
            $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?", 0, "Caution", 48 + 4)
            if ($answer2 -eq 6) { CopyLicenses }

            elseif ($answer2 -ne 6) { $Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" ) }

            $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
            Get-ItemProperty | Where-Object { 
                $_.Displayname -match "Tobii Dynavox Communicator" 

            } | Select-Object Publisher, Displayname, UninstallString

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
    )
    $Button2.Add_Click( {
        
            #If second answer equals yes or no - if "Yes" then it will call the function CopyLicenses and then continue.
            $answer2 = $wshell.Popup("Do you want to save your licenses on your computer before continuing?", 0, "Caution", 48 + 4)
            if ($answer2 -eq 6) { CopyLicenses }

            elseif ($answer2 -ne 6) { $Outputbox.Appendtext( "Action canceled: Copy Licenses`r`n" ) }

            $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
            Get-ItemProperty | Where-Object { 
                $_.Displayname -match "Tobii Dynavox Communicator" -or
                $_.Displayname -match "Sono" -or
                $_.Displayname -match "LiterAACy" -or
                $_.Displayname -match "SymbolStix" -or
                $_.Displayname -match "METACOM" -or
                $_.Displayname -match "PCS" -or
                $_.Displayname -match "Voices for Tobii Communicator"
            } | Select-Object Publisher, Displayname, UninstallString

            ForEach ($ver in $TobiiVer) {
                $Uninstname = $ver.Displayname
                $Outputbox.Appendtext( "Removing - " + "$Uninstname`r`n" )
                $uninst = $ver.UninstallString -replace "msiexec.exe", "" -Replace "/I", "" -Replace "/X", "" -replace "/uninstall", ""
                $uninst = $uninst.Trim()
                start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
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
    )
    #Show the Form 
    $form.ShowDialog() | Out-Null 
 
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

#A12
Function DeleteEmailsC5 {
    $outputBox.clear()
    $outputBox.appendtext( "running Delete Emails Communicator.exe...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "Delete Emails Communicator.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        Start-Process "Delete Emails Communicator.exe"
    }
    else { 
        $outputbox.appendtext("File handle.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
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

    $outputBox.AppendText( "Done Copy Licenses!`r`n" )
    Return
}

Function Write-Log {
    #https://www.script-example.com/en-powershell-logging
    Param ($Message, $filename)
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    Set-Location $fpath
    "$($Message)" | out-file "$fpath\$filename.txt" -Append
    $OutputBox.AppendText("$Message" + "`r`n" )
}

#B1
#TODO Sean to fix total cpu and memory usage in percent
Function Listapps {
    $Outputbox.clear()
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path

    #Gettings all Tobii SW versions
    $Listapps = Get-ChildItem -Recurse -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, 
				HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\, 
				HKLM:\Software\WOW6432Node\Tobii\ | Get-ItemProperty | Where-Object { 
        $_.Publisher -like '*Tobii*' -or 
        $_.Displayname -like '*Tobii Experience Software*' -or
        $_.Displayname -like '*Tobii Device Drivers*' -or
        $_.Displayname -like '*Tobii Eye Tracking For Windows*'
				} | Select-Object Displayname, Displayversion | Sort-Object Displayname | format-table -HideTableHeaders | out-string
    
    #Gettings all TD windows app versions
    $Listwindowsapp = Get-AppxPackage | Where-Object { ($_.Publisher -like '*Tobii*') -or ($_.Name -like '*Snap*') } | Select-Object name , version | format-table -HideTableHeaders | out-string

    #Getting SW for TT components
    if (Test-path "C:\Program Files\Tobii\Tobii EyeX") { 
        Set-Location "C:\Program Files\Tobii\Tobii EyeX"
        $TTComponents = Get-childitem * -include platform_runtime_IS5GIBBONGAZE_service.exe, InstallerPackageRemovalTool.exe, Tobii.Configuration.exe, Tobii.EyeX.Engine.exe, Tobii.EyeX.Interaction.exe, Tobii.Service.exe, tobii_stream_engine.dll  | foreach-object { "{0}`t{1}" -f $_.Name, [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion }
    }

    #Getting FW version
    $fpathfw = Get-ChildItem -Path $PSScriptRoot -Filter "Tdx.EyeTrackerInfo.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpathfw.count -gt 0) {
        Set-Location $fpathfw
        $ETSN = .\Tdx.EyeTrackerInfo.exe --serialnumber
        $ETFWV = .\Tdx.EyeTrackerInfo.exe --firmwareversion
        $ETModel = .\Tdx.EyeTrackerInfo.exe --model
        $ETPDKV = .\Tdx.EyeTrackerInfo.exe --runtimebuildversion
    }

    $GetProcess = Get-process "*GazeSelection*", "*Tobii*", "*Tdx*" | Select-Object Processname | Format-table -hidetableheaders | Out-string
    $GetServices = Get-Service -Name '*Tobii*', '*Tdx*' | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
	
    $regOEMs = @("HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation", "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation\", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\OEMInformation\")
    foreach ($regOEM in $regOEMs) {
        if (Test-Path $regOEM) {
            $TobiiOEMETModel = (Get-ItemProperty -Path $regOEM).EyeTrackerModel
            $TobiiOEMProductType = (Get-ItemProperty -Path $regOEM).ProductType

            $referenceOEMModel = "PCEye5", "EM12", "PCEye2", "PCEyeExplore" , "I-1\d\+"
            $referenceOEMProductType = "TDG13", "TDG16", "TDH10", "TDTW7", "TDG10" , "I-Series"

            if ($TobiiOEMETModel.count -gt 0) {
                $Compares1 = (Compare-Object -DifferenceObject $TobiiOEMETModel -ReferenceObject $referenceOEMModel -CaseSensitive -ExcludeDifferent -IncludeEqual | Select-Object InputObject).InputObject
            }
            elseif ($TobiiOEMProductType.count -gt 0) {
                $Compares2 = (Compare-Object -DifferenceObject $TobiiOEMProductType -ReferenceObject $referenceOEMProductType -CaseSensitive -ExcludeDifferent -IncludeEqual | Select-Object InputObject).InputObject
            }     
        }
    }
	
    $ETDrivers = Get-WmiObject Win32_PnPSignedDriver | Where-Object { $_.Manufacturer -match "Tobii" } | Select-Object DeviceName, DriverVersion
    if ($ETDrivers.Count -gt 0) {
        ForEach ($ETDriver in $ETDrivers) {
            $ETdrivername = $ETDriver.DeviceName
            $ETdriverversion = $ETDriver.DriverVersion
        }
    }

    $Outputbox.Appendtext( "TOBII INSTALLED SOFTWARE:$Listapps`r`n") 
    $Outputbox.Appendtext( "TOBII WINDOWS STORE APPS:$Listwindowsapp")
    $Outputbox.Appendtext( "$TTComponent") 
    $Outputbox.Appendtext( "ET S/N: $ETSN`r`nET FW version: $ETFWV`r`nET Model: $ETModel `r`nPDK version: $ETPDKV `r`n") 
    $Outputbox.Appendtext( "`r`nACTIVE PROCESSES:$GetProcess`r`n") 
    $Outputbox.Appendtext( "`r`nACTIVE Services:$GetServices`r`n")
    if ($Compares1) { $Outputbox.Appendtext( "In $regOEM EyeTrackerModel is $Compares1`r`n") }
    if ($Compares2) { $Outputbox.Appendtext( "In $regOEM ProductType is $Compares2`r`n") }
    $Outputbox.Appendtext( "$ETdrivername $ETdriverversion")
	
    #If first answer equals yes or no
    $answer1 = $wshell.Popup("Do you wish to continue and save info to folder?`r`n", 0, "Caution", 48 + 4)
    if ($answer1 -eq 6) {
        $Outputbox.Appendtext( "Continue..`r`n" )
    }
    elseif ($answer1 -ne 6) {
        $Outputbox.Appendtext( "Action canceled`r`n" )
        Return
    }
	
    "TOBII INSTALLED SOFTWARE:$Listapps`r`n" | out-file "$fpath\SoftwareVersions.txt"
    "TOBII WINDOWS STORE APPS:$Listwindowsapp" | out-file "$fpath\SoftwareVersions.txt" -Append
    "$TTComponent" | out-file "$fpath\SoftwareVersions.txt" -Append
    "ET S/N: $ETSN`r`nET FW version: $ETFWV`r`nET Model: $ETModel `r`nPDK version: $ETPDKV `r`n" | out-file "$fpath\SoftwareVersions.txt" -Append
    "`r`nACTIVE PROCESSES:$GetProcess`r`n" | out-file "$fpath\SoftwareVersions.txt" -Append
    "`r`nACTIVE Services:$GetServices`r`n" | out-file "$fpath\SoftwareVersions.txt" -Append
    if ($Compares1) { "In $regOEM EyeTrackerModel is $Compares1`r`n" | out-file "$fpath\SoftwareVersions.txt" -Append }
    if ($Compares2) { "In $regOEM ProductType is $Compares2`r`n" | out-file "$fpath\SoftwareVersions.txt" -Append }
    "$ETdrivername $ETdriverversion" | out-file "$fpath\SoftwareVersions.txt" -Append
	
    $TobiiDrivers = Get-WindowsDriver -Online | Where-Object { $_.ProviderName -match "Tobii" }  | Select-Object Driver , OriginalFileName
    if ($TobiiDrivers.count -gt 0) {
        ForEach ($Tobiidrivers in $TobiiDrivers) {
            $Tobiiinf = $Tobiidrivers.Driver 
            $TobiiList = $Tobiidrivers.OriginalFileName
            $TobiiList = $TobiiList.Replace("C:\Windows\System32\DriverStore\FileRepository\", "")
            Write-Log -Message "$Tobiiinf : $TobiiList"  -filename "SoftwareVersions"
        }
    }

    $fpathusb = Get-ChildItem -Path $PSScriptRoot -Filter "CastorUsbCli.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if (test-path $fpathusb) {
        Set-Location $fpathusb
        Get-Service -Name '*TobiiIS*' | Stop-Service -Force -passthru -ErrorAction ignore

        .\CastorUsbCli.exe "--unit-info" | Out-File "$fpath\SoftwareVersions.txt" -Append
        .\CastorUsbCli.exe "--list" | Out-File "$fpath\SoftwareVersions.txt" -Append
        .\CastorUsbCli.exe "--platform" | Out-File "$fpath\SoftwareVersions.txt" -Append

        Get-Service -Name '*TobiiIS*' | start-Service  -passthru -ErrorAction ignore
    }

    pnputil /enum-drivers > "$fpath\systemDrivers.txt"

    $key = 'HKLM:\SOFTWARE\WOW6432Node\Tobii Dynavox\Device' 
    if (Test-Path $key) {
        $SerialNumber = (Get-ItemProperty -Path $key)."Serial Number" 
        $OEMImage = (Get-ItemProperty -Path $key)."OEM Image"
        $ProductKey = (Get-ItemProperty -Path $key)."Product Key" 
    }
    else {
        $SerialNumber = (Get-CimInstance -ClassName Win32_bios).SerialNumber
        $Model = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
    }

    
    powercfg /batteryreport /output "$infofolder\$SerialNumber-battery-report.html"

    $DesignedCapacity = (Get-WmiObject -Class BatteryStaticData -Namespace ROOT\WMI).DesignedCapacity / 1000
    $FullChargedCapacity = (Get-WmiObject -Class BatteryFullChargedCapacity -Namespace ROOT\WMI).FullChargedCapacity / 1000 
    $BatteryHealth = [Math]::Round($FullChargedCapacity / $DesignedCapacity * 100)



    Write-Log -Message "Serial Number: $SerialNumber`r`n" -filename "SoftwareVersions"
    if ($OEMImage) { Write-Log -Message "OEM Image is $OEMImage`r`n" -filename "SoftwareVersions" }
    if ($OEMImage) { Write-Log -Message "Product Key is $ProductKey`r`n" -filename "SoftwareVersions" }
    if ($OEMImage) { Write-Log -Message "Model $Model`r`n" -filename "SoftwareVersions" }

    Write-Log -Message "Battery Designed Capacity is $DesignedCapacity mWh" -filename "SoftwareVersions"
    Write-Log -Message "Battery Full Charged Capacity is $FullChargedCapacity mWh" -filename "SoftwareVersions"
    Write-Log -Message "Battery Health is $BatteryHealth %`r`n" -filename "SoftwareVersions"
    
    #Getting installed .NET version
    $installedFrameworks = @()
    if (IsKeyPresent "HKLM:\Software\Microsoft\.NETFramework\Policy\v1.0" "3705") { 
        $installedFrameworks += Write-Log -Message "Installed .Net Framework 1.0" -filename "SoftwareVersions"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v1.1.4322" "Install") { 
        $installedFrameworks += Write-Log -Message "Installed .Net Framework 1.0" -filename "SoftwareVersions"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v2.0.50727" "Install") { 
        $installedFrameworks += Write-Log -Message "Installed .Net Framework 2.0" -filename "SoftwareVersions"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.0\Setup" "InstallSuccess") {
        $installedFrameworks += Write-Log -Message "Installed .Net Framework 3.0" -filename "SoftwareVersions"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.5" "Install") { 
        $installedFrameworks += Write-Log -Message "Installed .Net Framework 3.5" -filename "SoftwareVersions"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client" "Install") { 
        $installedFrameworks += Write-Log -Message "Installed .Net Framework 4.0c" -filename "SoftwareVersions"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Install") {
        $installedFrameworks += Write-Log -Message "Installed .Net Framework 4.0" -filename "SoftwareVersions"
    }

    $result = -1
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client" "Install" -or IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Install") {
        # .net 4.0 is installed
        $result = 0
        $version = GetFrameworkValue "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Release"
        
        if ($version -ge 528040 -Or $version -ge 528372 -Or $version -ge 528049) {
            # .net 4.8
            Write-Log -Message "Installed .Net Framework 4.8"  -filename "SoftwareVersions"
            $result = 10
        }
        elseif ($version -ge 461808 -Or $version -ge 461814) {
            # .net 4.7.2
            Write-Log -Message "Installed .Net Framework 4.7.2" -filename "SoftwareVersions"
            $result = 9
        }
        elseif ($version -ge 461308 -Or $version -ge 461310) {
            # .net 4.7.1
            Write-Log -Message "Installed .Net Framework 4.7.1" -filename "SoftwareVersions"
            $result = 8
        }
        elseif ($version -ge 460798 -Or $version -ge 460805) {
            # .net 4.7
            Write-Log -Message "Installed .Net Framework 4.7" -filename "SoftwareVersions"
            $result = 7
        }
        elseif ($version -ge 394802 -Or $version -ge 394806) {
            # .net 4.6.2
            Write-Log -Message "Installed .Net Framework 4.6.2" -filename "SoftwareVersions"
            $result = 6
        }
        elseif ($version -ge 394254 -Or $version -ge 394271) {
            # .net 4.6.1
            Write-Log -Message "Installed .Net Framework 4.6.1" -filename "SoftwareVersions"
            $result = 5
        }
        elseif ($version -ge 393295 -Or $version -ge 393297) {
            # .net 4.6
            Write-Log -Message "Installed .Net Framework 4.6" -filename "SoftwareVersions"
            Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.6"
            $result = 4
        }
        elseif ($version -ge 379893) {
            # .net 4.5.2
            Write-Log -Message "Installed .Net Framework 4.5.2" -filename "SoftwareVersions"
            $result = 3
        }
        elseif ($version -ge 378675) {
            # .net 4.5.1
            Write-Log -Message "Installed .Net Framework 4.5.1" -filename "SoftwareVersions"
            $result = 2
        }
        elseif ($version -ge 378389) {
            # .net 4.5
            Write-Log -Message "Installed .Net Framework 4.5" -filename "SoftwareVersions"
            $result = 1
        }   
    
        $outputbox.appendtext("Done! `r`n")
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

#B2
Function GetServices {
    $outputBox.clear()
    $GetProcess = Get-process "*GazeSelection*", "*Tobii*", "*Tdx*" | Select-Object Processname | Format-table -hidetableheaders | Out-string
    $GetServices = Get-Service -Name '*Tobii*', '*Tdx*' | Select-Object Name, Status | Format-table -hidetableheaders | Out-string

    $outputBox.appendtext( "Listing active Tobii processes...`r`n" )
    if ($GetProcess) {
        $outputbox.appendtext("ACTIVE PROCESSES:$GetProcess`r`n")
    }
    if ($GetServices) {
        $outputbox.appendtext("ACTIVE Services:$GetServices`r`n")
    }

    $outputbox.appendtext("Done!`r`n")
}

#B3
Function RestartProcesses {
    
    Add-Type -AssemblyName System.Windows.Forms    
    Add-Type -AssemblyName System.Drawing

    # Build Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Restart Services"
    $Form.Size = New-Object System.Drawing.Size(400, 400)
    $Form.StartPosition = "CenterScreen"
    #$Form.Topmost = $True

    # Add Button1
    $Button1 = New-Object System.Windows.Forms.Button
    $Button1.Location = New-Object System.Drawing.Size(40, 40)
    $Button1.Size = New-Object System.Drawing.Size(150, 50)
    $Button1.Text = "Restart ET Services"
    $Form.Controls.Add($Button1)
    
    # Add Button2
    $Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Size(190, 40)
    $Button2.Size = New-Object System.Drawing.Size(150, 50)
    $Button2.Text = "Restart TD Browse"
    $Form.Controls.Add($Button2)
    
    # Add Button3
    $Button3 = New-Object System.Windows.Forms.Button
    $Button3.Location = New-Object System.Drawing.Size(40, 90)
    $Button3.Size = New-Object System.Drawing.Size(150, 50)
    $Button3.Text = "Restart TD Control"
    $Form.Controls.Add($Button3)    
    
    # Add Button4
    $Button4 = New-Object System.Windows.Forms.Button
    $Button4.Location = New-Object System.Drawing.Size(190, 90)
    $Button4.Size = New-Object System.Drawing.Size(150, 50)
    $Button4.Text = "Restart TD Phone"
    $Form.Controls.Add($Button4)

    # Add Button5
    $Button5 = New-Object System.Windows.Forms.Button
    $Button5.Location = New-Object System.Drawing.Size(40, 140)
    $Button5.Size = New-Object System.Drawing.Size(150, 50)
    $Button5.Text = "Restart TD Talk"
    $Form.Controls.Add($Button5)

    #Add Button event 
    $Button1.Add_Click( {
            $outputBox.clear()
            $Outputbox.Appendtext( "Restart Services...`r`n")
            $StopServices = Get-Service -Name '*Tobii*' | Stop-Service -force -Passthru -erroraction ignore | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            $Outputbox.Appendtext( "Stopping following Services:$StopServices`r`n")

            Start-Sleep -s 3
            $Processkill = get-process "GazeSelection" , "*TobiiDynavox*", "*Tobii.EyeX*", "Notifier" -erroraction ignore | Stop-process -force -Passthru -erroraction ignore | Select-Object Processname | Format-table -Hidetableheaders | Out-string

            $Outputbox.Appendtext( "Stopping following processes:$Processkill`r`n")

            #start all processes and services
            Start-Sleep -s 3
            try {
                Start-Service -Name '*Tobii*' -ErrorAction Stop | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
                Start-process "C:\Program Files (x86)\Tobii Dynavox\Eye Assist\TobiiDynavox.EyeAssist.Engine.exe"
            }
            Catch {
                $Outputbox.Appendtext( "Failed to start!`r`n" )
            }
            Start-Sleep -s 5
            $StopServices = Get-Service -Name '*Tobii*' 
            $ProcessNames = Get-process "GazeSelection" , "*TobiiDynavox*", "*Tobii.EyeX*", "Notifier" -erroraction ignore | Select-Object Processname | Format-table -Hidetableheaders | Out-string

            $Outputbox.Appendtext( "Running Services:$StopServices`r`n" )
            Foreach ($ProcessName in $ProcessNames) {
                $Outputbox.Appendtext( "Running Processes:$ProcessName`r`n" )
            }
            $outputBox.Appendtext( "Done!`r`n" )
        }
    )

    $Button2.Add_Click( {
            $outputBox.clear()
            $Outputbox.Appendtext( "Restart Services...`r`n")
            $StopServices = Get-Service -Name '*Tdx.Browse*' | Stop-Service -force -Passthru -erroraction ignore | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            $Outputbox.Appendtext( "Stopping following Services:$StopServices`r`n")

            Start-Sleep -s 3
            try {
                Start-Service -Name '*Tdx.Browse*' -ErrorAction Stop | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            }
            Catch {
                $Outputbox.Appendtext( "Failed to start!`r`n" )
            }
            Start-Sleep -s 5
            $StopServices = Get-Service -Name '*Tdx.Browse*' | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            $ProcessNames = Get-process '*Tdx.Browse*' -erroraction ignore | Select-Object Processname | Format-table -Hidetableheaders | Out-string
        
            Foreach ($StopService in $StopServices) {
                $Outputbox.Appendtext( "Running Services:$StopService`r`n" )
            }  
            Foreach ($ProcessName in $ProcessNames) {
                $Outputbox.Appendtext( "`r`nRunning Processes:$ProcessName`r`n" )
            }
            $outputBox.Appendtext( "Done!`r`n" )
        }
    )

    $Button3.Add_Click( {
            $outputBox.clear()
            $Outputbox.Appendtext( "Restart Services...`r`n")
            $StopServices = Get-Service -Name '*Tdx.ComputerControl*' | Stop-Service -force -Passthru -erroraction ignore | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            $Outputbox.Appendtext( "Stopping following Services:$StopServices`r`n")
        
            Start-Sleep -s 3
            $Processkill = get-process "*Tdx.ComputerControl*" -erroraction ignore | Stop-process -force -Passthru -erroraction ignore | Select-Object Processname | Format-table -Hidetableheaders | Out-string

            Start-Sleep -s 3
            try {
                Start-Service -Name '*Tdx.ComputerControl*' -ErrorAction Stop | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            }
            Catch {
                $Outputbox.Appendtext( "Failed to start!`r`n" )
            }
            Start-Sleep -s 5
            $StopServices = Get-Service -Name '*Tdx.ComputerControl*' | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            $ProcessNames = Get-process '*Tdx.ComputerControl*' -erroraction ignore | Select-Object Processname | Format-table -Hidetableheaders | Out-string
        
            Foreach ($StopService in $StopServices) {
                $Outputbox.Appendtext( "Running Services:$StopService`r`n" )
            }  
            Foreach ($ProcessName in $ProcessNames) {
                $Outputbox.Appendtext( "`r`nRunning Processes:$ProcessName`r`n" )
            }
            $outputBox.Appendtext( "Done!`r`n" )
        }
    )

    $Button4.Add_Click( {
            $outputBox.clear()
            $Outputbox.Appendtext( "Restart Services...`r`n")
            $StopServices = Get-Service -Name '*Tdx.Phone*' | Stop-Service -force -Passthru -erroraction ignore | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            $Outputbox.Appendtext( "Stopping following Services:$StopServices`r`n")
        
            Start-Sleep -s 3
            $Processkill = get-process "*Tdx.Phone*" -erroraction ignore | Stop-process -force -Passthru -erroraction ignore | Select-Object Processname | Format-table -Hidetableheaders | Out-string

            Start-Sleep -s 3
            try {
                Start-Service -Name '*Tdx.Phone*' -ErrorAction Stop | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            }
            Catch {
                $Outputbox.Appendtext( "Failed to start!`r`n" )
            }
            Start-Sleep -s 5
            $StopServices = Get-Service -Name '*Tdx.Phone*' | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            $ProcessNames = Get-process '*Tdx.Phone*' -erroraction ignore | Select-Object Processname | Format-table -Hidetableheaders | Out-string
        
            Foreach ($StopService in $StopServices) {
                $Outputbox.Appendtext( "Running Services:$StopService`r`n" )
            }  
            Foreach ($ProcessName in $ProcessNames) {
                $Outputbox.Appendtext( "`r`nRunning Processes:$ProcessName`r`n" )
            }
            $outputBox.Appendtext( "Done!`r`n" )
        }
    )

    $Button5.Add_Click( {
            $outputBox.clear()
            $Outputbox.Appendtext( "Restart Services...`r`n")
            $StopServices = Get-Service -Name '*Tdx.Talk*' | Stop-Service -force -Passthru -erroraction ignore | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            $Outputbox.Appendtext( "Stopping following Services:$StopServices`r`n")
        
            Start-Sleep -s 3
            $Processkill = get-process "*Tdx.Talk*" -erroraction ignore | Stop-process -force -Passthru -erroraction ignore | Select-Object Processname | Format-table -Hidetableheaders | Out-string

            Start-Sleep -s 3
            try {
                Start-Service -Name '*Tdx.Talk*' -ErrorAction Stop | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            }
            Catch {
                $Outputbox.Appendtext( "Failed to start!`r`n" )
            }
            Start-Sleep -s 5
            $StopServices = Get-Service -Name '*Tdx.Talk*' | Select-Object Name, Status | Format-table -hidetableheaders | Out-string
            $ProcessNames = Get-process '*Tdx.Talk*' -erroraction ignore | Select-Object Processname | Format-table -Hidetableheaders | Out-string
        
            Foreach ($StopService in $StopServices) {
                $Outputbox.Appendtext( "Running Services:$StopService`r`n" )
            }  
            Foreach ($ProcessName in $ProcessNames) {
                $Outputbox.Appendtext( "`r`nRunning Processes:$ProcessName`r`n" )
            }
            $outputBox.Appendtext( "Done!`r`n" )
        }
    )

    $form.ShowDialog() | Out-Null 

}

#B4
Function ETfw {
    $outputBox.clear()
    $outputBox.appendtext( "Checking Eye tracker Firmware...`r`n" )
    $fpathfw = Get-ChildItem -Path $PSScriptRoot -Filter "Tdx.EyeTrackerInfo.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpathfw.count -gt 0) {
        Set-Location $fpathfw
        $ETSN = .\Tdx.EyeTrackerInfo.exe --serialnumber
        $ETFWV = .\Tdx.EyeTrackerInfo.exe --firmwareversion
        $ETModel = .\Tdx.EyeTrackerInfo.exe --model
        if ($ETSN.count -gt 0) {
            $outputbox.appendtext("ET S/N: $ETSN`r`nET FW version: $ETFWV`r`nET Model: $ETModel`r`n")
            if ($ETModel -match "IS4") {
                $answer1 = $wshell.Popup("This will upgrade IS4 firmware.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
                if ($answer1 -eq 6) {
                    $Outputbox.Appendtext( "Starting upgrade... Do NOT close this window while it is in progress..`r`n" )
                }
                elseif ($answer1 -ne 6) {
                    $Outputbox.Appendtext( "Action canceled`r`n" )
                    Return
                }
                $path = "C:\Program Files (x86)\Tobii\Service"
                #If first answer equals yes or no
                if (Test-Path $path) {             
                    Set-Location -path $path
                    if ($ETModel -match "IS4_PCEYE_MINI") {
                        #PCEye Mini: tobii-ttp://PCE1M-010106010685
                        $outputbox.appendtext("Upgrading PCEye mini FW..`r`n")
                        $PCEyeMini = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4pceyemini_firmware_2.27.0-4014648.tobiipkg" --no-version-check
                        $outputbox.appendtext("$PCEyeMini`r`n")
                        $outputbox.appendtext("Upgrade is Done! `r`n")
                    }
                    elseif ($ETModel -match "IS4_Large_102") {
                        $outputbox.appendtext("Upgrading PCEye Plus FW..`r`n")
                        $PCEyePlus = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4large102_firmware_2.27.0-4014648.tobiipkg" --no-version-check
                        $outputbox.appendtext("$PCEyePlus`r`n")
                        $outputbox.appendtext("Upgrade is Done! `r`n")
                    }
                    elseif ($ETModel -match "IS4_Large_Peripheral") {
                        $outputbox.appendtext("Upgrading Tobii4C FW..`r`n")
                        $4C = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4largetobiiperipheral_firmware_2.27.0-4014648.tobiipkg" --no-version-check
                        $outputbox.appendtext("$4C`r`n")
                        $outputbox.appendtext("Upgrade is Done! `r`n")
                    }
                    elseif ($ETModel -match "IS4_Base_I-series") {
                        $outputbox.appendtext("Upgrading I-Series+ FW..`r`n")
                        $ISeries = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii Dynavox\Gaze Interaction\Eye Tracker Firmware Releases\IS4B1\is4iseriesb_firmware_2.9.0.tobiipkg" --no-version-check
                        $outputbox.appendtext("$ISeries`r`n")
                        $outputbox.appendtext("Upgrade is Done. Restart ET through Control Center `r`n")
                    }
                    elseif ($ETModel -match "tet-tcp") {
                        #Tobii Firmware Upgrade Tool Automatically selected eye tracker tet-tcp://172.28.195.1 Failed to open file
                        $outputbox.appendtext("ET model is IS20. Use ET Browser to upgrade. Make sure that Bonjure is installed.`r`n")
                    }
                    #Get-Service -Name 'Tobii Service'  | Where-Object { $_.Status -ne "Running" } | Start-Service
                    Get-Service -Name 'Tobii Service'  | stop-service 
                    Get-Service -Name 'Tobii Service'  | Start-Service
                }
            } 
        }
        else {
            $outputbox.appendtext("Could not read ET SN!`r`n" )
        }
    } 
    else {
        $outputbox.appendtext("File Tdx.EyeTrackerInfo.exe is missing!`r`n")
    }
    $outputbox.appendtext("Done! `r`n")
} 

#B5
Function WCF {
    $outputBox.clear()
    $outputBox.appendtext( "Checking WCF Endpoint Blocking Software...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "handle.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        Start-Process cmd "/c  `"handle.exe net.pipe & pause `""
    }
    else { 
        $outputbox.appendtext("File handle.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B6
Function SMBios {
    Add-Type -AssemblyName System.Windows.Forms    
    Add-Type -AssemblyName System.Drawing

    # Build Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "SMBIOS tool"
    $Form.Size = New-Object System.Drawing.Size(400, 400)
    $Form.StartPosition = "CenterScreen"
    #$Form.Topmost = $True

    # Add Button1
    $Button1 = New-Object System.Windows.Forms.Button
    $Button1.Location = New-Object System.Drawing.Size(40, 40)
    $Button1.Size = New-Object System.Drawing.Size(150, 50)
    $Button1.Text = "getSMBIOSvalue"
    $Form.Controls.Add($Button1)
    
    # Add Button2
    $Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Size(190, 40)
    $Button2.Size = New-Object System.Drawing.Size(150, 50)
    $Button2.Text = "setName"
    $Form.Controls.Add($Button2)
    
    # Add Button3
    $Button3 = New-Object System.Windows.Forms.Button
    $Button3.Location = New-Object System.Drawing.Size(40, 90)
    $Button3.Size = New-Object System.Drawing.Size(150, 50)
    $Button3.Text = "setSerialNumber"
    $Form.Controls.Add($Button3)
    
    # Add Button4
    $Button4 = New-Object System.Windows.Forms.Button
    $Button4.Location = New-Object System.Drawing.Size(190, 90)
    $Button4.Size = New-Object System.Drawing.Size(150, 50)
    $Button4.Text = "setVendor"
    $Form.Controls.Add($Button4)
    
    # Add Button5
    $Button5 = New-Object System.Windows.Forms.Button
    $Button5.Location = New-Object System.Drawing.Size(40, 140)
    $Button5.Size = New-Object System.Drawing.Size(150, 50)
    $Button5.Text = "GB2SmbiosTool"
    $Form.Controls.Add($Button5)
    
    $outputBox.clear()
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "getSMBIOSvalues.cmd" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        
        #Add Button event 
        $Button1.Add_Click( {
                Start-Process -FilePath .\getSMBIOSvalues.cmd
            }
        )
        $Button2.Add_Click( {
                Start-Process -FilePath .\setName.cmd
            }
        )
        $Button3.Add_Click( {
                Start-Process -FilePath .\setSerialNumber.cmd
            }
        )
        $Button4.Add_Click( {
                Start-Process -FilePath .\setVendor.cmd
            }
        )
        $Button5.Add_Click( {
                $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "GB2SmbiosTool.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
                if ($fpath.count -gt 0) {
                    Set-Location $fpath
                    Start-Process cmd -Verb runAs "/c `"GB2SmbiosTool.exe`"" 
                }
                else { 
                    $outputbox.appendtext("File GB2SmbiosTool.exe is missing!`r`n" )
                }  
            }
        )

        #Show the Form 
        $form.ShowDialog() | Out-Null 
    }
    else { 
        $outputbox.appendtext("File getSMBIOSvalues.cmd is missing!`r`n" )
    }
    $outputbox.appendtext("Done!`r`n")
}

#B7 
Function IRUtility {
    $outputBox.clear()
    $outputBox.appendtext( "Opening TD IR Utility...`r`n" )
    $fpath = "C:\Program Files (x86)\Tobii Dynavox\Hardware Test Utility"
    if ($fpath.count -gt 0) {
        $fpath = (Get-ChildItem -Path "C:\Program Files (x86)\Tobii Dynavox\Hardware Test Utility" -Filter "TobiiDynavox.IRUtility.exe" -Recurse).FullName | Split-Path 
        Set-Location $fpath
        Start-Process .\TobiiDynavox.IRUtility.exe
    }
    else { 
        $outputbox.appendtext("File TobiiDynavox.IRUtility.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B8
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

    $fpathfw = Get-ChildItem -Path $PSScriptRoot -Filter "Tdx.EyeTrackerInfo.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpathfw.count -gt 0) {
        $outputbox.appendtext("Pinging ET..`r`n")
        Set-Location $fpathfw
        $ETSN = .\Tdx.EyeTrackerInfo.exe --serialnumber
        $ETFWV = .\Tdx.EyeTrackerInfo.exe --firmwareversion
        $ETModel = .\Tdx.EyeTrackerInfo.exe --model
        if ($ETSN.count -gt 0) {
            $outputbox.appendtext("ET S/N: $ETSN`r`nET FW version: $ETFWV`r`nET Model: $ETModel`r`n")
        }
        elseif ($ETSN.count -eq 0) {
            $outputbox.appendtext("No eye tracker could be found`r`n")
        }
    }
    else {
        $outputbox.appendtext("File Tdx.EyeTrackerInfo.exe is missing!`r`n" )
    }   
    $serviceNames = @("Tobii Service", "TobiiIS5LARGEPCEYE5", "TobiiIS5GIBBON", "TobiiGeneric")
    foreach ($serviceName in $serviceNames) {
        If (Get-Service $serviceName -ErrorAction SilentlyContinue) {
            If ((Get-Service $serviceName).Status -eq 'Running') {
                Stop-Service $serviceName
                $outputbox.appendtext("Stopping $serviceName`r`n")
            }
            else {
                $outputbox.appendtext("$serviceName found, but it is not running.`r`n")
            }
        }
        else {
            $outputbox.appendtext("$serviceName not found`r`n")
        }
    }
    Try {
        $outputBox.appendtext( "reseting is5 to bootloader...`r`n" )
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "CastorUsbCli.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        if ($fpath.count -gt 0) {
            Set-Location $fpath
            .\CastorUsbCli.exe --reset BOOT
        }
        else { 
            $outputbox.appendtext("File CastorUsbCli.exe is missing!`r`n" )
        }
        $getPID = Get-WmiObject Win32_PnPSignedDriver | Where-Object devicename -Like "*WinUSB Device*" | Select-Object DeviceID
        #$getdeviceids2 = Get-CimInstance Win32_PnPSignedDriver | Where-Object Description -Like "*WinUSB Device*" | Select-Object DeviceID
  
        if ($getPID) {
            FOREACH ($getPIDs in  $getPID) {
                $outputbox.appendtext("The reset is done. ET PID is now:$getPIDs`r`n")
            }
        }
        else {
            $outputbox.appendtext("Not able to read PID`r`n")
        }
    }
    Catch [System.Management.Automation.RemoteException] {
        $outputbox.appendtext("No Eye Tracker Connected`r`n")
    }
    #If second answer equals yes or no
    $answer2 = $wshell.Popup("Verify Hardware ID for EyeChip is set to 102", 0, "Caution", 48 + 4)
    if ($answer2 -eq 6) {

        foreach ($serviceName in $serviceNames) {
            if (Get-Service $serviceName -ErrorAction SilentlyContinue) {

                if ((Get-Service $serviceName).Status -ne 'Running') {
                    start-Service $serviceName
                    $outputbox.appendtext("Starting $serviceName`r`n")
                }
                else {
                    $outputbox.appendtext("$serviceName found, running.`r`n")
                }
            }
            else {
                $outputbox.appendtext("$serviceName not found`r`n")
            }
        }
    }     
    elseif ($answer2 -ne 6) {

    }
    $outputbox.appendtext("Done! `r`n")
}

#B9
Function RetrieveUnreleased {
    Add-Type -AssemblyName System.Windows.Forms    
    Add-Type -AssemblyName System.Drawing

    # Build Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Retrieve Unreleased UN"
    $Form.Size = New-Object System.Drawing.Size(500, 300)
    $Form.StartPosition = "CenterScreen"
    #$Form.Topmost = $True

    # Add Button1
    $Button1 = New-Object System.Windows.Forms.Button
    $Button1.Location = New-Object System.Drawing.Size(50, 30)
    $Button1.Size = New-Object System.Drawing.Size(150, 50)
    $Button1.Text = "Set value to True"
    $Form.Controls.Add($Button1)
    
    # Add Button2
    $Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Size(50, 80)
    $Button2.Size = New-Object System.Drawing.Size(150, 50)
    $Button2.Text = "Set the value to False"
    $Form.Controls.Add($Button2)
    
    # Add Button3
    $Button3 = New-Object System.Windows.Forms.Button
    $Button3.Location = New-Object System.Drawing.Size(50, 130)
    $Button3.Size = New-Object System.Drawing.Size(150, 50)
    $Button3.Text = "remove the key"
    $Form.Controls.Add($Button3)    
    
    $outputBox.clear()
    $regpath = 'HKLM:\SOFTWARE\WOW6432Node\Tobii\Update Notifier'
    if (!(Test-Path $regpath)) {
        $regpath = 'HKLM:\SOFTWARE\WOW6432Node\Tobii\I-Series\Update Notifier'
    }
    
    $Check = Get-ItemProperty -Path "$regpath" -Name RetrieveUnreleasedVersions -ErrorAction SilentlyContinue


    #Add Button event 
    $Button1.Add_Click( {
            if ($Check) {
                Set-ItemProperty -Path "$regpath" -Name "RetrieveUnreleasedVersions" -Value 'True'
            }
            else {
                New-ItemProperty -Path "$regpath" -Name "RetrieveUnreleasedVersions" -PropertyType "String" -Value 'True'
            }
            $outputbox.appendtext("Value set to True`r`n")
        }
    )

    $Button2.Add_Click( {
            if ($Check) { 
                Set-ItemProperty -Path "$regpath" -Name "RetrieveUnreleasedVersions" -Value 'False'
            }
            else {
                New-ItemProperty -Path "$regpath" -Name "RetrieveUnreleasedVersions" -PropertyType "String" -Value 'False'
            }
            $outputbox.appendtext("Value set to False`r`n") 
        }
    )

    $Button3.Add_Click( {
            if ($Check) {
                Remove-ItemProperty -Path "$regpath" -Name "RetrieveUnreleasedVersions"
                $outputbox.appendtext("String has been removed`r`n")
            }
        }
    )

    $outputbox.appendtext("Done!`r`n")
    $form.ShowDialog() | Out-Null 

}

#B10
Function RemoveDrivers {
    $outputBox.clear()

    Add-Type -AssemblyName System.Windows.Forms    
    Add-Type -AssemblyName System.Drawing

    # Build Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Remove ET Drivers"
    $Form.Size = New-Object System.Drawing.Size(500, 300)
    $Form.StartPosition = "CenterScreen"
    #$Form.Topmost = $True

    # Add Button1
    $Button1 = New-Object System.Windows.Forms.Button
    $Button1.Location = New-Object System.Drawing.Size(50, 30)
    $Button1.Size = New-Object System.Drawing.Size(150, 50)
    $Button1.Text = "Tobii Win Drivers"
    $Form.Controls.Add($Button1)
    
    # Add Button2
    $Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Size(50, 80)
    $Button2.Size = New-Object System.Drawing.Size(150, 50)
    $Button2.Text = "Tobii pnp Devices"
    $Form.Controls.Add($Button2)
    
    # Add Button3
    $Button3 = New-Object System.Windows.Forms.Button
    $Button3.Location = New-Object System.Drawing.Size(50, 130)
    $Button3.Size = New-Object System.Drawing.Size(150, 50)
    $Button3.Text = "Tobii wmi Objects"
    $Form.Controls.Add($Button3) 

    #Add Button event 
    $Button1.Add_Click( {
            $TobiiDrivers = Get-WindowsDriver -Online | Where-Object { $_.ProviderName -match "Tobii" } | Select-Object Driver, OriginalFileName
            $NewTobiiDrivers = $TobiiDrivers.originalfilename -replace "C:", "" -replace "(?<=\\).+?(?=\\)", "" -replace "\\\\\\", "" 
        
            $Outputbox.Appendtext( "Tobii Drivers are:`r`n")
            Foreach ($NewTobiiDriver in $NewTobiiDrivers) {
                $Outputbox.Appendtext( "$NewTobiiDriver`r`n")
            }

            #If first answer equals yes or no
            $answer1 = $wshell.Popup("Are you sure you want to remove all those drivers?`r`n", 0, "Caution", 48 + 4)
            if ($answer1 -eq 6) {
                $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

            }
            elseif ($answer1 -ne 6) {
                $Outputbox.Appendtext( "Action canceled: Remove driver`r`n" )
                Return
            }
            $outputBox.appendtext( "Removing Tobii Drivers...`r`n" )
            ForEach ($ver in $TobiiDrivers) {
                $outputBox.appendtext( "Removing - " + "$ver`r`n" )
                pnputil /delete-driver $ver.Driver /force /uninstall
            }   

        }
    )

    $Button2.Add_Click( {

            $TobiiDevices = Get-PnpDevice | where-object { ($_.FriendlyName -like '*Tobii*') }  | Sort-object FriendlyName -desc  | Select-Object Status, FriendlyName, InstanceId
            #$unknownTobiiDevices = Get-PnpDevice | Where-Object {($_.FriendlyName -like '*Tobii*') -and ($_.Status -like '*Unknown*')} | Sort-object FriendlyName -desc  | Select-Object Status, FriendlyName , InstanceId

            $Outputbox.Appendtext( "Tobii Drivers are:`r`n")
            Foreach ($TobiiDevice in $TobiiDevices) {
                $Outputbox.Appendtext( "$TobiiDevice`r`n")
            }

            #If first answer equals yes or no
            $answer1 = $wshell.Popup("Are you sure you want to remove all drivers?`r`n", 0, "Caution", 48 + 4)
            if ($answer1 -eq 6) {
                $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

            }
            elseif ($answer1 -ne 6) {
                $Outputbox.Appendtext( "Action canceled: Remove driver`r`n" )
                Return
            }
            $outputBox.appendtext( "Removing Tobii Drivers...`r`n" )
        
            foreach ($TobiiDevice in $TobiiDevices) {
                &"pnputil" /remove-device $TobiiDevice.InstanceId 
            }
        
        }
    )
    
    $Button3.Add_Click( {
        
            $TobiiObjects = Get-WmiObject Win32_PnPSignedDriver | Where-Object { $_.Manufacturer -match "Tobii" } | Select-Object DeviceName, DriverVersion, InfName

            $Outputbox.Appendtext( "Tobii Drivers are:`r`n")
            Foreach ($TobiiObject in $TobiiObjects) {
                $Outputbox.Appendtext( "$TobiiObject`r`n")
            }

            #If first answer equals yes or no
            $answer1 = $wshell.Popup("Are you sure you want to remove all those drivers?`r`n", 0, "Caution", 48 + 4)
            if ($answer1 -eq 6) {
                $Outputbox.Appendtext( "Starting... Do NOT close this window while it is in progress.`r`n" )

            }
            elseif ($answer1 -ne 6) {
                $Outputbox.Appendtext( "Action canceled: Remove driver`r`n" )
                Return
            }
            $outputBox.appendtext( "Removing Tobii Drivers...`r`n" )
            ForEach ($ver in $TobiiObjects) {
                $outputBox.appendtext( "Removing - " + "$ver`r`n" )
                pnputil /delete-driver $ver.InfName /force /uninstall
            }     

        }
    )

    $form.ShowDialog() | Out-Null 
    $outputbox.appendtext("Done!`r`n")
}

#B11
Function LongTest {
    $outputBox.clear()
    Add-Type -AssemblyName System.Windows.Forms    
    Add-Type -AssemblyName System.Drawing

    # Build Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Long Running Test"
    $Form.Size = New-Object System.Drawing.Size(500, 300)
    $Form.StartPosition = "CenterScreen"
    #$Form.Topmost = $True

    # Add Button1
    $Button1 = New-Object System.Windows.Forms.Button
    $Button1.Location = New-Object System.Drawing.Size(50, 30)
    $Button1.Size = New-Object System.Drawing.Size(200, 50)
    $Button1.Text = "ET Connection"
    $Form.Controls.Add($Button1)
    
    # Add Button2
    $Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Size(250, 30)
    $Button2.Size = New-Object System.Drawing.Size(200, 50)
    $Button2.Text = "ET Profiles"
    $Form.Controls.Add($Button2)
    
    # Add Button3
    $Button3 = New-Object System.Windows.Forms.Button
    $Button3.Location = New-Object System.Drawing.Size(50, 80)
    $Button3.Size = New-Object System.Drawing.Size(200, 50)
    $Button3.Text = "TD-RI Sample"
    $Form.Controls.Add($Button3) 

    # Add Button4
    $Button4 = New-Object System.Windows.Forms.Button
    $Button4.Location = New-Object System.Drawing.Size(250, 80)
    $Button4.Size = New-Object System.Drawing.Size(200, 50)
    $Button4.Text = "Sleep test"
    $Form.Controls.Add($Button4)
    
    # Add Button5
    $Button5 = New-Object System.Windows.Forms.Button
    $Button5.Location = New-Object System.Drawing.Size(50, 130)
    $Button5.Size = New-Object System.Drawing.Size(200, 50)
    $Button5.Text = "USB Log"
    $Form.Controls.Add($Button5)
    
    # Add Button6
    $Button6 = New-Object System.Windows.Forms.Button
    $Button6.Location = New-Object System.Drawing.Size(250, 130)
    $Button6.Size = New-Object System.Drawing.Size(200, 50)
    $Button6.Text = "Battery test"
    $Form.Controls.Add($Button6) 

    # Add Button7
    $Button7 = New-Object System.Windows.Forms.Button
    $Button7.Location = New-Object System.Drawing.Size(50, 180)
    $Button7.Size = New-Object System.Drawing.Size(200, 50)
    $Button7.Text = "OLI"
    $Form.Controls.Add($Button7) 

    #Add Button event 
    $Button1.Add_Click( { ETConnection } )
    $Button2.Add_Click( { EAProfileCreation } )
    $Button3.Add_Click( { RISamples } )
    $Button4.Add_Click( { Sleeper } )
    $Button5.Add_Click( { USBLogView } )
    $Button6.Add_Click( { BatteryLog } )
    $Button6.Add_Click( { OLI } )


    $form.ShowDialog() | Out-Null 
    $outputbox.appendtext("Done!`r`n")
}

#B11_A
Function ETConnection {
    $outputBox.clear()
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    USBLogView
    if ($fpath.count -gt 0) {
        Set-Location $fpath
    }
    
    $ETConnection = "$fpath\ETConnection"
    if (!(Test-Path $ETConnection)) {
        New-Item -Path "$ETConnection" -ItemType Directory  
    }
    $infofolder = "$ETConnection\ETConnectionOutput.txt"
    if (!(Test-Path "$infofolder")) {
        New-Item -Path "$infofolder" -ItemType file 
    }
  
    $outputBox.appendtext( "Running ET connection check...`r`n" )
    $outputBox.appendtext( "Results of output will be stored in $infofolder...`r`n" )
    $a = 1
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'Loop'
    $msg = 'Enter number of loops:'
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

    Do {
        Start-sleep -s 1
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        if ($fpath.count -gt 0) {
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
        Add-content $infofolder $time, $getinfo
        $a
        $outputbox.appendtext("$getinfo`r`n")
        $a++
    } while ($a -le $b)

    #Copy logs into result folder
    $LogPath = "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox", "$ENV:ProgramData\Tobii Dynavox", "$ENV:USERPROFILE\AppData\Local\Tobii", "$ENV:ProgramData\Tobii"
    $files = Get-ChildItem -Path $LogPath -Recurse | Where-Object {
                                                    ($_.Name -eq 'EyeAssistEngine.log') -or
                                                    ($_.Name -eq 'EyeTrackingSettings.log') -or
                                                    ($_.Name -eq 'RegionInteraction.log') -or
                                                    ($_.Name -eq 'ServerLog.txt') -or
                                                    ($_.Name -eq 'InteractionLog.txt') -or
                                                    ($_.Name -eq 'ServiceLog.txt') -or
                                                    ($_.Name -eq 'pr_log0.txt') 
    } | Select-Object -expand Fullname
    foreach ($file in $files) {
        Copy-Item -Path $file -Destination $ETConnection -Recurse
    }
    $outputbox.appendtext("Done! `r`n")
}

#B11_B
Function EAProfileCreation {
    $outputBox.clear()
    Add-Type -AssemblyName System.Windows.Forms    
    Add-Type -AssemblyName System.Drawing

    # Build Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "ET Profile"
    $Form.Size = New-Object System.Drawing.Size(500, 300)
    $Form.StartPosition = "CenterScreen"
    #$Form.Topmost = $True

    # Add Button1
    $Button1 = New-Object System.Windows.Forms.Button
    $Button1.Location = New-Object System.Drawing.Size(50, 30)
    $Button1.Size = New-Object System.Drawing.Size(200, 50)
    $Button1.Text = "create profile based on default"
    $Form.Controls.Add($Button1)
    
    # Add Button2
    $Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Size(50, 80)
    $Button2.Size = New-Object System.Drawing.Size(200, 50)
    $Button2.Text = "create as many profiles and calibrate"
    $Form.Controls.Add($Button2)
    
    # Add Button3
    $Button3 = New-Object System.Windows.Forms.Button
    $Button3.Location = New-Object System.Drawing.Size(50, 130)
    $Button3.Size = New-Object System.Drawing.Size(200, 50)
    $Button3.Text = "remove all created profiles"
    $Form.Controls.Add($Button3) 

    #Add Button event 
    $Button1.Add_Click( { 
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
            $title = 'Profile Name'
            $msg = "Write a profile name"
            $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
            Set-Location 'C:\Program Files (x86)\Tobii Dynavox\Eye Assist'
            .\TobiiDynavox.EyeAssist.Smorgasbord.exe --createprofilewithdefaultcalibration --profile $b
            $outputbox.appendtext("Profile with $b has been created`r`n")
        } )
    $Button2.Add_Click( {
            $xml = [xml](Get-Content "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist\EyeAssist.Settings.xml")
            $xml.Settings.CalibrationStimulusSpeed = "Fast"
            $xml.Settings.EasyCalibrationNumberOfPoints = [string]2
            $xml.Save("$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist\EyeAssist.Settings.xml")
            
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
        } )
    $Button3.Add_Click( { 
            $Keys = ("HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\UserProfiles" )
            Remove-item $Keys -Recurse -ErrorAction Ignore
        } )

    $form.ShowDialog() | Out-Null 
    $outputbox.appendtext("Done!`r`n")

}

#B11_C
Function RISamples {
    #Copy trace file from C: to RISample folder.
    $outputBox.clear()
    $outputBox.appendtext( "Starting TD region interaction sample...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    $fpathsample = Get-ChildItem -Path $PSScriptRoot -Filter "Tdx.EyeTracking.RegionInteraction.EyeAssist.Sample.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    USBLogView
    $RISamples = "$fpath\RISamples"
    if (!(Test-Path $RISamples)) {
        New-Item -Path "$RISamples" -ItemType Directory  
    }
    $ConResults = "$RISamples\ETConnectionSample.txt"
    $SamResults = "$RISamples\SampleResults.txt"
    $ProcessList = @("Tdx.EyeTracking.RegionInteraction.EyeAssist.Sample" )
    $SampleLog = "C:\trace\tobii\Tdx.EyeTracking.RegionInteraction.EyeAssist.Sample.log"
    if (!(Test-Path "$ConResults") -or !(Test-Path "$SamResults")) {
        New-Item -Path "$ConResults" -ItemType file 
        New-Item -Path "$SamResults" -ItemType file 
    }
    if ($fpathsample.count -gt 0) {
        Set-Location $fpathsample
        .\Tdx.EyeTracking.RegionInteraction.EyeAssist.Sample.exe
    }
    else { 
        $outputbox.appendtext("File Tdx.EyeTracking.RegionInteraction.EyeAssist.Sample.exe is missing!`r`n")
    }

    Do {  
        $ProcessesFound = Get-Process | Where-Object { $ProcessList -contains $_.Name } | Select-Object -ExpandProperty Name
        If ($ProcessesFound) {
            $fpathfw = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
            If ($fpathfw.count -gt 0) {
                Set-Location $fpathfw
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
            Add-content $ConResults $time, $getinfo
            Start-Sleep 1
        }
    } Until (!$ProcessesFound)

    [datetime[]] $timestamps = @(Get-Content -path $SampleLog -raw | Select-String '\d{4}\-(0?[1-9]|1[012])\-(0?[1-9]|[12][0-9]|3[01])*\s(\d+:\d+:\d+)' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }) 

    if ($timestamps.Count -lt 2) {
        Write-Host "Only one result: " $timestamps[0]
        return
    }

    for ($i = 0; $i -lt $timestamps.Count; $i++) {
        $previous = $timestamps[$i]
        $current = $timestamps[$i + 1]
        $difference = ($current - $previous)

        if (($difference) -gt ("00:00:05")) {
            Add-Content "$SamResults" "Gap between $current and $previous with ($difference)`n"
        } 
    }
    $outputbox.appendtext("Results are saved in $fpath! `r`n")
    #Remove-Variable * -ErrorAction SilentlyContinue
    $LogPath = "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox", "$ENV:ProgramData\Tobii Dynavox", "$ENV:USERPROFILE\AppData\Local\Tobii", "$ENV:ProgramData\Tobii"

    $files = Get-ChildItem -Path $LogPath -Recurse | Where-Object {
                                                    ($_.Name -eq 'EyeAssistEngine.log') -or
                                                    ($_.Name -eq 'EyeTrackingSettings.log') -or
                                                    ($_.Name -eq 'RegionInteraction.log') -or
                                                    ($_.Name -eq 'ServerLog.txt') -or
                                                    ($_.Name -eq 'InteractionLog.txt') -or
                                                    ($_.Name -eq 'ServiceLog.txt') -or
                                                    ($_.Name -eq 'pr_log0.txt') 
    } | Select-Object -expand Fullname
    foreach ($file in $files) {
        Copy-Item -Path $file -Destination $RISamples -Recurse
    }
    $outputbox.appendtext("Done! `r`n")
}

#B11_D
Function Sleeper {
    $outputBox.clear()
    $outputBox.appendtext( "30,30,60,60, Only S3, Cycle forever, update log to file, update externa actions...`r`n" )
    $fpath = (Get-ChildItem -Path "$PSScriptRoot" -Filter "Sleeper.exe" -Recurse).FullName | Split-Path 
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        $files = "$fpath\sleeper.log", "$fpath\sleeperoutput.txt"
        foreach ($file in $files) {
            Clear-Content -Path "$file"
        }
        if (test-path "$fpath\sleeperCfg.dat") {
            remove-item $fpath\sleeperCfg.dat
        }
        start-process .\Sleeper.exe
    }
    else { 
        $outputbox.appendtext("File Sleeper.exe is missing!`r`n" )
    }	
    $outputbox.appendtext("Done! `r`n")
}

#B11_E
Function USBLogView {
    $outputBox.clear()
    $outputBox.appendtext( "running USBLogView.exe...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "USBLogView.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        Start-Process "USBLogView.exe"
    }
    else { 
        $outputbox.appendtext("File handle.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B11_F
Function BatteryLog {
    $outputBox.clear()
    $outputBox.appendtext( "Starting TobiiDynavox.QA.BatteryMonitor.exe...`r`n" )
    $fpath = (Get-ChildItem -Path "$PSScriptRoot" -Filter "TobiiDynavox.QA.BatteryMonitor.exe" -Recurse).FullName | Split-Path 
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        start-process .\TobiiDynavox.QA.BatteryMonitor.exe
        $outputbox.appendtext("Results will be saves in $fpath\battery_log.csv`r`n")
    }
    else { 
        $outputbox.appendtext("File BatteryMonitor.exe is missing!`r`n" )
    }	
    $outputbox.appendtext("Done! `r`n")
}

#B11_G
Function OLI {
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $a = 1
    $title = 'OLI'
    $msg = 'ENTER the name of OLI inc. path:'
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    
    #$offlineInstaller = "Tobii.I-Series.Offline.Installer_4.180.0.29190.msi"
	
    #Set-Location "C:\Users\Qa\Desktop\SupportTools"
    $regPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\"
	
    for ($i = 1; $i -le 5; $i++) {
        $tobiiVer = Get-ChildItem -Path $regPath | Get-ItemProperty | Where-Object { $_.Displayname -Match "Tobii Experience Software For Windows" } | Select-Object UninstallString
		
        if ($tobiiVer) {
            # Uninstall the software
            Read-Host "Press ENTER to uninstall Tobii Experience Software"
            $uninstallString = $tobiiVer.UninstallString -replace "msiexec.exe", "" -replace "/I", ""
            $uninstallProcess = Start-Process "msiexec.exe" -ArgumentList "/X $uninstallString /quiet /norestart" -Wait -PassThru
			
            if ($uninstallProcess.ExitCode -ne 0) {
                Write-Error "Error $uninstallProcess.ExitCode: Failed to uninstall Tobii Experience Software."
                break
            }
        }
		
        # Install the offline installer
        Read-Host "Press ENTER to install Tobii Experience Software"
        try {
            Start-Process "$b" -ArgumentList "/quiet" -Wait
            #Start-Process "$offlineInstaller" -ArgumentList "/quiet" -Wait

        }
        catch {
            Write-Error "Error: Failed to install Tobii Experience Software."
            break
        }
		
        # Output the iteration number to indicate progress
        Write-Output "Iteration: $i"
    }
}

#B12
#TODO
Function Logging {
    $outputBox.clear()
    Add-Type -AssemblyName System.Windows.Forms    
    Add-Type -AssemblyName System.Drawing

    # Build Form
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Logging"
    $Form.Size = New-Object System.Drawing.Size(500, 300)
    $Form.StartPosition = "CenterScreen"
    #$Form.Topmost = $True

    # Add Button1
    $Button1 = New-Object System.Windows.Forms.Button
    $Button1.Location = New-Object System.Drawing.Size(50, 30)
    $Button1.Size = New-Object System.Drawing.Size(150, 50)
    $Button1.Text = "MiddleWare debug"
    $Form.Controls.Add($Button1)
    
    # Add Button2
    $Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Size(50, 80)
    $Button2.Size = New-Object System.Drawing.Size(150, 50)
    $Button2.Text = "startIS5 Debug Tool"
    $Form.Controls.Add($Button2)

    # Add Button3
    $Button3 = New-Object System.Windows.Forms.Button
    $Button3.Location = New-Object System.Drawing.Size(50, 130)
    $Button3.Size = New-Object System.Drawing.Size(150, 50)
    $Button3.Text = "finish IS5 Debug Tool"
    $Form.Controls.Add($Button3)
        
    #Add Button event 
    $Button1.Add_Click( { DebugLoggning } )
    $Button2.Add_Click( { startIS5DebugTool } )
    $Button3.Add_Click( { finishIS5DebugTool } )

    $form.ShowDialog() | Out-Null 
    $outputbox.appendtext("Done!`r`n")
}

#B12_A
Function DebugLoggning {
    #Instructions
    #Start PowerShell as admin
    #Run "Set-ExecutionPolicy Unrestricted"
    #Run the attached script.
    #Tobii Service is then restarted and new log levels are set. If you are experiencing issues, please restart Tobii Service.
    #https://confluence.tobii.intra/pages/viewpage.action?spaceKey=EYEX&title=Changing+log+level
        


    param (    
        [switch]$reset,
        [ValidateSet( 'DEBUG', 'INFO', 'WARNING', 'ERROR', 'FATAL')]
        [string]$MinLevel = 'DEBUG',
        [ValidateSet( 'DEBUG', 'INFO', 'WARNING', 'ERROR', 'FATAL')] 
        [string]$MaxLevel = 'FATAL')
    
    if ($reset) {
        $MinLevel = 'INFO'
        $MaxLevel = 'ERROR'
        $OutputBox.AppendText( "Reset log levels to  + $MinLevel +  and  + $MaxLevel`r`n" )
        
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
            write-host "     $minLevelNode.Value = $MinLevel   "
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

Function SetDebugLogging {
    $outputBox.clear()
    
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'debug'
    $msg = 'Enter 1 for normal level or 2 for debug :'
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    if ($b -eq 1) {
        DebugLoggning -MinLevel INFO -MaxLevel ERROR
    }
    elseif ($b -eq 2) {
        DebugLoggning -MinLevel DEBUG -MaxLevel FATAL
    }
}

#B12_B
Function startIS5DebugTool {
    $outputBox.clear()
    $outputBox.appendtext( "Starting TobiiDeviceDebugTool.exe...`r`n1. Tobii Service shall be running.`r`n2. Do not close the window`r`n3. Run test scenario for which the user wants to collect the logs.`r`n4. When the test is done, stop the tool by pressing any key.`r`n5. collect the logs from $ENV:USERPROFILE\AppData\Local\tobiipdk\<platform_type>`r`n" )
    
    $fpath = (Get-ChildItem -Path "$PSScriptRoot" -Filter "TobiiDeviceDebugTool.exe" -Recurse).FullName | Split-Path
    $Tempfolder = Get-ChildItem -Path "$ENV:USERPROFILE\AppData\Local\Temp" -Filter TobiiDeviceDebugTool
    
    if ($fpath.count -gt 0) {
        if ($Tempfolder.count -eq 0) {
            Copy-Item -Path "$fpath" -Destination "$ENV:USERPROFILE\AppData\Local\Temp" -Recurse
        }
        Set-Location "$ENV:USERPROFILE\AppData\Local\Temp\TobiiDeviceDebugTool"
        Start-Process -FilePath "$ENV:USERPROFILE\AppData\Local\Temp\TobiiDeviceDebugTool\RunApp.bat"
        copy-Item -Path "$ENV:USERPROFILE\AppData\Local\Temp\TobiiDeviceDebugTool\RunApp.bat - Shortcut.lnk" -Destination "$ENV:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    }
    else { 
        $outputbox.appendtext("File TobiiDeviceDebugTool.exe is missing!`r`n" )
    }	
    $outputbox.appendtext("Done! `r`n")
}

#B12_C
#Todo check when the device restart
Function finishIS5DebugTool {
    $outputBox.clear()
    $outputBox.appendtext( "Finishing TobiiDeviceDebugTool.exe...`r`n" )
    
    $processid = (Get-Process -Name "TobiiDeviceDebugTool").Id

    $wshell = New-Object -ComObject wscript.shell
    $wshell.AppActivate($processid)
    $wshell.SendKeys("{ENTER}")
    $wshell.SendKeys("{ENTER}")

    Invoke-Item -Path "$ENV:USERPROFILE\AppData\Local\tobiipdk"
    $outputbox.appendtext("Done! `r`n")
}

#B13
Function InternalSE {
    $outputBox.clear()
    $outputBox.appendtext( "Starting Stream Engine Sample app...`r`n" )
    $fpath = (Get-ChildItem -Path "$PSScriptRoot" -Filter "sample.exe" -Recurse).FullName | Split-Path 
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        Start-Process .\sample.exe
    }
    else { 
        $outputbox.appendtext("File sample.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B14
Function Diagnostic {
    $outputBox.clear()
    $outputBox.appendtext( "Run diagnostics application for Interaction...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "Tobii.EyeX.Diagnostics.Application.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        start-process cmd "/c `"Tobii.EyeX.Diagnostics.Application.exe`""
    }
    else { 
        $outputbox.appendtext("File Tobii.EyeX.Diagnostics.Application.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B15 Deploy from v0.3
Function Deployment {
    #USB namings:
    #"Indi_AD"             "Indi_AB"             "Indi_A_XX"
    #"Indi_BD"             "Indi_BB"             "Indi_B_XX"
    #"Indi_7_AD"           "Indi_7_AB"           "Indi_7_A_XX"
    #"Indi_7_BD"           "Indi_7_BB"           "Indi_7_B_XX"
    #"I-110D"              "I-110B"              "I-110_XX"
    #"I-Series+D"          "I-SeriesB"           "I-Series+_XX"
    #"Surface_SP6D"        "Surface_SP6"         "Surface_Pro_SP6_XX"
    #"Surface_SP7D"        "Surface_SP7"         "Surface_Pro_SP7_XX"
    #"ISeries_IOTD"        "ISeries_IOT"         "ISeries_IOT"
    #"ISeriesD"            "ISeriesB"            "ISeries_XX"
    #"I-110-8W10D"         "I-110-8W10B"         "I-110-850_W10_REPAIR_XX"
    #"I-110-8W11D"         "I-110-8W11B"         "I-110-850_W11_XX"
    #"ISeries_MPD"         "ISeries_MPB"         ""

    $ReferenceD = @("Indi_A", "Indi_B", "Indi_7_A", "Indi_7_B", "I-110", "I-Series+", "Surface_Pro_SP6", 
        "ISeries", "Surface_Pro_SP7", "I-110-850_W11" , "I-110-850_W10", "ISeries_IOT",
        "ISeries_MP")

    $availableUSBs = (@(Get-Volume | Where-Object DriveType -eq Removable | Where-Object FileSystemType -eq NTFS |  Select-Object FileSystemLabel).FileSystemLabel ) -replace ".$"
    $outputbox.appendtext("Available USB drives: $availableUSBs`r`n")

    foreach ($availableUSB in $availableUSBs) {
        if ( $availableUSB -match "Surface_SP6") { $newReferenceDs = "Surface_Pro_SP6" }
        elseif ( $availableUSB -match "Surface_SP7") { $newReferenceDs = "Surface_Pro_SP7" }
        elseif ( $availableUSB -match "I-110-8W11") { $newReferenceDs = "I-110-850_W11" }
        elseif ( $availableUSB -match "I-110-8W10") { $newReferenceDs = "I-110-850_W10" }
        else { $newReferenceDs = (Compare-Object -DifferenceObject $ReferenceD -ReferenceObject $availableUSB -CaseSensitive -ExcludeDifferent -IncludeEqual | Select-Object InputObject).InputObject }

        $outputbox.appendtext("`r`nSelecting: $newReferenceDs`r`n")

        foreach ($newComparesUSBs in $newReferenceDs) {

            $DeployName = "$availableUSB" + "D"

            if ("$newComparesUSBs" -eq "I-Series+") {
                $BootName = $newComparesUSBs.Replace( "+", "") + "B"
            }
            elseif ("$newComparesUSBs" -eq "Surface_Pro_SP6") {
                $BootName = $newComparesUSBs.Replace("_Pro", "")
            }
            elseif ("$newComparesUSBs" -eq "Surface_Pro_SP7") {
                $BootName = $newComparesUSBs.Replace("_Pro", "")
            }
            elseif ("$newComparesUSBs" -eq "I-110-850_W11") {
                $BootName = $newComparesUSBs.Replace("50_", "") + "B"
            }
            elseif ("$newComparesUSBs" -eq "I-110-850_W10") {
                $BootName = $newComparesUSBs.Replace("50_", "") + "B"
            }
            elseif ("$newComparesUSBs" -eq "ISeries_IOT") {
                $BootName = $newComparesUSBs.Replace("B", "")
            }
            else {
                $BootName = "$newComparesUSBs" + "B"
            }
            $outputbox.appendtext("Setting deploy name to $DeployName and $BootName`r`n")
            $paths = "$env:USERPROFILE\Downloads"#, "D:\"

            $CheckDownload = (Get-ChildItem -Path $paths | Where-Object { $_.Name -match "$newComparesUSBs" -and $_.Name -match ".7z" }).Name 

            if ($CheckDownload -gt 1) {
                foreach ($CheckDownloads in $CheckDownload) {
                    #for I-SeriesIOT
                    if (($CheckDownloads -match "IOT") -and ($newComparesUSBs -eq "ISeries_IOT")) {
                        $Download = ((Get-ChildItem -Path $paths | Where-Object { $_.Name -eq "$CheckDownloads" }).Name ) -replace ".7z", ""
                        $Download2 = Get-ChildItem -Path $paths  | Where-Object { $_.Name -eq "$CheckDownloads" } | Select-Object -expand Fullname | Split-Path
                    }
                    elseif ($CheckDownloads -match "\+") {
                        $Download = ((Get-ChildItem -Path $paths | Where-Object { $_.Name -eq "$CheckDownloads" }).Name ) -replace ".7z", ""
                        $Download2 = Get-ChildItem -Path $paths  | Where-Object { $_.Name -eq "$CheckDownloads" } | Select-Object -expand Fullname | Split-Path
                    }
                    elseif (($CheckDownloads -notmatch "IOT") -and ($newComparesUSBs -eq "ISeries")) {
                        $Download = ((Get-ChildItem -Path $paths | Where-Object { $_.Name -eq "$CheckDownloads" }).Name ) -replace ".7z", ""
                        $Download2 = Get-ChildItem -Path $paths  | Where-Object { $_.Name -eq "$CheckDownloads" } | Select-Object -expand Fullname | Split-Path
                    }

                }
            }
            else {

                $Download = ((Get-ChildItem -Path $paths | Where-Object { $_.Name -match "$newComparesUSBs" }).Name ) -replace ".7z", ""
                $Download2 = Get-ChildItem -Path $paths  | Where-Object { $_.Name -match "$newComparesUSBs" } | Select-Object -expand Fullname | Split-Path
            }

            Set-Location  "C:\Program Files\7-Zip"
            if ($Download) {
                $outputbox.appendtext("Found deploy: $Download`r`n")

                # clear content in USB
                $outputbox.appendtext("Formatting both $DeployName and $BootName`r`n")
                $Test1 = Format-Volume -FriendlyName $DeployName -FileSystem NTFS -NewFileSystemLabel $DeployName
                $Test2 = Format-Volume -FriendlyName $BootName -FileSystem FAT32 -NewFileSystemLabel $BootName

                # Find and select driver for USB
                $getDepLetter = (Get-Volume | Where-Object { ($_.FileSystemLabel -eq "$DeployName") }).DriveLetter
                $getBootLetter = (Get-Volume | Where-Object { ($_.FileSystemLabel -eq "$BootName") }).DriveLetter
                $outputbox.appendtext("Found following driver letters: $getDepLetter & $getBootLetter`r`n")

                # Select correct deploy from Download
                $outputbox.appendtext("Unpacking...`r`n")
                $unpack = .\7z.exe x "$Download2\$Download.7z" -o"$getDepLetter":\  -p5rd4c5vgcTvuKC -r
                
                # Move files to its proper path
                $USBpath = "$getDepLetter':\'$Download" -replace "'", ""
                $newgetDepLetter = "$getDepLetter':\'" -replace "'", ""
                $newgetBootLetter = "$getBootLetter':\'" -replace "'", ""

                Get-ChildItem -Path "$USBpath\winpe" -Recurse | Move-Item -Destination $newgetBootLetter
                Get-ChildItem -Path "$USBpath\deploy" -Recurse | Move-Item -Destination $newgetDepLetter 
                Remove-Item -Path "$USBpath" -Force -Recurse
                $outputbox.appendtext("Moving folders to its right path and cleaning..`r`n")
            }
            else {
                $outputbox.appendtext("No Deploy match $newComparesUSBs`r`n")
            }
        }
    }
    $outputbox.appendtext("DONE`r`n")
}

#B16
Function LogCollector {
    #File version 
    $LogCollectorTool = "LogCollectorTool_V0.9"

    #Forces powershell to run as an admin
    if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
    { Start-Process powershell.exe "-NoProfile -Windowstyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
    $PSScriptRoot

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$LogCollectorTool"
    $form.Size = New-Object System.Drawing.Size(420, 320)
    $form.StartPosition = 'CenterScreen'

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(135, 250)
    $okButton.Size = New-Object System.Drawing.Size(75, 25)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(210, 250)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 25)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(75, 20)
    $label.Text = "Logs location:"
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(85, 10)
    $textBox.Size = New-Object System.Drawing.Size(270, 20)
    $form.Controls.Add($textBox)

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10, 30)
    $label2.Size = New-Object System.Drawing.Size(390, 20)
    $label2.Text = "                        ex. C:\Users\TobiiDynavox_SysInfo_xxxx"
    $form.Controls.Add($label2)

    $label3 = New-Object System.Windows.Forms.Label
    $label3.Location = New-Object System.Drawing.Point(10, 50)
    $label3.Size = New-Object System.Drawing.Size(240, 30)
    $label3.Text = "A. Choose one of following (only the number):"
    $form.Controls.Add($label3)

    $textBox2 = New-Object System.Windows.Forms.TextBox
    $textBox2.Location = New-Object System.Drawing.Point(250, 50)
    $textBox2.Size = New-Object System.Drawing.Size(60, 20)
    $form.Controls.Add($textBox2)

    $label3 = New-Object System.Windows.Forms.Label
    $label3.Location = New-Object System.Drawing.Point(10, 80)
    $label3.Size = New-Object System.Drawing.Size(350, 50)
    $label3.Text = "1- Latest logs                                 2- Eye Assist logs`n3- Driver software logs                  4- Driver installer logs`n5- Any other file or folder               6- Convert BW logs to Windows`n7- Timing Issue EventLog Finder  8- RI Sample results"
    $form.Controls.Add($label3)

    $label4 = New-Object System.Windows.Forms.Label
    $label4.Location = New-Object System.Drawing.Point(10, 140)
    $label4.Size = New-Object System.Drawing.Size(400, 20)
    $label4.Text = "B. Logs between two timestamps: format should be as: yyyy-mm-dd hh:mm"
    $form.Controls.Add($label4)

    $textBox3 = New-Object System.Windows.Forms.TextBox
    $textBox3.Location = New-Object System.Drawing.Point(10, 160)
    $textBox3.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($textBox3)

    $textBox4 = New-Object System.Windows.Forms.TextBox
    $textBox4.Location = New-Object System.Drawing.Point(180, 160)
    $textBox4.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($textBox4)

    $form.Topmost = $true

    $form.Add_Shown( { $textBox.Select() })
    $result = $form.ShowDialog()
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force

    if ($x -and $x2 -and $x3) {
        Clear-Variable x3
        Clear-Variable x2
        Clear-Variable x
    }

    Function LatestErrorLogs {
        if ($x) {
            $LogPath = $x
            $ErrorPath = "$LogPath\ErrorLogs"
        }
        else {
            $LogPath = "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox", "$ENV:ProgramData\Tobii Dynavox", "$ENV:USERPROFILE\AppData\Local\Tobii", "$ENV:ProgramData\Tobii"
            $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
            $ErrorPath = "$fpath\ErrorLogs"
        }    
        $files = Get-ChildItem -Path $LogPath -Recurse | Where-Object {
                                                    ($_.Name -match 'ComputerControl.log') -or
                                                    ($_.Name -match 'ComputerControl.Updater.log') -or
                                                    ($_.Name -match 'ComputerControl.Launcher.log') -or
                                                    ($_.Name -match 'EyeAssistEngine.log') -or
                                                    ($_.Name -match 'EyeTrackingSettings.log') -or
                                                    ($_.Name -match 'RegionInteraction.log') -or
                                                    ($_.Name -match 'Switcher.log') -or
                                                    ($_.Name -match 'Switcher.Updater.log') -or
                                                    ($_.Name -match 'ServerLog.txt') -or
                                                    ($_.Name -match 'InteractionLog.txt') -or
                                                    ($_.Name -match 'ConfigurationLog.txt') -or
                                                    ($_.Name -match 'ServiceLog.txt') -or
                                                    ($_.Name -match 'pr_log0.txt') -or
                                                    ($_.Name -match 'Updater.log') -or
                                                    ($_.Name -match 'Talk.Launcher.log')
        } | Select-Object -expand Fullname
 
   
        #Creating folder
        if (!(Test-Path "$ErrorPath")) {
            Write-Host "Creating ErrorLogs folder in $ErrorPath.."
            New-Item -Path "$ErrorPath" -ItemType Directory  
        }
        #Creating files
        if (!(Test-Path "$ErrorPath\LatestErrors.txt")) {
            New-Item -Path $ErrorPath -Name "LatestErrors.txt" -ItemType "file"
        }
        else {
            Clear-Content -Path "$ErrorPath\LatestErrors.txt"
        }

        foreach ($file in $files) {
            if (![System.IO.File]::Exists($file)) {
                Write-Host "file with path $file doesn't exist"
            }
            else {
                $test = New-Item -Path $ErrorPath -Name "temp.txt" -ItemType "file"
                Get-Content -Path "$file" -Raw | ForEach-Object -Process { $_ -replace "- `r`n", '- ' } | Add-Content -Path "$ErrorPath\temp.txt"
                #$content1 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "error" -AllMatches | ForEach-Object { $_.Line }
                $content1 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "error", "WixRemoveFoldersEx" -AllMatches | select-string -pattern 'NO_ERROR', 'Single error vector evaluation' -NotMatch |  ForEach-Object { $_.Line }
            
                if ($content1.length -ne 0) {
                    Add-Content -path "$ErrorPath\LatestErrors.txt" -Value $file
                }
                Add-Content -path "$ErrorPath\LatestErrors.txt" -Value $content1, "`n"
                Remove-Item "$ErrorPath\temp.txt"
            }
        }
        [System.Windows.MessageBox]::Show('Done')
    }

    Function EALogs {
        if ($x) {
            $LogPath = $x
            $ErrorPath = "$LogPath\ErrorLogs"
        }
        else {
            $LogPath = "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist\Logs"
            $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
            $ErrorPath = "$fpath\ErrorLogs"
        }
 
        $EALogs = Get-ChildItem -Path $LogPath -Recurse | Where-Object {
                                                    ($_.Name -match "EyeAssistEngine.*.log") -or
                                                    ($_.Name -match "EyeTrackingSettings.*.log") -or
                                                    ($_.Name -match "RegionInteraction.*.log")
        } | Select-Object -expand Fullname
        #Creating folder
        if (!(Test-Path "$ErrorPath")) {
            Write-Host "Creating ErrorLogs folder in $ErrorPath .."
            New-Item -Path "$ErrorPath" -ItemType Directory   
        }
        #Creating files
        if (!(Test-Path "$ErrorPath\EALogs.txt")) {
            New-Item -Path $ErrorPath -Name "EALogs.txt" -ItemType "file"
        } 
        else {
            Clear-Content -Path "$ErrorPath\EALogs.txt"
        }

        $EAcontent = Get-ChildItem -Path $EALogs -Recurse | Sort-Object name -desc | Select-Object -expand Fullname

        foreach ($NewEAContent in $EAcontent) {
            $test = New-Item -Path $ErrorPath -Name "temp.txt" -ItemType "file"
            Get-Content -Path "$NewEAContent" -Raw | ForEach-Object -Process { $_ -replace "- `r`n", '- ' } | Add-Content -Path "$ErrorPath\temp.txt"
            #$content2 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "error" -AllMatches | ForEach-Object { $_.Line }
            $content2 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "error", "WixRemoveFoldersEx" -AllMatches | select-string -pattern 'Single error vector evaluation' -NotMatch | ForEach-Object { $_.Line }
            if ($content2.length -ne 0) {
                Add-Content -path "$ErrorPath\EALogs.txt" -Value $NewEAContent
            } 
            Add-Content -path "$ErrorPath\EALogs.txt" -Value $content2, "`n"
            Remove-Item "$ErrorPath\temp.txt"
        }
        [System.Windows.MessageBox]::Show('Done')
    }

    Function TTechLogs {
        if ($x) {
            $LogPath = $x
            $ErrorPath = "$LogPath\ErrorLogs"
        }
        else {
            $LogPath = "$ENV:ProgramData\Tobii", "$ENV:USERPROFILE\AppData\Local\Tobii"
            $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
            $ErrorPath = "$fpath\ErrorLogs"
        }

        #Creating folder
        if (!(Test-Path "$ErrorPath")) {
            Write-Host "Creating ErrorLogs folder in $ErrorPath.."
            New-Item -Path "$ErrorPath" -ItemType Directory   
        }
     
        $files = @("pr_log", "ServerLog" , "InteractionLog", "ServiceLog", "ConfigurationLog", "TrayLog")

        foreach ($path in $files) {
            $newfile = "$ErrorPath\$path.txt"
            if (!(Test-path $newfile)) {
                New-Item -ItemType File -Path $newfile
            }
            else {
                Clear-Content -Path "$newfile"
            }
            $TTcontents = Get-ChildItem -Include "$path*.*" -Path $LogPath -Recurse  | Sort-Object name -desc | Where-Object fullname -NotLike "$ErrorPath\$path.txt"   | Select-Object -expand Fullname
            foreach ($TTcontent in $TTcontents) {
                $test = New-Item -Path $ErrorPath -Name "temp.txt" -ItemType "file"
                Get-Content -LiteralPath "$TTcontent" -Raw | ForEach-Object -Process { $_ -replace "- `r`n", '- ' } | Add-Content -Path "$ErrorPath\temp.txt"
                $content3 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "error", "WixRemoveFoldersEx" -AllMatches | select-string -pattern 'NO_ERROR', 'NoError' -NotMatch | ForEach-Object { $_.Line }
                if ($content3.length -ne 0) { 
                    Add-Content -path $newfile -Value $TTcontent
                }	
                Add-Content -path $newfile -value $content3, "`n"
                Remove-Item "$ErrorPath\temp.txt"
            }
        }
        [System.Windows.MessageBox]::Show('Done')
    }

    Function InstallerLogs {
        if ($x) {
            $InstallerLogs = "$x\TOBII_INSTALLER_LOGS\TEMP"
            $ErrorPath = "$x\ErrorLogs"
        }
        else {
            $InstallerLogs = "$ENV:USERPROFILE\AppData\Local\Temp"
            $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
            $ErrorPath = "$fpath\ErrorLogs"
        }
   
        if (!(Test-Path "$ErrorPath")) {
            Write-Host "Creating ErrorLogs folder.."
            New-Item -Path "$ErrorPath" -ItemType Directory   
        }
        if (!(Test-Path "$ErrorPath\InstallerError.txt")) {
            $ErrorFile = New-Item -Path $ErrorPath -Name "InstallerError.txt" -ItemType "file"
        }
        else {
            Clear-Content -Path "$ErrorPath\InstallerError.txt"
        }

        if (Test-path $InstallerLogs) { 
            $Installercontent = Get-ChildItem -Include "Tobii*.*" -Path $InstallerLogs -Recurse -File |  Sort-Object name -desc | Select-Object -expand Fullname
            foreach ($NewInstallercontent in $Installercontent) {
                $test = New-Item -Path $ErrorPath -Name "temp.txt" -ItemType "file"
                Get-Content -Path "$NewInstallercontent" -Raw | ForEach-Object -Process { $_ -replace "- `r`n", '- ' } | Add-Content -Path "$ErrorPath\temp.txt"
                #$content9 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "error" -AllMatches | ForEach-Object { $_.Line }
                $content9 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "error", "WixRemoveFoldersEx:  Error" -AllMatches | select-string -pattern "3: Error", "error status: 0", 'ErrorDialog' -NotMatch | ForEach-Object { $_.Line }
                if ($content9.length -ne 0) { 
                    Add-Content -path "$ErrorPath\InstallerError.txt" -Value $NewInstallercontent
                }	
                add-Content "$ErrorPath\InstallerError.txt" -value $content9, "`n"
                Remove-Item "$ErrorPath\temp.txt"	
            }
        }
        else { Write-Host "Files are not existed" }
        [System.Windows.MessageBox]::Show('Done')
    }

    Function OtherLogs {
        $LogPath = $x
        #if given path is a folder:
        if ((Get-Item $LogPath) -is [System.IO.DirectoryInfo]) {
            #Creating folder
            $ErrorPath = "$LogPath\ErrorLogs"
            if (!(Test-Path "$ErrorPath")) {
                Write-Host "Creating ErrorLogs folder.."
                New-Item -Path "$ErrorPath" -ItemType Directory   
            }
            #Creating files
            if (!(Test-Path "$ErrorPath\errorlogs.txt")) {
                New-Item -Path $ErrorPath -Name "errorlogs.txt" -ItemType "file"
                Write-Host "creating file"
            }
            else {
                Clear-Content -Path "$ErrorPath\errorlogs.txt"
                Write-Host "cleaing"
            }
            $Othercontent = Get-ChildItem -Path $LogPath -file | Sort-Object name -desc | Select-Object -expand Fullname
            foreach ($NewOthercontent in $Othercontent) {
                New-Item -Path $ErrorPath -Name "temp.txt" -ItemType "file"
                Get-Content -Path "$NewOthercontent" -Raw | ForEach-Object -Process { $_ -replace "- `r`n", '- ' } | Add-Content -Path "$ErrorPath\temp.txt"
                #$content10 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "error" -AllMatches | ForEach-Object { $_.Line }
                $content10 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "error", "WixRemoveFoldersEx", "WixQuietExec" -AllMatches | ForEach-Object { $_.Line }
                if ($content10.length -eq 0) {
                    Write-Host "empty"
                } 
                else {
                    Add-Content -path "$ErrorPath\errorlogs.txt" -Value $NewOthercontent
                }	
                Add-Content -path "$ErrorPath\errorlogs.txt" -Value $content10, "`n"
                Remove-Item "$ErrorPath\temp.txt"
            }

        }
        else {
            #or if given path is a file 
            $LogPath2 = Split-Path -Path $LogPath
            $ErrorPath = "$LogPath2\ErrorLogs"
            if (!(Test-Path "$ErrorPath")) {
                Write-Host "Creating ErrorLogs folder.."
                New-Item -Path "$ErrorPath" -ItemType Directory   
            }
            #Creating files
            if (!(Test-Path "$ErrorPath\errorlogs.txt")) {
                New-Item -Path $ErrorPath -Name "errorlogs.txt" -ItemType "file"
                Write-Host "creating file"
            }
            else {
                Clear-Content -Path "$ErrorPath\errorlogs.txt"
                Write-Host "cleaing"
            }
            New-Item -Path $ErrorPath -Name "temp.txt" -ItemType "file"
            Get-Content -Path "$LogPath" -Raw | ForEach-Object -Process { $_ -replace "- `r`n", '- ' } | Add-Content -Path "$ErrorPath\temp.txt"
            #$content11 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "error" -AllMatches | ForEach-Object { $_.Line }
            $content11 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse | Select-String -Pattern "error", "WixRemoveFoldersEx", "WixQuietExec" -AllMatches | ForEach-Object { $_.Line }
            if ($content11.length -eq 0) {
                Write-Host "empty"
            } 
            else {
                Add-Content -path "$ErrorPath\errorlogs.txt" -Value $LogPath
            }	
            Add-Content -path "$ErrorPath\errorlogs.txt" -Value $content11, "`n"
            Remove-Item "$ErrorPath\temp.txt"
        }
        [System.Windows.MessageBox]::Show('Done')
    }

    #C:\Users\aes\Desktop\SupportTools\accessory.tdl
    Function BWLogConvertor {
        Write-Host "Running BW log convertor `r`n"
        $LogPath = $x
        if ($LogPath -match "accessory.tdl") { 
            $newLogPath = $LogPath -replace "accessory.tdl", ""
        }
        elseif ($LogPath -match "accessory") {
            $newLogPath = $LogPath -replace "accessory", ""
        }
        else {
            $newLogPath = $LogPath
        }

        #Creating files
        if (!(Test-Path "$newLogPath\accessory.txt")) {
            New-Item -Path $newLogPath -Name "accessory.txt" -ItemType "file"
            Write-Host "creating file"
        }
        else {
            Clear-Content -Path "$newLogPath\accessory.txt"
            Write-Host "cleaing"
        }
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "extract_logs.py" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        Set-Location $fpath
        #$logfile = "C:\Users\aes\Desktop\accessory.tdl"
        #$textfile = "$fpath"\test2.txt"
        $newnewLogPath = "$newLogPath\accessory.tdl"
        $Test = Start-Process cmd "/c `"extract_logs.py  $newnewLogPath > $newLogPath\accessory.txt`""
    }

    Function TimingIssueEventLogFinder {

        Write-Host "Running Timing Issue EventLog Finder `r`n"
        $LogPath = $x

        $Path = Get-ChildItem -Path $LogPath -Filter "*.evtx" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
   
        #Creating files
        if (!(Test-Path "$Path\TimingIssue.txt")) {
            New-Item -Path $Path -Name "TimingIssue.txt" -ItemType "file"
            Write-Host "creating file"
        }
        else {
            Clear-Content -Path "$Path\TimingIssue.txt"
            Write-Host "cleaing"
        }
        $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "TimeSyncIssueFinder.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
        Set-Location $fpath
    
        $Test2 = Start-Process cmd "/c `"TimeSyncIssueFinder.exe -f  $LogPath > $Path\TimingIssue.txt`""

    }

    Function RISample {
        Write-Host "Running RI Sample results `r`n"
        $LogPath = $x
        #$path = "C:\Users\aes\Desktop\tobii"
        $content = Get-ChildItem -Path $LogPath -Recurse | Where-Object { $_.Name -match 'Tdx.EyeTracking.RegionInteraction.EyeAssist.Sample' } | Select-Object -expand Fullname

        foreach ($newcontent in $content) {

            $lines = Get-Content -path $newcontent -raw
            $lines | Select-String '\d{4}\-(0?[1-9]|1[012])\-(0?[1-9]|[12][0-9]|3[01])*\s(\d+:\d+:\d+)' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value } | Set-Content "$LogPath\text.txt"
            #$timestamps = @([datetime]"03:37:51", [datetime]"03:37:53", [datetime]"03:37:54")

            [datetime[]] $timestamps = @(Get-Content -Path "$LogPath\text.txt")

            if ($timestamps.Count -lt 2) {
                Write-Host "Only one result: " $timestamps[0]
                return
            }

            for ($i = 0; $i -lt $timestamps.Count; $i++) {
                $previous = $timestamps[$i]
                $current = $timestamps[$i + 1]
                $difference = ($current - $previous)
                #($current - $previous) | Out-File "$path\text2.txt" -Append
                #Add-content $Logfile -value $logstring
                #Add-Content "$path\text2.txt" ($current - $previous)

                if (($difference) -gt ("00:00:05")) {
                    Add-Content -path "$LogPath\Results.txt" -Value $newcontent
                    Add-Content "$LogPath\Results.txt" "Gap between $current and $previous with ($difference)`n"
                } 
            }
        }

        Remove-Item "$LogPath\text.txt"
        #Remove-Variable * -ErrorAction SilentlyContinue

    }

    Function TimeStampBetween {
        if ($x) {
            $LogPath = $x
            $ErrorPath = "$LogPath\ErrorLogs"
        }
        else {
            $LogPath = "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox", "$ENV:ProgramData\Tobii Dynavox", "$ENV:USERPROFILE\AppData\Local\Tobii", "$ENV:ProgramData\Tobii"
            $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
            $ErrorPath = "$fpath\ErrorLogs"

        }

        #$date = "2020-11-22"
        $start = Get-Date -format "yyyy-MM-dd hh:mm:ss" "$x3"
        $end = Get-Date -format "yyyy-MM-dd hh:mm:ss" "$x4"
  
        write-host "start: $start"
        write-host "end: $end"
        # Pattern explaination # ^ matches the begginning of each line # \d matches a decimal character
        # {4},{2},{3} repeats the previous character # so \d{4} matches any four numerals # / and : are literally / and :
        # a period is a special regex character so it needs escaped \.
        $pattern = "^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\."

        #Creating folder
        if (!(Test-Path "$ErrorPath")) {
            Write-Host "Creating ErrorLogs folder in $ErrorPath.." 
            $NewFolder = New-Item -Path "$ErrorPath" -ItemType Directory  
        }

        if ($start -match ":") {
            $newStart = $start -replace ":" , "."
        }
        else { $newstart = $start }

        if ($end -match ":") {
            $newEnd = $end -replace ":" , "."
        }
        else { $newEnd = $end }

        $textfile = "$newStart - $newEnd"
        if (!(Test-Path "$ErrorPath\$textfile.txt")) {
            $NewItem = New-Item -Path $ErrorPath -Name "$textfile.txt" -ItemType "file"
            Write-Host "Creating file in $NewItem"
        }
        elseif (Test-Path "$ErrorPath\$textfile.txt") {
            Clear-Content -Path "$ErrorPath\$newStart - $newEnd.txt"
            Write-Host "$NewItem is already existing, cleaning the file"
        }

        $files = @("ServerLog", "ServiceLog", "pr_log", "InteractionLog", "ConfigurationLog", "Tray", "EyeAssistEngine", "EyeTrackingSettings", "RegionInteraction", "ComputerControl" , "Switcher", "Phone", "Talk", "Browse", "Updater", "Tdx.Overlays")

        foreach ($path in $files) {
            Write-Host "Analysing $path and collecting logs"
            $Servicecontent2 = Get-ChildItem -Include "*$path*.log", "$path*.log.*" , "$path*.txt", "$path*.txt.*" -Path $LogPath -Recurse  | Sort-Object name -desc | Where-Object fullname -NotLike "$ErrorPath\$path.txt" | Select-Object -expand Fullname
            foreach ($NewServicecontent2 in $Servicecontent2) {
                $newtemp = New-Item -Path $ErrorPath -Name "temp.txt" -ItemType "file"
                Get-Content -Path "$NewServicecontent2" -Raw | ForEach-Object -Process { $_ -replace "- `r`n", '- ' -replace ",", "." } | Add-Content -Path "$ErrorPath\temp.txt"
                $content21 = Get-ChildItem -path "$ErrorPath\temp.txt" -Recurse #| Select-String -Pattern "$date" -AllMatches | ForEach-Object { $_.Line }
                $entries = $content21 | Select-String -Pattern $pattern | ForEach-Object {
                    [pscustomobject]@{ 
                        'Date' = [datetime]::Parse($_.Matches[0].Value) 
                        'Line' = $_.LineNumber 
                        'Text' = $_.Line
                    }										 
                }
                $filtered = $entries | Where-Object { $_.Date -ge $start -and $_.Date -le $end } | Sort-Object Date 
                if ($filtered) {
                    $first = $filtered[0].Line - 1 
                    $last = $filtered[-1].Line - 1 
                    $content21[$first..$last] 
                }
                if ($filtered.length -ne 0) { 
                    Add-Content "$ErrorPath\$textfile.txt" -Value $NewServicecontent2
                }
                Add-Content "$ErrorPath\$textfile.txt" -value $filtered.text, "`n"
                Remove-Item "$ErrorPath\temp.txt"
            }
        }
        [System.Windows.MessageBox]::Show('Done')
    }
 
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $x = $textBox.Text
        $x2 = $textBox2.Text
        $x3 = $textBox3.Text
        $x4 = $textBox4.Text
        $x
        $x2
        $x3
        $x4

        if ($x2 -match "1") { 
            LatestErrorLogs
        }
        elseif ($x2 -match "2") { 
            EALogs
        }
        elseif ($x2 -match "3") { 
            TTechLogs
        }
        elseif ($x2 -match "4") { 
            InstallerLogs
        }
        elseif ($x2 -match "5") {
            OtherLogs
        }
        elseif ($x2 -match "6") {
            BWLogConvertor
        }
        elseif ($x2 -match "7") {
            TimingIssueEventLogFinder
        }    
        elseif ($x2 -match "8") {
            RISample
        }
        elseif (!($x2)) {
            if ("$x3" -and "$x4") {
                TimeStampBetween
            }
        }

    }
}

#B17
Function InstalledAppsTool {
    $outputBox.clear()
    $outputBox.appendtext( "Tdx.InstalledAppsTool.exe...`r`n" )

    $fpath = (Get-ChildItem -Path "$PSScriptRoot" -Filter "Tdx.InstalledAppsTool.exe" -Recurse).FullName | Split-Path 
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        start-process .\Tdx.InstalledAppsTool.exe
    }
    else { 
        $outputbox.appendtext("File Tdx.InstalledAppsTool.exe is missing!`r`n" )
    }	
    $outputbox.appendtext("Done! `r`n")
}

Function Write-LogT {
    Param ($Message)
    $fpath1 = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path  
    Set-Location $fpath1
    "$(get-date -format "yyyy-MM-dd HH:mm:ss"): $($Message)" | out-file "$fpath1\TroubleshootingsLog.txt" -Append
    $OutputBox.AppendText("$Message" + "`r`n" )
}

#B18 "Troubleshoot"
#Todo Check
#add new section for HW check and other 
Function Troubleshoot {
    $outputBox.clear()
    #File version 
    $VersionLatest = "4.180.0.29190"
    $PDKVersionLatest = "1.42.2.0_46cc4824dc"

    $PCEyeDisplayName = "Tobii Experience Software For Windows (PCEye5)"
    $ISeriesDisplayName = "Tobii Experience Software For Windows (I-Series)"
    $ServicePCEye5 = "TobiiIS5LARGEPCEYE5"
    $ServiceGibbon = "TobiiIS5GIBBONGAZE"

    #1 Pinging ET and checking HW & fw
    Write-LogT -Message "******Start analyze of Eye Tracker******`r`n"
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "Tdx.EyeTrackerInfo.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        #Start PDK service
        try {
            $getService = Get-Service -Name '*TobiiIS5*'  | start-Service -PassThru -ErrorAction Ignore
        }
        catch {
            Write-LogT -Message "Error starting PDK. Make sure that ET is connected and (TobiiIS5XXXX) available in Task Manager-Services."
        }

        $global:serialnumber = .\Tdx.EyeTrackerInfo.exe --serialnumber
        if ($serialnumber.count -gt 0) {
            if ($serialnumber -match "IS514") {
                $global:LatestDisplayName = "$PCEyeDisplayName"
                Write-LogT -Message "PASS: Connected Eye Tracker is PCEye5 with S/N $serialnumber"
            }
            elseif ($serialnumber -match "IS502") {
                $global:LatestDisplayName = "$ISeriesDisplayName"
                Write-LogT -Message "PASS: Connected Eye Tracker is I-Series with S/N $serialnumber"
            }
        }
        elseif ($serialnumber.count -eq 0) {
            Write-LogT -Message "FAIL: No Eye Tracker Connected. Make sure that ET is connected and there is PDK on this device."
        }
    }

    #2 Listing Tobii Experience software that installed on this device
    Write-LogT -Message "******Check Eye Tracker software******`r`n"
    $AppLists = Get-ChildItem -Recurse -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Tobii\ | 
    Get-ItemProperty | Where-Object { 
        $_.Displayname -like '*Tobii Experience Software*' -or
        $_.Displayname -like '*Tobii Device Drivers*' -or
        $_.Displayname -like '*Tobii Eye Tracking For Windows*' -or
        $_.Displayname -like '*Tobii Eye Tracking*'
    }

    $DisplayNameApps = $AppLists.DisplayName
    $DisplayVersion = $AppLists.displayversion
	
    if ($DisplayNameApps.count -eq 1) {
        #Write-LogT -Message "PASS: $DisplayNameApps is installed."
        if ($DisplayNameApps -eq $LatestDisplayName) { 
            Write-LogT -Message "PASS: $DisplayNameApps is correct SW for the HW."
        }
        elseif ($DisplayNameApps -ne $LatestDisplayName) {
            Write-LogT -Message "FAIL: No HW attached."
        }	
    } 
    elseif ($DisplayNameApps.count -gt 1) {
        Write-LogT -Message "FAIL: Installed ET software on this device are following:"
        foreach ($L in $DisplayNameApps) {
            Write-LogT -Message "$L`r`n"
        }
        Write-LogT -Message "Uninstall all sw named above and install only $LatestDisplayName $VersionLatest."
    }
	
    if ($DisplayVersion -eq $VersionLatest) {
        Write-LogT -Message "PASS: Latest version of $DisplayNameApps is installed, $DisplayVersion."
    }
    else {
        Write-LogT -Message "FAIL: $DisplayNameApps is not the latest. Upgrade the software through Update Notifier."
    }
	
    # Check for Experience app
    $AppPackage = Get-AppxPackage -Name *TobiiAB.TobiiEyeTrackingPortal*
    if ($AppPackage) {
        Write-LogT -Message "FAIL: $AppPackage Shall be removed."
        $regpaths = "HKLM:\SYSTEM\CurrentControlSet\Services\Tobii Interaction Engine",
        "HKLM:\SYSTEM\CurrentControlSet\Services\Tobii Service",
        "HKLM:\SYSTEM\CurrentControlSet\Services\TobiiGeneric",
        "HKLM:\SYSTEM\CurrentControlSet\Services\TobiiIS5LARGEPCEYE5",
        "HKLM:\SYSTEM\CurrentControlSet\Services\TobiiIS5EYETRACKER5"
        if (test-path $regpaths) {
            Write-LogT -Message "FAIL: Delete $regpaths"
        }
    }

    $AllTDListApps = (Get-ChildItem -Recurse -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Tobii\ | 
        Get-ItemProperty | Where-Object { 
        ($_.Displayname -Match "Windows Control") -or
        ($_.Displayname -Match "Tobii Dynavox Gaze Point") -or
        ($_.Displayname -Match "GazeSelection") 
        } | Select-Object Displayname, UninstallString).DisplayName
    if ($AllTDListApps -gt 0) {
        Write-LogT -Message "FAIL: $AllTDListApps shall be removed. Uninstall also Tobii Dynavox Eye Tracking and re-install it again."
    }

    #3 Getting installed Services that installed on this device
    Write-LogT -Message "******Check services******`r`n"
    $GetService = Get-Service -Name '*Tobii*'
    #Listing all installed Services
    if ($GetService.count -ne 0) {
        $EyeXPath = "C:\Program Files\Tobii\Tobii EyeX"
        if (Test-Path $EyeXPath) {
            if ($global:serialnumber -match "IS502") {
                $global:ReqService = $ServiceGibbon
                $PDKversions = Get-ChildItem -Path $EyeXPath -Recurse -file -include "platform_runtime_IS5GIBBONGAZE_service.exe" | foreach-object { "{0}`t{1}" -f $_.Name, [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion }
            }
            elseif ($global:serialnumber -match "IS514") {
                $global:ReqService = $ServicePCEye5
                $PDKversions = Get-ChildItem -Path $EyeXPath -Recurse -file -include "platform_runtime_IS5LARGEPCEYE5_service.exe" | foreach-object { "{0}`t{1}" -f $_.Name, [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion }
            } 
        }
        $TobiiService = $GetService | Where-Object { $_.Name -eq "Tobii Service" }
        if ( $TobiiService) {
            Write-LogT -Message "PASS: Tobii Service is installed."
        }
        elseif (!($TobiiService)) {
            Write-LogT -Message "FAIL: Tobii Service is not intsalled. Uninstall Tobii Experince Software and re-install it again."
        }
        $ServiceStatus = ($GetService | Where-Object { $_.Status -ne "Running" }).name
        If ($ServiceStatus) {
            foreach ($ServiceStatuss in $ServiceStatus) {
                Write-LogT -Message "FAIL: $ServiceStatuss is not running. Open Task Manager and run the service."
            }
        }

        if ($PDKversions) {
            $Compares = (Compare-Object -DifferenceObject $GetService -ReferenceObject $ReqService -CaseSensitive -ExcludeDifferent -IncludeEqual | Select-Object InputObject).InputObject
            if ($Compares -eq $ReqService ) {
                if ($PDKversions -match $PDKVersionLatest) {
                    Write-LogT -Message "PASS: Latest PDK ($ReqService) $PDKversions is installed."
                }
                else {
                    Write-LogT -Message "FAIL: PDK ($ReqService $PDKversions) is not the latest, make sure that $LatestDisplayName $VersionLatest is installed."
                }
                
                $AvaliablePDK = (Get-Service -DisplayName "*Tobii Runtime Service*").name
                $ComparesPDK = (Compare-Object -DifferenceObject $AvaliablePDK -ReferenceObject $Compares -CaseSensitive  | Select-Object InputObject).InputObject
                if (($ComparesPDK)) {
                    Write-LogT -Message "FAIL: $ComparesPDK should not be installed. Please remove it!"
                }
            } 
            elseif (!($Compares) ) { 
                Write-LogT -Message "FAIL: NO PDK INSTALLED. Make sure $LatestDisplayName is installed."
            }
        } 
        else {
            Write-LogT -Message "FAIL: No ET HW found. Make sure ET is connected."
        }
    }
    else {
        Write-LogT -Message "FAIL: No ET Service found. Make sure $LatestDisplayName is installed."
    }
    
    #4 Getting Processes that running on this device
    Write-LogT -Message "******Check processes******`r`n"
    $TobiiProcesses = "Tobii.EyeX.Engine", "Tobii.EyeX.Interaction", "Tobii.Service", "TobiiDynavox.EyeAssist.Engine", "TobiiDynavox.EyeAssist.RegionInteraction.Startup", "TobiiDynavox.EyeAssist.Smorgasbord", "TobiiDynavox.EyeAssist.TrayIcon", "TobiiDynavox.EyeTrackingSettings"
    foreach ($TobiiProcess in $TobiiProcesses) {
        Try {
            $erroractionpreference = "Stop"
            $GetTobiiProcess = Get-Process $TobiiProcess | Select-Object ProcessName
        }
        catch {
            Write-LogT -Message "FAIL: $TobiiProcess is not running. Open Task Manager and run the process."
        }
    }
    Write-LogT -Message "Completed analyze processes"

    #5 Getting drivers that installed on this device
    Write-LogT -Message "******Check drivers******`r`n"
    $TobiiWindowsDrivers = Get-WindowsDriver -Online | Where-Object { $_.OriginalFileName -match "Tobii" } | Sort-object OriginalFileName -desc | Select-Object OriginalFileName, Driver
    if ($TobiiWindowsDrivers.count -ne 0) {

        $NewTobiiDrivers = $TobiiWindowsDrivers.originalfilename -replace "C:", "" -replace "(?<=\\).+?(?=\\)", "" -replace "\\\\\\", "" 

        $b = $NewTobiiDrivers | Select-Object -Unique
        $CompareDrivers = (Compare-Object -ReferenceObject $b -DifferenceObject $NewTobiiDrivers | Select-Object InputObject).InputObject
        if ($CompareDrivers.count -gt 0) {
            Write-LogT -Message "FAIL: There are two drivers of $CompareDrivers. Uninstall Tobii Experience Software and re-install it again."
        } 
         
        foreach ($NewTobiiDriver in $NewTobiiDrivers) {
            if (($NewTobiiDriver -match "is") -or ($NewTobiiDriver -match "dmft")) {
                Write-LogT -Message "FAIL: $NewTobiiDriver is not belong to this HW! Remove all sw in second step and install only Tobii Experience Software and Tobii Dynavox Eye Tracking."
            }
            
            if (($NewTobiiDriver -match "318") -and ($global:serialnumber -match "IS502")) {
                Write-LogT -Message "FAIL: $NewTobiiDriver belong to PCEye5."
            }
            elseif (($NewTobiiDriver -match "304") -and ($global:serialnumber -match "IS514")) {
                Write-LogT -Message "FAIL: $NewTobiiDriver belong to I-Series."
            }
            else {
                Write-LogT -Message "PASS: Installed Drivers are $NewTobiiDrivers"
            }
        }
    }

    $SignedDrivers = (Get-WmiObject Win32_PnPSignedDriver | Where-Object { $_.Manufacturer -match "Tobii" } | Select-Object DeviceName).DeviceName
    $d = "Tobii Hello Sensor", "Tobii Eye Tracker HID", "Tobii Device"
    if ($SignedDrivers.Count -eq 0) {
        Write-LogT -Message "FAIL: No signed drivers found. Uninstall Tobii Experience Software and re-install it again." 
        
    }
    elseif ($SignedDrivers.Count -gt 0) {
        $CompareSignedDrivers = (Compare-Object -ReferenceObject $d -DifferenceObject $SignedDrivers | Where-Object { $_.SideIndicator -eq "<=" }).InputObject
        if ($CompareSignedDrivers.count -gt 0) {
            Write-LogT -Message "FAIL: $CompareSignedDrivers is missing. Uninstall Tobii Experience Software and re-install it again." 
        }
        else {
            Write-LogT -Message "PASS: Installed drivers are $SignedDrivers"
        }
    }
    #List from Device Manager
    #$GetDriverStatus = (Get-PnpDevice -FriendlyName '*Tobii*' | Where-Object { $_.Status -ne "OK" } | Select-Object FriendlyName, InstanceId).FriendlyName
    $GetPnpDrivers = Get-PnpDevice -FriendlyName '*Tobii*' | Select-Object Status, Class, FriendlyName, InstanceId
    $GetPnpDriversName = $GetPnpDrivers.FriendlyName
    $ReferencePnpDrivers = "Tobii Device", "Tobii Hello Sensor", "Tobii Eye Tracker HID"

    #foreach ($GetPnpDriver in $GetPnpDrivers ) {
    #    if ($GetPnpDriver.Status -ne "OK") {
    #        $getPnpDriverName = $GetPnpDriver.FriendlyName
    #        $getPnpDriverStatus = $GetPnpDriver.status
    #        Write-LogT -Message "FAIL: $getPnpDriverName Status is $getPnpDriverStatus..re-install Tobii Experience."
    #        #write-host $GetPnpDriver.FriendlyName "Status is" $GetPnpDriver.status
    #    }
    #}
    $ComparePnpDrivers = (Compare-Object -ReferenceObject $ReferencePnpDrivers -DifferenceObject $GetPnpDrivers.FriendlyName | Where-Object { $_.SideIndicator -eq "<=" }).InputObject
    if ($ComparePnpDrivers -gt 0) {
        foreach ($ComparePnpDriver in  $ComparePnpDrivers) {
            Write-LogT -Message "FAIL: $ComparePnpDriver is missing, Uninstall Tobii Experience Software and re-install it again."
        }
    }
    else {
        Write-LogT -Message "PASS: Devices are: $GetPnpDriversName"
    }
    
    #6 Check if there are valid calibration profiles
    Write-LogT -Message "******Check calibration profiles/display setup******`r`n"
    $EyeXConfig = "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig"
    if (Test-path $EyeXConfig) {
        $CurrentProfile = (Get-itemproperty -Path $EyeXConfig).currentuserprofile
        if ($CurrentProfile.count -gt 0) {
            Write-LogT -Message "PASS: Active profile is $currentprofile."
        }
        else {
            Write-LogT -Message "FAIL: No active profile. Open Tobii Dynavox Eye Tracking and create a new calibration profile."
        }
    }

    $UserProfile = "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\UserProfiles"
    if (Test-Path $UserProfile) {
        $getCalbfolders = (Get-ChildItem -Path $UserProfile).name | Split-Path -Leaf
        if ($getCalbfolders.count -gt 0) {
            Write-LogT -Message "Following Calibration profiles are created in this device:"
            foreach ($getCalbfolder in $getCalbfolders) {
                $f = (Get-ChildItem -Path "$UserProfile\$getCalbfolder" -Recurse).property
                if ($f.contains('Data')) {
                    Write-LogT -Message "PASS: $getCalbfolder is created, Data file is exist."
                }
                else { 
                    Write-LogT -Message "FAIL: $getCalbfolder is created, but Data file is not exist. Open Tobii Dynavox Eye Tracking and re-calibrate $getCalbfolder."
                }
            }
        } 
        elseif ($getCalbfolders.count -eq 0) {
            Write-LogT -Message "FAIL: No Calibration Profile stored in this device. Open Tobii Dynavox Eye Tracking and create a new calibration profile."
        }   
    }    

    #display-setup
    $regEntryPath = 'HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig\MonitorMappings'
    $referenceValues = "ActiveDisplayArea", "AspectRatioHeight", "AspectRatioWidth"
    try {
        $erroractionpreference = "Stop"
        $keyValue = (Get-ChildItem -Path $regEntryPath).property 
        $keyValue = $keyValue | Select-Object -Unique
        $comparekeys = (Compare-Object -DifferenceObject $keyValue -ReferenceObject $referenceValues -CaseSensitive).inputobject
        if ($comparekeys.count -gt 0) {
            Write-LogT -Message "FAIL: No display values has been found! Open Tobii Dynavox Eye Tracking and perform display setup if possible."
        }
        elseif ($comparekeys.count -eq 0) { 
            Write-LogT -Message "PASS: Display setup has been performed!"
        }
    } 
    catch {
        Write-LogT -Message "FAIL: No display setup has been found! Open Tobii Dynavox Eye Tracking and perform display setup if possible."
    }

    Write-LogT -Message "Done!"
}


#Windows forms
$Optionlist = @("Remove Progressive Suite", "Remove PCEye5 Bundle", "Remove all ET SW", "Remove WC&GP Bundle", "Remove VC++", "Remove PCEye Package", "Remove Communicator", "Remove Compass", "Remove TGIS only", "Remove TGIS profile calibrations", "Remove all users C5", "Remove C5 Emails", "Backup Gaze Interaction", "Copy License")
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(600, 590)
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False

#Informationtext above the dropdown list.
$DropDownLabel = new-object System.Windows.Forms.Label
$DropDownLabel.Location = new-object System.Drawing.Size(10, 10)
$DropDownLabel.size = new-object System.Drawing.Size(200, 30)
$DropDownLabel.Text = "Select an option"
$Form.Controls.Add($DropDownLabel)

#Dropdown list with options
$DropDownBox = New-Object System.Windows.Forms.ComboBox
$DropDownBox.Location = New-Object System.Drawing.Size(10, 40)
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
$Button.Location = New-Object System.Drawing.Size(10, 80)
$Button.Size = New-Object System.Drawing.Size(180, 50)
$Button.Text = "Start"
$Button.Font = New-Object System.Drawing.Font ("" , 12, [System.Drawing.FontStyle]::Regular)
$Form.Controls.Add($Button)
$Button.Add_Click{ selectedscript }

#B1 Button1 
$Button1 = New-Object System.Windows.Forms.Button
$Button1.Location = New-Object System.Drawing.Size(420, 0)
$Button1.Size = New-Object System.Drawing.Size(150, 30)
$Button1.Text = "All versions"
$Button1.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button1)
$Button1.Add_Click{ Listapps }

#B2 Button2 
$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(420, 30)
$Button2.Size = New-Object System.Drawing.Size(150, 30)
$Button2.Text = "Get Services"
$Button2.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button2)
$Button2.Add_Click{ GetServices }

#B3 Button3 
$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Size(420, 60)
$Button3.Size = New-Object System.Drawing.Size(150, 30)
$Button3.Text = "Restart Services"
$Button3.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button3)
$Button3.Add_Click{ RestartProcesses }

#B4 Button4 
$Button4 = New-Object System.Windows.Forms.Button
$Button4.Location = New-Object System.Drawing.Size(420, 90)
$Button4.Size = New-Object System.Drawing.Size(150, 30)
$Button4.Text = "Firmware v / Upgrade"
$Button4.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button4)
$Button4.Add_Click{ ETfw }

#B5 Button5 
$Button5 = New-Object System.Windows.Forms.Button
$Button5.Location = New-Object System.Drawing.Size(420, 120)
$Button5.Size = New-Object System.Drawing.Size(150, 30)
$Button5.Text = "WCF"
$Button5.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button5)
$Button5.Add_Click{ WCF }

#B6 Button6 
$Button6 = New-Object System.Windows.Forms.Button
$Button6.Location = New-Object System.Drawing.Size(420, 150)
$Button6.Size = New-Object System.Drawing.Size(150, 30)
$Button6.Text = "SMBIOS"
$Button6.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button6)
$Button6.Add_Click{ SMBios }

#B7 Button7 
$Button7 = New-Object System.Windows.Forms.Button
$Button7.Location = New-Object System.Drawing.Size(420, 180)
$Button7.Size = New-Object System.Drawing.Size(150, 30)
$Button7.Text = "IR Utility"
$Button7.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button7)
$Button7.Add_Click{ IRUtility }

#B8 Button8 
$Button8 = New-Object System.Windows.Forms.Button
$Button8.Location = New-Object System.Drawing.Size(420, 210)
$Button8.Size = New-Object System.Drawing.Size(150, 30)
$Button8.Text = "Reset ET"
$Button8.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button8)
$Button8.Add_Click{ resetBOOT }

#B9 Button9 
$Button9 = New-Object System.Windows.Forms.Button
$Button9.Location = New-Object System.Drawing.Size(420, 240)
$Button9.Size = New-Object System.Drawing.Size(150, 30)
$Button9.Text = "Retrieve UN"
$Button9.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button9)
$Button9.Add_Click{ RetrieveUnreleased }

#B10 Button10 
$Button10 = New-Object System.Windows.Forms.Button
$Button10.Location = New-Object System.Drawing.Size(420, 270)
$Button10.Size = New-Object System.Drawing.Size(150, 30)
$Button10.Text = "Remove Drivers"
$Button10.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button10)
$Button10.Add_Click{ RemoveDrivers }

#B11 Button11 
$Button11 = New-Object System.Windows.Forms.Button
$Button11.Location = New-Object System.Drawing.Size(420, 300)
$Button11.Size = New-Object System.Drawing.Size(150, 30)
$Button11.Text = "Long Test"
$Button11.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button11)
$Button11.Add_Click{ LongTest }

#B12 Button12 
$Button12 = New-Object System.Windows.Forms.Button
$Button12.Location = New-Object System.Drawing.Size(420, 330)
$Button12.Size = New-Object System.Drawing.Size(150, 30)
$Button12.Text = "Logging"
$Button12.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button12)
$Button12.Add_Click{ Logging }

#B13 Button13 
$Button13 = New-Object System.Windows.Forms.Button
$Button13.Location = New-Object System.Drawing.Size(420, 360)
$Button13.Size = New-Object System.Drawing.Size(150, 30)
$Button13.Text = "InternSE"
$Button13.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button13)
$Button13.Add_Click{ InternalSE }

#B14 Button14 
$Button14 = New-Object System.Windows.Forms.Button
$Button14.Location = New-Object System.Drawing.Size(420, 390)
$Button14.Size = New-Object System.Drawing.Size(150, 30)
$Button14.Text = "Diagnostic"
$Button14.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button14)
$Button14.Add_Click{ Diagnostic }

#B15 Button15 
$Button15 = New-Object System.Windows.Forms.Button
$Button15.Location = New-Object System.Drawing.Size(420, 420)
$Button15.Size = New-Object System.Drawing.Size(150, 30)
$Button15.Text = "Deployment"
$Button15.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button15)
$Button15.Add_Click{ Deployment }

#B16 Button16 
$Button16 = New-Object System.Windows.Forms.Button
$Button16.Location = New-Object System.Drawing.Size(420, 450)
$Button16.Size = New-Object System.Drawing.Size(150, 30)
$Button16.Text = "Log Collector"
$Button16.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button16)
$Button16.Add_Click{ LogCollector }

#B17 Button17 
$Button17 = New-Object System.Windows.Forms.Button
$Button17.Location = New-Object System.Drawing.Size(420, 480)
$Button17.Size = New-Object System.Drawing.Size(150, 30)
$Button17.Text = "InstalledAppsTool"
$Button17.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button17)
$Button17.Add_Click{ InstalledAppsTool }

#B18 Button18 <#
<#
$Button18 = New-Object System.Windows.Forms.Button
$Button18.Location = New-Object System.Drawing.Size(420, 510)
$Button18.Size = New-Object System.Drawing.Size(150, 30)
$Button18.Text = ""
$Button18.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button18)
$Button18.Add_Click{  }
#>

#B19 Button19 "Troubleshoot"
$Button19 = New-Object System.Windows.Forms.Button
$Button19.Location = New-Object System.Drawing.Size(190, 80)
$Button19.Size = New-Object System.Drawing.Size(180, 50)
$Button19.Text = "Troubleshoot"
$Button19.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button19)
$Button19.Add_Click{ Troubleshoot }

#Form name + activate form.
$Form.Text = $fileversion
$Form.Add_Shown( { $Form.Activate() })
$Form.ShowDialog()