#File version 
$fileversion = "SupportTool v1.6.7"
$LogCollectorTool = "LogCollectorTool_V0.9.ps1"

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
    $buttoncancle = New-Object System.Windows.Forms.Button

    $TobiiVer = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ |
    Get-ItemProperty  | Where-Object { 
    ($_.Displayname -eq "Tobii Dynavox Switcher") -or
    ($_.Displayname -eq "Tobii Dynavox Switcher (Beta)") -or
    ($_.Displayname -eq "Tobii Dynavox Browse") -or
    ($_.Displayname -eq "Tobii Dynavox Browse (Beta)") -or
    ($_.Displayname -eq "Tobii Dynavox Phone") -or
    ($_.Displayname -eq "Tobii Dynavox Phone (Beta)") -or
    ($_.Displayname -eq "Tobii Dynavox Talk") -or
    ($_.Displayname -eq "Tobii Dynavox Talk (Beta)") -or
    ($_.Displayname -eq "Tobii Dynavox Control") -or
    ($_.Displayname -eq "Tobii Dynavox Control (Beta)") -or
    ($_.Displayname -eq "Tobii Dynavox Control (Development)") -or
    ($_.Displayname -eq "Tobii Dynavox Telephony Bluetooth Driver")

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
    }
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(283, $buttonplacement)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    $form.ShowDialog()

    foreach ($cbox in $CheckBoxArray) {
        if ($cbox.CheckState -eq "Checked") {
            #If first answer equals yes or no
            $Uninstname = (Compare-Object -DifferenceObject $TobiiVer.displayname -ReferenceObject $cbox.Name -CaseSensitive -ExcludeDifferent -IncludeEqual | Select-Object InputObject).InputObject
            #$Outputbox.Appendtext( "Following apps will be removed $Uninstname`r`n" ) 
            $Outputbox.Appendtext( "Uninstname =$Uninstname")

            if ($Uninstname -eq "Tobii Dynavox Browse") {
                $Browses = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
            ($_.Displayname -eq "Tobii Dynavox Browse Launcher") -or 
            ($_.Displayname -eq "Tobii Dynavox Browse Updater Service") -or
            ($_.Displayname -eq "Tobii Dynavox Browse") -or
            ($_.Displayname -eq "Tobii Dynavox Prediction Service") 
                } | Select-Object Displayname, UninstallString 
        
                foreach ( $Browse in $Browses) {
                    $Displayname = $Browse.Displayname
                    $Outputbox.Appendtext(  "Removing - " + "$Displayname`r`n" )
                    $uninst = $Browse.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
                }
                $pathbrowses = (
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Browse",
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Prediction Service",
                    "$ENV:ProgramData\Tobii Dynavox\Browse",
                    "C:\Program Files\Tobii Dynavox\Browse",
                    "C:\Program Files\Tobii Dynavox\Prediction Service",
                    "HKCU:\Software\Tobii Dynavox\Browse",
                    "HKCU:\Software\Tobii Dynavox\Browse Launcher",
                    "HKCU:\Software\Tobii Dynavox\Browse Updater Service",
                    "HKCU:\Software\Tobii Dynavox\Prediction Service",
                    "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Browse",
                    "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Browse Launcher",
                    "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Browse Updater Service")
                foreach ($pathbrowse in $pathbrowses) {
                    if ($pathbrowse) {
                        $Outputbox.Appendtext(  "Removing - " + "$pathbrowse`r`n" )
                        Remove-item $pathbrowse -Recurse -ErrorAction Ignore
                    }
                } 
            }
            elseif ($Uninstname -eq "Tobii Dynavox Control") {
                $Controls = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
            ($_.Displayname -eq "Tobii Dynavox Control Updater Service") -or
            ($_.Displayname -eq "Tobii Dynavox Control") 
                } | Select-Object Displayname, UninstallString 
        
                foreach ( $Control in $Controls) {
                    $Displayname = $Control.Displayname
                    $Outputbox.Appendtext(  "Removing - " + "$Displayname`r`n" )
                    $uninst = $Control.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
                }
                $pathControls = (
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control",
                    "$ENV:ProgramData\Tobii Dynavox\Computer Control",
                    "C:\Program Files\Tobii Dynavox\Computer Control",
                    "HKCU:\Software\Tobii Dynavox\Computer Control",
                    "HKCU:\Software\Tobii Dynavox\Computer Control Launcher",
                    "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Computer Control",
                    "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Computer Control Updater Service")
                foreach ($pathControl in $pathControls) {
                    if ($pathControl) {
                        $Outputbox.Appendtext(  "Removing - " + "$pathControl`r`n" )
                        Remove-item $pathControl -Recurse -ErrorAction Ignore
                    }
                } 
            }
            elseif ($Uninstname -eq "Tobii Dynavox Phone") {
                $Phones = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
            ($_.Displayname -eq "Tobii Dynavox Phone Launcher") -or 
            ($_.Displayname -eq "Tobii Dynavox Phone Updater Service") -or
            ($_.Displayname -eq "Tobii Dynavox Phone") 
                } | Select-Object Displayname, UninstallString 
        
                stop-process -Name "*Tdx.Phone*" -Force
                foreach ( $Phone in $Phones) {
                    $Displayname = $Phone.Displayname
                    $Outputbox.Appendtext( "Removing - " + "$Displayname`r`n" )
                    $uninst = $Phone.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
                }
                $pathPhones = (
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Phone",
                    "$ENV:USERPROFILE\AppData\Local\Tobii Dynavox\Phone",
                    "$ENV:ProgramData\Tobii Dynavox\Phone",
                    "$ENV:ProgramData\Tobii Dynavox\Phone Launcher",
                    "$ENV:ProgramData\Tobii Dynavox\Phone Updater Service",
                    "C:\Program Files\Tobii Dynavox\Phone",
                    "C:\Program Files\Tobii Dynavox\Phone Launcher",
                    "C:\Program Files\Tobii Dynavox\Phone Updater Service",
                    "HKCU:\Software\Tobii Dynavox\Phone",
                    "HKCU:\Software\Tobii Dynavox\Phone Launcher",
                    "HKCU:\Software\Tobii Dynavox\Phone Updater Service")
                foreach ($pathPhone in $pathPhones) {
                    if ($pathPhone) {
                        $Outputbox.Appendtext(  "Removing - " + "$pathPhone`r`n" )
                        Remove-item $pathPhone -Recurse -ErrorAction Ignore
                    }
                } 
            }
            elseif ($Uninstname -eq "Tobii Dynavox Switcher") {
                $Switchers = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
            ($_.Displayname -eq "Tobii Dynavox Switcher Updater Service") -or
            ($_.Displayname -eq "Tobii Dynavox Switcher") 
                } | Select-Object Displayname, UninstallString 

                foreach ( $Switcher in $Switchers) {
                    $Displayname = $Switcher.Displayname
                    $Outputbox.Appendtext(  "Removing - " + "$Displayname`r`n" )
                    $uninst = $Switcher.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
                }
                $pathSwitchers = (
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Switcher",
                    "$ENV:ProgramData\Tobii Dynavox\Switcher",
                    "C:\Program Files\Tobii Dynavox\Switcher",
                    "C:\Program Files\Tobii Dynavox\Switcher Updater Service",
                    "HKCU:\Software\Tobii Dynavox\Switcher")
                foreach ($pathSwitcher in $pathSwitchers) {
                    if ($pathSwitcher) {
                        $Outputbox.Appendtext(  "Removing - " + "$pathSwitcher`r`n" )
                        Remove-item $pathSwitcher -Recurse -ErrorAction Ignore
                    }
                } 
            }
            elseif ($Uninstname -eq "Tobii Dynavox Talk") {
                $Talks = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
            ($_.Displayname -eq "Tobii Dynavox Talk Launcher") -or 
            ($_.Displayname -eq "Tobii Dynavox Talk Updater Service") -or
            ($_.Displayname -eq "Tobii Dynavox Talk") 
                } | Select-Object Displayname, UninstallString 

                foreach ( $Talk in $Talks) {
                    $Displayname = $Talk.Displayname
                    $Outputbox.Appendtext( "Removing - " + "$Displayname`r`n") 
                    $uninst = $Talk.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
                }
                $pathTalks = (
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Talk",
                    "$ENV:ProgramData\Tobii Dynavox\Talk",
                    "C:\Program Files\Tobii Dynavox\Talk",
                    "C:\Program Files\Tobii Dynavox\Talk Launcher",
                    "C:\Program Files\Tobii Dynavox\Talk Updater Service",
                    "HKCU:\Software\Tobii Dynavox\Talk",
                    "HKCU:\Software\Tobii Dynavox\Talk Launcher",
                    "HKCU:\Software\Tobii Dynavox\Talk Updater Service")
                foreach ($pathTalk in $pathTalks) {
                    if ($pathTalk) {
                        $Outputbox.Appendtext(  "Removing - " + "$pathTalk`r`n" )
                        Remove-item $pathTalk -Recurse -ErrorAction Ignore
                    }
                }  
            }
            elseif ($Uninstname -eq "Tobii Dynavox Browse (Beta)") {
                $Browsebetas = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
            ($_.Displayname -eq "Tobii Dynavox Browse Launcher (Beta)") -or 
            ($_.Displayname -eq "Tobii Dynavox Browse Updater Service (Beta)") -or
            ($_.Displayname -eq "Tobii Dynavox Browse (Beta)") -or
            ($_.Displayname -eq "Tobii Dynavox Prediction Service (Beta)") 
                } | Select-Object Displayname, UninstallString 

                foreach ( $Browsebeta in $Browsebetas) {
                    $Displayname = $Browsebeta.Displayname
                    $Outputbox.Appendtext( "Removing - " + "$Displayname`r`n" )
                    $uninst = $Browsebeta.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
                }
                $pathbrowsebetas = (
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Browse Beta",
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Prediction Service Beta",
                    "$ENV:ProgramData\Tobii Dynavox\Browse Beta",
                    "C:\Program Files\Tobii Dynavox\Browse Beta",
                    "C:\Program Files\Tobii Dynavox\Prediction Service Beta",
                    "HKCU:\Software\Tobii Dynavox\Browse Beta",
                    "HKCU:\Software\Tobii Dynavox\Browse Launcher Beta",
                    "HKCU:\Software\Tobii Dynavox\Browse Updater Service Beta",
                    "HKCU:\Software\Tobii Dynavox\Prediction Service Beta",
                    "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Browse Beta",
                    "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Browse Launcher Beta",
                    "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Browse Updater Service Beta")
                foreach ($pathbrowsebeta in $pathbrowsebetas) {
                    if ($pathbrowsebeta) {
                        $Outputbox.Appendtext(  "Removing - " + "$pathbrowsebeta`r`n" )
                        Remove-item $pathbrowsebeta -Recurse -ErrorAction Ignore
                    }
                }
            }
            elseif ($Uninstname -eq "Tobii Dynavox Control (Beta)") {
                $Controlbetas = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
            ($_.Displayname -eq "Tobii Dynavox Control Updater Service (Beta)") -or
            ($_.Displayname -eq "Tobii Dynavox Control (Beta)") 
                } | Select-Object Displayname, UninstallString 
        
                foreach ( $Controlbeta in $Controlbetas) {
                    $Displayname = $Controlbeta.Displayname
                    $Outputbox.Appendtext( "Removing - " + "$Displayname`r`n") 
                    $uninst = $Controlbeta.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
                }
                $pathControlbetas = (
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control Review",
                    "$ENV:ProgramData\Tobii Dynavox\Computer Control Review",
                    "C:\Program Files\Tobii Dynavox\Computer Control Review",
                    "HKCU:\Software\Tobii Dynavox\Computer Control Review",
                    "HKCU:\Software\Tobii Dynavox\Computer Control Launcher Review",
                    "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Computer Control Review",
                    "HKLM:\SOFTWARE\Wow6432Node\Tobii Dynavox\Computer Control Updater Service Review")
                foreach ($pathControlbeta in $pathControlbetas) {
                    if ($pathControlbeta) {
                        $Outputbox.Appendtext(  "Removing - " + "$pathControlbeta`r`n" )
                        Remove-item $pathControlbeta -Recurse -ErrorAction Ignore
                    }
                }
            }
            elseif ($Uninstname -eq "Tobii Dynavox Phone (Beta)") {
                $Phonebetas = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
            ($_.Displayname -eq "Tobii Dynavox Phone Launcher (Beta)") -or 
            ($_.Displayname -eq "Tobii Dynavox Phone Updater Service (Beta)") -or
            ($_.Displayname -eq "Tobii Dynavox Phone (Beta)") 
                } | Select-Object Displayname, UninstallString 
        
                stop-process -Name "*Tdx.Phone*" -Force
                foreach ( $Phonebeta in $Phonebetas) {
                    $Displayname = $Phonebeta.Displayname
                    $Outputbox.Appendtext( "Removing - " + "$Displayname`r`n") 
                    $uninst = $Phonebeta.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
                }
                $pathPhoneBetas = (
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Phone Beta",
                    "$ENV:USERPROFILE\AppData\Local\Tobii Dynavox\Phone Beta",
                    "$ENV:ProgramData\Tobii Dynavox\Phone Beta",
                    "$ENV:ProgramData\Tobii Dynavox\Phone Launcher Beta",
                    "$ENV:ProgramData\Tobii Dynavox\Phone Updater Service Beta",
                    "C:\Program Files\Tobii Dynavox\Phone Beta",
                    "C:\Program Files\Tobii Dynavox\Phone Launcher Beta",
                    "C:\Program Files\Tobii Dynavox\Phone Updater Service Beta",
                    "HKCU:\Software\Tobii Dynavox\Phone Beta",
                    "HKCU:\Software\Tobii Dynavox\Phone Launcher Beta",
                    "HKCU:\Software\Tobii Dynavox\Phone Updater Service Review")
                foreach ($pathPhoneBeta in $pathPhoneBetas) {
                    if ($pathPhoneBeta) {
                        $Outputbox.Appendtext( "Removing - " + "$pathPhoneBeta`r`n" )
                        Remove-item $pathPhoneBetas -Recurse -ErrorAction Ignore
                    }
                } 
            }
            elseif ($Uninstname -eq "Tobii Dynavox Switcher (Beta)") {
                $Switcherbetas = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
            ($_.Displayname -eq "Tobii Dynavox Switcher Updater Service (Beta)") -or
            ($_.Displayname -eq "Tobii Dynavox Switcher (Beta)") 
                } | Select-Object Displayname, UninstallString 
        
                foreach ( $Switcherbeta in $Switcherbetas) {
                    $Displayname = $Switcherbeta.Displayname
                    $Outputbox.Appendtext( "Removing - " + "$Displayname`r`n" )
                    $uninst = $Switcherbeta.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
                }
                $pathSwitcherBetas = (
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Switcher Beta",
                    "$ENV:ProgramData\Tobii Dynavox\Switcher Beta",
                    "C:\Program Files\Tobii Dynavox\Switcher Beta",
                    "C:\Program Files\Tobii Dynavox\Switcher Updater Service Beta",
                    "HKCU:\Software\Tobii Dynavox\Switcher Beta")
                foreach ($pathSwitcherBeta in $pathSwitcherBetas) {
                    if ($pathSwitcherBeta) {
                        $Outputbox.Appendtext( "Removing - " + "$pathSwitcherBeta`r`n" )
                        Remove-item $pathSwitcherBeta -Recurse -ErrorAction Ignore
                    }
                }
            }
            elseif ($Uninstname -eq "Tobii Dynavox Talk (Beta)") {
                $Talkbetas = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
            ($_.Displayname -eq "Tobii Dynavox Talk Launcher (Beta)") -or 
            ($_.Displayname -eq "Tobii Dynavox Talk Updater Service (Beta)") -or
            ($_.Displayname -eq "Tobii Dynavox Talk (Beta)") 
                } | Select-Object Displayname, UninstallString 

                foreach ( $Talkbeta in $Talkbetas) {
                    $Displayname = $Talkbeta.Displayname
                    $Outputbox.Appendtext("Removing - " + "$Displayname`r`n")
                    $uninst = $Talkbeta.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
                }
                $pathTalkBetas = (
                    "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Talk Beta",
                    "$ENV:ProgramData\Tobii Dynavox\Talk Beta",
                    "C:\Program Files\Tobii Dynavox\Talk Beta",
                    "C:\Program Files\Tobii Dynavox\Talk Launcher Beta",
                    "C:\Program Files\Tobii Dynavox\Talk Updater Service Beta",
                    "HKCU:\Software\Tobii Dynavox\Talk Beta",
                    "HKCU:\Software\Tobii Dynavox\Talk Launcher Beta",
                    "HKCU:\Software\Tobii Dynavox\Talk Updater Service Beta")
                foreach ($pathTalkBeta in $pathTalkBetas) {
                    if ($pathTalkBeta) {
                        Write-Host ( "Removing - " + "$pathTalkBeta`r`n" )
                        Remove-item $pathTalkBeta -Recurse -ErrorAction Ignore
                    }
                } 
            }
            elseif ($Uninstname -eq "Tobii Dynavox Telephony Bluetooth Driver") {
                $PhoneBTHs = Get-ChildItem -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | Get-ItemProperty  | Where-Object { 
            ($_.Displayname -eq "Tobii Dynavox Telephony Bluetooth Driver")
                } | Select-Object Displayname, UninstallString 

                foreach ( $PhoneBTH in $PhoneBTHs) {
                    $Displayname = $PhoneBTH.Displayname
                    $Outputbox.Appendtext("Removing - " + "$Displayname`r`n")
                    $uninst = $PhoneBTH.UninstallString -replace "msiexec.exe ", "" -Replace "/I", "" -Replace "/X", ""
                    start-process "msiexec.exe" -arg "/X $uninst /quiet /norestart" -Wait
                }
                $pathPhoneBTHs = (
                    "$ENV:ProgramData\Tobii Dynavox\Telephony Bluetooth Driver",
                    "HKCU:\Software\Tobii Dynavox\Telephony Bluetooth Driver")
                foreach ($pathPhoneBTH in $pathPhoneBTHs) {
                    if ($pathPhoneBTH) {
                        Write-Host ( "Removing - " + "$pathPhoneBTH`r`n" )
                        Remove-item $pathPhoneBTH -Recurse -ErrorAction Ignore
                    }
                } 
            

            }
        }
    }
    Remove-Variable * -ErrorAction SilentlyContinue
    Remove-Variable checkbox*

    $Outputbox.Appendtext("Done!`r`n")
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

    $TobiiVer = Get-WindowsDriver -Online | Where-Object { $_.ProviderName -match "Tobii" } | Select-Object Driver
    ForEach ($ver in $TobiiVer) {
        $outputBox.appendtext( "Removing Drivers - " + "$TobiiVer`r`n" )
        pnputil /delete-driver $ver.Driver /force /uninstall
    }
    stop-process -Name "*switcher*" -Force
    #Removes WC related folders
    $paths = (
        "C:\Program Files (x86)\Tobii Dynavox\Computer Control", #1 NEW
        "C:\Program Files (x86)\Tobii Dynavox\Eye Tracking Settings",	
        "C:\Program Files (x86)\Tobii Dynavox\Eye Assist",
        "C:\Program Files (x86)\Tobii Dynavox\Update Notifier",
        "C:\Program Files\Tobii Dynavox\Switcher", #2
        "C:\Program Files\Tobii\Tobii EyeX",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\App Switcher",
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control", #3
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Computer Control Bundle", #4
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Eye Tracking", #5
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\EyeAssist", #6
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Switcher", #7
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Shared Settings", #7 settings
        "$ENV:USERPROFILE\AppData\Roaming\Tobii Dynavox\Update Notifier",
        "$ENV:ProgramData\Tobii Dynavox\EyeAssist",
        "$ENV:ProgramData\Tobii Dynavox\Computer Control", #9
        "$ENV:ProgramData\Tobii Dynavox\Switcher", #10
        "$ENV:ProgramData\Tobii Dynavox\Update Notifier",
        "$ENV:ProgramData\Tobii\EulaHasBeenAccepted.txt",
        "$ENV:ProgramData\Tobii\Statistics", #11
        "$ENV:ProgramData\Tobii\Tobii Interaction", #12
        "$ENV:ProgramData\Tobii\Tobii Platform Runtime", #13
        "$ENV:ProgramData\HelloDMFT" )

    $PDKPath = "$ENV:ProgramData\Tobii\Tobii Platform Runtime\IS5LARGEPCEYE5"
    if (Test-path $PDKPath) {
        Write-Host "inside"
        Get-ChildItem -Path "$ENV:ProgramData\Tobii\Tobii Platform Runtime\IS5LARGEPCEYE5" -Recurse -af |  foreach-object { $_.FullName }
        Remove-Item "$ENV:ProgramData\Tobii\Tobii Platform Runtime\IS5LARGEPCEYE5" -Recurse 
    } 
    foreach ($path in $paths) {
        if (Test-Path $path) {
            $Outputbox.appendtext( "Removing - " + "$path`r`n" )
            Remove-Item "$path" -Recurse -Force # -ErrorAction Ignore
        }
    }
    $Keys = (
        "HKCU:\Software\Tobii\EyeAssist", #1
        "HKCU:\Software\Tobii\Update Notifier",
        "HKCU:\Software\Tobii Dynavox\Analytics", #2
        "HKCU:\Software\Tobii Dynavox\Computer Control", #3
        "HKLM:\SOFTWARE\WOW6432Node\Tobii Dynavox\Computer Control Updater Service", #4
        "HKLM:\SOFTWARE\WOW6432Node\Tobii\Update Notifier",
        "HKLM:\SOFTWARE\WOW6432Node\Tobii Dynavox\Computer Control Updater Service Review")

    foreach ($Key in $Keys) {
        if (test-path $Key) {
            $Outputbox.appendtext( "Removing - " + "$Key`r`n" )
            Remove-item $Key -Recurse -ErrorAction Ignore
        }
    }
    if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\OEMInformation\EyeTrackerModel") {
        Remove-ItemProperty -path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\OEMInformation" -Name "EyeTrackerModel"
    }

    Get-Item -Path "HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation"

    if (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation\EyeTrackerModel") {
        Remove-ItemProperty -path "HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation" -Name "EyeTrackerModel"
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
    $LogPath = "$ENV:USERPROFILE\AppData\Local\Temp"

    $ErrorPath = "$LogPath\ErrorLogs"
    if (!(Test-Path "$ErrorPath")) {
        $outputbox.appendtext( "Creating ErrorLogs folder.. `r`n")
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
    }
    else {
        
        foreach ($line in $tester4) {
            $array = $line.split("\")
            $path = [string]::Join("\", $array[0..($array.length - 2)]) 
            Add-Content -Path "$ErrorPath\InstallerError3.txt" -Value $path
        }
    }

    if ($Null -eq (Get-Content "$ErrorPath\InstallerError3.txt")) {
        $OutputBox.AppendText( "File is empty, no need to copy driver setup`r`n")
    }
    else {
        $test4 = Get-Content -Path "$ErrorPath\InstallerError3.txt"

        foreach ($tests4 in $test4) {
            $OutputBox.AppendText("Copy DriverSetup to specific path`r`n")
            New-Item -ItemType Directory -Force -Path $tests4
            $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "DriverSetup.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
            Set-Location $fpath
            Copy-Item -Path ("DriverSetup.exe") -Destination $tests4
        }
    }

   	$RegPath = "HKLM:\SOFTWARE\WOW6432Node\Tobii\EyeXConfig"
    $TempPath = "$ENV:USERPROFILE\AppData\Local\Temp\EyeXConfig.reg"
    if ((Test-Path -Path $RegPath) -and (!(Test-Path -path $TempPath))) {
        $Outputbox.Appendtext("Backup profiles in %temp%\EyeXConfig.reg`r`n" )
        Invoke-Command { reg export "HKLM\SOFTWARE\WOW6432Node\Tobii\EyeXConfig" $TempPath }
    }

    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
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
    else { $outputbox.appendtext( "No need to run BeforeUninstall.bat script`r`n") }

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
        
    $TobiiVer = Get-WindowsDriver -Online | Where-Object { $_.ProviderName -match "Tobii" } | Select-Object Driver
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

#A12
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

#A13 Copy licenses function. If any path to $Licensepaths exists, it will make a folder "Tobii Licenses", copy the licensefolders to the new folder(Does not contain the keys.xml, it is only the folder)
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

#B1 Function listapps - outputs all installed apps with the publisher Tobii
Function Listapps {
    $Outputbox.clear()
    #Creating folder
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion.ps1" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path

    if ($fpath.count -gt 0) {
        Set-Location $fpath
    }
    else { 
        $outputbox.appendtext("File $fileversion is missing!`r`n" )
    }
    $infofolder = "$fpath\infofolder"
    if (!(Test-Path "$infofolder")) {
        New-Item -Path "$infofolder" -ItemType Directory  
    }

    #Creating files
    if (!(Test-Path "$infofolder\SoftwareVersions.txt")) {
        New-Item -Path $infofolder -Name "SoftwareVersions.txt" -ItemType "file"
    }
    else {
        Clear-Content -Path "$infofolder\SoftwareVersions.txt"
    }
    $Outputbox.Appendtext( "Listing Tobii installed versions...`r`n" )
    $Outputbox.Appendtext( "Saving info in $infofolder\SoftwareVersions.txt`r`n" )
	
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
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
	
    $Listapps = Get-ChildItem -Recurse -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, 
    HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\,
    HKLM:\Software\WOW6432Node\Tobii\ |
    Get-ItemProperty | Where-Object { $_.Publisher -like '*Tobii*' } | Select-Object Displayname, Displayversion | Sort-Object Displayname | format-table -HideTableHeaders | out-string
    
    $TechListapps = Get-ChildItem -Recurse -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\, 
    HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\ | 
    Get-ItemProperty | Where-Object { 
        $_.Displayname -like '*Tobii Experience Software*' -or 
        $_.Displayname -like '*Tobii Device Drivers*' -or 
        $_.Displayname -like '*Tobii Eye Tracking For Windows*' 
    } | Select-Object Displayname, Displayversion | Sort-Object Displayname | format-table -HideTableHeaders   | out-string    
	
    $Listwindowsapp = Get-AppxPackage | Where-Object { ($_.Publisher -like '*Tobii*') -or
        ($_.Name -like '*Snap*') } | Select-Object name , version | format-table -HideTableHeaders | out-string
	
    $testpath = "C:\Program Files\Tobii\Tobii EyeX"
    #Fix installerpackageremovaltool.exe in driver setup folder
    if (Test-path $testpath) { 
        Set-Location "C:\Program Files\Tobii\Tobii EyeX"
        $Components = Get-childitem * -include platform_runtime_IS5GIBBONGAZE_service.exe, InstallerPackageRemovalTool.exe, Tobii.Configuration.exe, Tobii.EyeX.Engine.exe, Tobii.EyeX.Interaction.exe, Tobii.Service.exe, tobii_stream_engine.dll  | foreach-object { "{0}`t{1}" -f $_.Name, [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion }
    }

    if ($Firmware -match "IS5_Gibbon_Gaze") {
        $PDKversions = Get-ChildItem -Path "C:\Program Files\Tobii\Tobii EyeX" -Recurse -file -include "platform_runtime_IS5GIBBONGAZE_service.exe" | foreach-object { "{0}`t{1}" -f $_.Name, [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion }

    }
    elseif ($Firmware -match "IS5_Large_PC_Eye_5") {
        $PDKversions = Get-ChildItem -Path "C:\Program Files\Tobii\Tobii EyeX" -Recurse -file -include "platform_runtime_IS5LARGEPCEYE5_service.exe" | foreach-object { "{0}`t{1}" -f $_.Name, [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion }
    }	
    $ETVandModels = $Firmware | Select-String -Pattern "Firmware version", "Model"
    $ETSN = $Firmware | Select-String -Pattern "tobii-ttp"
    $ETSN = "$ETSN"
    $NewETSN = $ETSN -replace "Automatically selected eye tracker", ""

    $outputBox.AppendText( "TOBII INSTALLED SOFTWARE:$Listapps`r`n" )
    Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "TOBII INSTALLED SOFTWARE:$Listapps"

    if ($Listwindowsapp) {
        $outputBox.AppendText( "TOBII WINDOWS STORE APPS:$Listwindowsapp`r`n" )
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "$Listwindowsapp"
    }

    $outputbox.appendtext("TOBII TECH INSTALLED SW:$TechListapps")
    Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "$TechListapps"

    foreach ($component in $Components) {
        $outputbox.appendtext("$component`r`n")
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "$component"
    }

    foreach ($PDKversion in $PDKversions) {
        $outputbox.appendtext("$PDKversion`r`n")
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "$PDKversion"
    }

    foreach ($ETVandModel in $ETVandModels) {
        $outputbox.appendtext("$ETVandModel`r`n")
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "$ETVandModel"
    }
    
    if ($NewETSN) {
        $outputbox.appendtext("Eye Tracker S/N: $NewETSN `r`n")
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "$NewETSN"
    }

    $Drivers = Get-WmiObject Win32_PnPSignedDriver | Where-Object { ($_.DeviceName -match "Tobii Hello") -or ($_.DeviceName -match "Tobii Eye Tracker") } | Select-Object DeviceName, DriverVersion
    foreach ($Driver in $Drivers) {
        $drivername = $Driver.DeviceName
        $driverversion = $Driver.DriverVersion
        $outputbox.appendtext("$drivername $driverversion`r`n")
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "$drivername $driverversion"
        $outputbox.appendtext("`r`n")
    }

    $TobiiVer = Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Tobii\ProductInformation, HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation\, HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\OEMInformation\ 
    #$OutputBox.AppendText("$TobiiVer")
     
    if ($TobiiVer.EyeTrackerModel -eq "PCEye5") {
        $OutputBox.AppendText("Eye tracker is " + $TobiiVer.EyeTrackerModel + " IS514`r`n")
    }
    elseif ($TobiiVer.ProductType -eq "TDG16") {
        $OutputBox.AppendText("Eye tracker is " + $TobiiVer.ProductType + " TDG16`r`n")
    }
    elseif ($TobiiVer.ProductType -eq "TDH10") {
        $OutputBox.AppendText("Eye tracker is " + $TobiiVer.ProductType + " TDH10`r`n")
    }
    elseif ($TobiiVer.ProductType -eq "TDTW7") {
        $OutputBox.AppendText("Eye tracker is " + $TobiiVer.ProductType + " TDTW7`r`n")
    }
    elseif ($TobiiVer.ProductType -eq "TDG10") {
        $OutputBox.AppendText("Eye tracker is " + $TobiiVer.ProductType + " TDG10`r`n")
    }
    elseif ($TobiiVer.ProductType -eq "I-Series" -and $TobiiVer1.ProductModel -eq "I-12+") {
        $OutputBox.AppendText("Eye tracker is " + $TobiiVer.ProductType + " TDI12-xxxxx`r`n")
    }
    elseif ($TobiiVer.EyeTrackerModel -eq "EM12") {
        $OutputBox.AppendText("Eye tracker is " + $TobiiVer.EyeTrackerModel + " TEM12`r`n")
    }
    elseif ($TobiiVer.EyeTrackerModel -eq "PCEye2") {
        $OutputBox.AppendText("Eye tracker is  " + $TobiiVer.EyeTrackerModel + " PCEGO or PCEye Mini`r`n")
    }
    elseif ($TobiiVer.EyeTrackerModel -eq "PCEyeExplore") {
        $OutputBox.AppendText("Eye tracker is  " + $TobiiVer.EyeTrackerModel + " PCEyeExplore`r`n")
    }
    else {
        $OutputBox.AppendText( "No match," + $TobiiVer.EyeTrackerModel + "and " + $TobiiVer.ProductType + "`r`n") 
    }

    pnputil /enum-drivers > $infofolder\systemDrivers.txt
    $TobiiDrivers = Get-WindowsDriver -Online | Where-Object { $_.ProviderName -match "Tobii" }  | Select-Object Driver , OriginalFileName
    ForEach ($drivers in $TobiiDrivers) {
        $inf = $drivers.Driver 
        $List = $drivers.OriginalFileName
        $List = $List.Replace("C:\Windows\System32\DriverStore\FileRepository\", "")
        $outputbox.appendtext("$inf : $List `r`n")
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "$inf : $List"
    }
   
    $getdeviceids = $null
    $getdeviceids2 = $null
    $getdeviceids = Get-WmiObject Win32_USBControllerDevice | ForEach-Object { [wmi]($_.Dependent) } | Where-Object DeviceID -Like "*Tobii*" | Select-object DeviceID
    $getdeviceids2 = Get-CimInstance Win32_PnPSignedDriver | Where-Object Description -Like "*WinUSB Device*" | Select-Object DeviceID
    # gwmi Win32_USBControllerDevice |%{[wmi]($_.Dependent)} | Sort Manufacturer,Description,DeviceID | Ft -GroupBy Manufacturer Description,Service,DeviceID | out-file c:\VidPid.txt

    Foreach ($getdeviceid in $getdeviceids) {
        $outputbox.appendtext("$getdeviceid `r`n")
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "$getdeviceid"
    }
    Start-Sleep -s 5
    Foreach ($getdeviceid2 in $getdeviceids2) {
        $outputbox.appendtext("$getdeviceid2`r`n")
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "$getdeviceid2"
    }	

    $installedFrameworks = @()
    if (IsKeyPresent "HKLM:\Software\Microsoft\.NETFramework\Policy\v1.0" "3705") { 
        $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 1.0`r`n")
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 1.0"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v1.1.4322" "Install") { 
        $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 1.1`r`n") 
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 1.1"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v2.0.50727" "Install") { 
        $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 2.0`r`n")
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 2.0"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.0\Setup" "InstallSuccess") { 
        $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 3.0`r`n") 
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 3.0"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.5" "Install") { 
        $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 3.5`r`n" ) 
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 3.5"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client" "Install") { 
        $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 4.0c`r`n" ) 
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.0c"
    }
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Install") { 
        $installedFrameworks += $outputbox.appendtext("Installed .Net Framework 4.0`r`n" ) 
        Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.0"
    }

    $result = -1
    if (IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client" "Install" -or IsKeyPresent "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Install") {
        # .net 4.0 is installed
        $result = 0
        $version = GetFrameworkValue "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full" "Release"
        
        if ($version -ge 528040 -Or $version -ge 528372 -Or $version -ge 528049) {
            # .net 4.8
            $outputbox.appendtext( "Installed .Net Framework 4.8`r`n")
            Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.8"
            $result = 10
        }
        elseif ($version -ge 461808 -Or $version -ge 461814) {
            # .net 4.7.2
            $outputbox.appendtext("Installed .Net Framework 4.7.2`r`n")
            Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.7.2"
            $result = 9
        }
        elseif ($version -ge 461308 -Or $version -ge 461310) {
            # .net 4.7.1
            $outputbox.appendtext( "Installed .Net Framework 4.7.1`r`n")
            Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.7.1"
            $result = 8
        }
        elseif ($version -ge 460798 -Or $version -ge 460805) {
            # .net 4.7
            $outputbox.appendtext( "Installed .Net Framework 4.7`r`n")
            Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.7"
            $result = 7
        }
        elseif ($version -ge 394802 -Or $version -ge 394806) {
            # .net 4.6.2
            $outputbox.appendtext( "Installed .Net Framework 4.6.2`r`n")
            Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.6.2"
            $result = 6
        }
        elseif ($version -ge 394254 -Or $version -ge 394271) {
            # .net 4.6.1
            $outputbox.appendtext( "Installed .Net Framework 4.6.1`r`n")
            Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.6.1"
            $result = 5
        }
        elseif ($version -ge 393295 -Or $version -ge 393297) {
            # .net 4.6
            $outputbox.appendtext( "Installed .Net Framework 4.6`r`n")
            Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.6"
            $result = 4
        }
        elseif ($version -ge 379893) {
            # .net 4.5.2
            $outputbox.appendtext( "Installed .Net Framework 4.5.2`r`n")
            Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.5.2"
            $result = 3
        }
        elseif ($version -ge 378675) {
            # .net 4.5.1
            $outputbox.appendtext( "Installed .Net Framework 4.5.1`r`n")
            Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.5.1"
            $result = 2
        }
        elseif ($version -ge 378389) {
            # .net 4.5
            $outputbox.appendtext( "Installed .Net Framework 4.5`r`n")
            Add-Content -path "$infofolder\SoftwareVersions.txt" -Value "Installed .Net Framework 4.5"
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
Function HWInfo {
    $outputBox.clear()
    #Creating folder
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion.ps1" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path

    if ($fpath.count -gt 0) {
        Set-Location $fpath
    }
    else { 
        $outputbox.appendtext("File $fileversion is missing!`r`n" )
    }
    $infofolder = "$fpath\infofolder"
    if (!(Test-Path "$infofolder")) {
        New-Item -Path "$infofolder" -ItemType Directory  
    }
    else {
        Remove-Item -Path $infofolder\* -Recurse 
        $OutputBox.AppendText( "InfoFolder is already created.`r`n")
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
        !(Test-Path "$infofolder\DeviceInfo.txt") -or
        !(Test-Path "$infofolder\AllSW.txt") -or
        !(Test-Path "$infofolder\ProcessPIDDrivers.txt")
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
        New-Item -Path $infofolder -Name "AllSW.txt" -ItemType "file"
        New-Item -Path $infofolder -Name "ProcessPIDDrivers.txt" -ItemType "file"
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
        Clear-Content -Path "$infofolder\AllSW.txt"
        Clear-Content -Path "$infofolder\ProcessPIDDrivers.txt"
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
    Add-Content -path "$infofolder\AllSW.txt" -Value $Listapps
    Add-Content -path "$infofolder\ProcessPIDDrivers.txt" -Value $GetOtherInfo

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

#B3 Lists currently active tobii processes & services
Function GetOtherInfo {
    $outputBox.clear()
    $GetProcess = get-process "*GazeSelection*", "*Tobii*" | Select-Object Processname | Format-table -hidetableheaders | Out-string
    $GetServices = Get-Service -Name '*Tobii*' | Select-Object Name, Status | Format-table -hidetableheaders | Out-string

    $outputBox.appendtext( "Listing active Tobii processes...`r`n" )
    if ($GetProcess) {
        $outputbox.appendtext("ACTIVE PROCESSES:$GetProcess`r`n")
    }
    if ($GetServices) {
        $outputbox.appendtext("ACTIVE Services:$GetServices`r`n")
    }

    $outputbox.appendtext("Done!`r`n")
}

#B4 Stops all currently active tobii processes
Function RestartProcesses {
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

#B5
Function ETfw {
    $outputBox.clear()
    $outputBox.appendtext( "Checking Eye tracker Firmware...`r`n" )
    Get-Service -Name 'Tobii Service'  | Start-Service
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "FWUpgrade32.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        try { 
            $erroractionpreference = "Stop"
            $Firmware = .\FWUpgrade32.exe --auto --info-only 
        }
        Catch [System.Management.Automation.RemoteException] {
            $outputbox.appendtext("No Eye Tracker Connected`r`n")
        }
        $outputbox.appendtext("$Firmware`r`n")
        if ($null -ne $Firmware) {
            $path = "C:\Program Files (x86)\Tobii\Service"
            if (Test-Path $path) {
                #If first answer equals yes or no
                $answer1 = $wshell.Popup("This will upgrade IS4 firmware.`r`nAre you sure you want to continue?`r`n", 0, "Caution", 48 + 4)
                if ($answer1 -eq 6) {
                    $Outputbox.Appendtext( "Starting upgrade... Do NOT close this window while it is in progress..`r`n" )
                }
                elseif ($answer1 -ne 6) {
                    $Outputbox.Appendtext( "Action canceled`r`n" )
                    Return
                }
                Set-Location -path $path
                if ($Firmware -match "PCE1M") {
                    #PCEye Mini: tobii-ttp://PCE1M-010106010685
                    $outputbox.appendtext("Upgrading PCEye mini FW..`r`n")
                    $PCEyeMini = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4pceyemini_firmware_2.27.0-4014648.tobiipkg" --no-version-check
                    $outputbox.appendtext("$PCEyeMini`r`n")
                    $outputbox.appendtext("Upgrade is Done! `r`n")
                }
                elseif ($Firmware -match "IS4_Large_102") {
                    $outputbox.appendtext("Upgrading PCEye Plus FW..`r`n")
                    $PCEyePlus = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4large102_firmware_2.27.0-4014648.tobiipkg" --no-version-check
                    $outputbox.appendtext("$PCEyePlus`r`n")
                    $outputbox.appendtext("Upgrade is Done! `r`n")
                }
                elseif ($Firmware -match "IS4_Large_Peripheral") {
                    $outputbox.appendtext("Upgrading 4C FW..`r`n")
                    $4C = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii\Tobii Firmware\is4largetobiiperipheral_firmware_2.27.0-4014648.tobiipkg" --no-version-check
                    $outputbox.appendtext("$4C`r`n")
                    $outputbox.appendtext("Upgrade is Done! `r`n")
                }
                elseif ($Firmware -match "IS4_Base_I-series") {
                    $outputbox.appendtext("Upgrading I-Series+ FW..`r`n")
                    $ISeries = .\FWUpgrade32.exe --auto "C:\Program Files (x86)\Tobii Dynavox\Gaze Interaction\Eye Tracker Firmware Releases\IS4B1\is4iseriesb_firmware_2.9.0.tobiipkg" --no-version-check
                    $outputbox.appendtext("$ISeries`r`n")
                    $outputbox.appendtext("Upgrade is Done. Restart ET through Control Center `r`n")
                }
                elseif ($Firmware -match "tet-tcp") {
                    #Tobii Firmware Upgrade Tool Automatically selected eye tracker tet-tcp://172.28.195.1 Failed to open file
                    $outputbox.appendtext("ET model is IS20. Use ET Browser to upgrade. Make sure that Bonjure is installed.`r`n")
                }
                #Get-Service -Name 'Tobii Service'  | Where-Object { $_.Status -ne "Running" } | Start-Service
                Get-Service -Name 'Tobii Service'  | stop-service 
                Get-Service -Name 'Tobii Service'  | Start-Service
            } 
        }
        else {
            $outputbox.appendtext("Could not read Firmware version!`r`n" )
        }
    } 
    else {
        $outputbox.appendtext("File FWUpgrade32.exe is missing!`r`n" )
    }

    $outputbox.appendtext("Done! `r`n")
} 

#B6
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

#B7
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

#B8
Function SMBios {
    $outputBox.clear()
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'SMBios tool'
    $msg = "Press 1 to run getSMBIOSvalues.cmd, `r`n 2 setName.cmd, `r`n 3 setSerialNumber.cmd, `r`n 4 setVendor.cmd, `r`n 5 GB2SmbiosTool"
    $b = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "getSMBIOSvalues.cmd" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        Set-Location $fpath

        if ($b -match "1") { 
            Start-Process -FilePath .\getSMBIOSvalues.cmd
        }
        elseif ($b -match "2") {
            Start-Process -FilePath .\setName.cmd
        }
        elseif ($b -match "3") { 
            Start-Process -FilePath .\setSerialNumber.cmd
        }
        elseif ($b -match "4") { 
            Start-Process -FilePath .\setVendor.cmd
        }
        elseif ($b -match "5") {
            $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "GB2SmbiosTool.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
            if ($fpath.count -gt 0) {
                Set-Location $fpath
                Start-Process cmd -Verb runAs "/c `"GB2SmbiosTool.exe`"" 
            }
            else { 
                $outputbox.appendtext("File GB2SmbiosTool.exe is missing!`r`n" )
            }
        }
        else { $outputbox.appendtext("N/A`r`n") }
    }
    else { 
        $outputbox.appendtext("File getSMBIOSvalues.cmd is missing!`r`n" )
    }
    $outputbox.appendtext("Done!`r`n")
}

#B9
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
        Set-Location $fpath
        try { 
            $erroractionpreference = "Stop"
            $Firmware = .\FWUpgrade32.exe --auto --info-only 
            $outputBox.appendtext( "Eye Tracker info: $Firmware`r`n" )
        }
        Catch [System.Management.Automation.RemoteException] {
            $outputbox.appendtext("No eye tracker could be found`r`n")
        }
    }
    else { 
        $outputbox.appendtext("File FWUpgrade32.exe is missing!`r`n" )
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
    
    $outputbox.appendtext("Done! `r`n")
}

#B10
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

#B11
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

#B12
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

    $TobiiVer = Get-WindowsDriver -Online | Where-Object { $_.ProviderName -match "Tobii" } | Select-Object Driver, OriginalFileName

    ForEach ($ver in $TobiiVer) {
        $outputBox.appendtext( "Removing - " + "$ver`r`n" )
        pnputil /delete-driver $ver.Driver /force /uninstall
    }
    $outputbox.appendtext("Done!`r`n")
}

#B13
Function SETest {
    $outputBox.clear()
    $outputBox.appendtext( "running Stream Engine Test...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "tests.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        .\tests.exe
    }
    else { 
        $outputbox.appendtext("File tests.exe is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}

#B14
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

#B15
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

#B16
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

#B17
Function ETConnection {
    $outputBox.clear()
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion.ps1" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        Set-Location $fpath
    }
    $infofolder = "$fpath\ETConnectionOutput.txt"
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
    $outputbox.appendtext("Done! `r`n")
}

#B18
Function EAProfileCreation {
    $outputBox.clear()
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'ET Profile'
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

#B19
Function ETSamples {
    $outputBox.clear()
    $outputBox.appendtext( "Starting TD region interaction sample...`r`n" )
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$fileversion.ps1" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    $fpathsample = Get-ChildItem -Path $PSScriptRoot -Filter "Tdx.EyeTracking.RegionInteraction.EyeAssist.Sample.exe" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
   

    $ConResults = "$fpath\ETConnectionSample.txt"
    $SamResults = "$fpath\SampleResults.txt"
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
    $outputbox.appendtext("Done! `r`n")
}

#B20
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

#B21
Function Sleeper {
    $outputBox.clear()
    $outputBox.appendtext( "Starting Sleeper.exe then load configuration and start sleep...`r`n" )
    $fpath = (Get-ChildItem -Path "$PSScriptRoot" -Filter "Sleeper.exe" -Recurse).FullName | Split-Path 
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        start-process .\Sleeper.exe
    }
    else { 
        $outputbox.appendtext("File Sleeper.exe is missing!`r`n" )
    }	
    $outputbox.appendtext("Done! `r`n")
}

#B22 Deploy
Function Deployment {
    #USB namings:
    #"ISeries_MPD"         "ISeries_MPB"
    #"Indi_AD"             "Indi_AB"
    #"Indi_BD"             "Indi_BB",
    #"ISeriesD"            "ISeriesB"
    #"I-110D"              "I-110B"
    #"Indi_7_AD"           "Indi_7_AB"
    #"Indi_7_BD"           "Indi_7_BB"
    #"Surface_SP6D"        "Surface_SP6"
    #"Surface_SP7D"        "Surface_SP7" 
    #"I-Series+D"          "I-SeriesB"
    #"I-110-8W10D"         "I-110-8W10B"
    #"I-110-8W11D"         "I-110-8W11B"

    $Deployments = @("ISeries_MP", "Indi_A", "Indi_B", "ISeries", "I-Series+", "I-110", "Indi_7_A", "Indi_7_B", "Surface_Pro_SP6", "Surface_Pro_SP7", "I-110-850_W11" , "I-110-850_W10" )
    Write-host "Deployments list: $Deployments"

    $availableUSB = (@(Get-Volume | Where-Object DriveType -eq Removable | Where-Object FileSystemType -eq NTFS |  Select-Object FileSystemLabel).FileSystemLabel ) -replace ".$"
    write-host "Available drives: $availableUSB"

    foreach ($availableUSBs in $availableUSB) {
        if ( $availableUSBs -match "Surface_SP6") {
            $ComparesUSB = "Surface_Pro_SP6"
        }
        elseif ( $availableUSBs -match "Surface_SP7") {
            $ComparesUSB = "Surface_Pro_SP7"
        }
        elseif ( $availableUSBs -match "I-110-8W11") {
            $ComparesUSB = "I-110-850_W11"
        }
        elseif ( $availableUSBs -match "I-110-8W10") {
            $ComparesUSB = "I-110-850_W10"
        }
        else {
            $ComparesUSB = (Compare-Object -DifferenceObject $Deployments -ReferenceObject $availableUSB -CaseSensitive -ExcludeDifferent -IncludeEqual | Select-Object InputObject).InputObject
        }
        Write-Host "Available deploys that match available USBs: $ComparesUSB"

        foreach ($ComparesUSBs in $ComparesUSB) {
            write-host "Selecting deploy: $ComparesUSBs" 

            $DeployName = "$availableUSBs" + "D"

            if ("$ComparesUSBs" -eq "I-Series+") {
                $NewComparesUSB = $ComparesUSBs.Replace( "+", "")
                $BootName = "$NewComparesUSB" + "B"
                Write-Host "Device is I-Series+"
            }
            elseif ("$ComparesUSBs" -eq "Surface_Pro_SP6") {
                $NewComparesUSB = $ComparesUSBs.Replace("_Pro", "")
                $BootName = $NewComparesUSB
            }
            elseif ("$ComparesUSBs" -eq "Surface_Pro_SP7") {
                $NewComparesUSB = $ComparesUSBs.Replace("_Pro", "")
                $BootName = $NewComparesUSB
            }
            elseif ("$ComparesUSBs" -eq "I-110-850_W11") {
                $NewComparesUSB = $ComparesUSBs.Replace("50_", "")
                $BootName = $NewComparesUSB + "B"
            }
            elseif ("$ComparesUSBs" -eq "I-110-850_W10") {
                $NewComparesUSB = $ComparesUSBs.Replace("50_", "")
                $BootName = $NewComparesUSB + "B"
            }
            else {
                $BootName = "$ComparesUSBs" + "B"
                Write-Host "Device is $BootName"
            }
            write-host "Setting deploy name to $DeployName and $BootName"

            write-host "DeployName$DeployName"
            if ($DeployName -match "Surface_Pro" -or $DeployName -match "ISeries" -or $ComparesUSBs -match "I-110-850_W" -or $ComparesUSBs -match "Surface_Pro_SP7") {
                $Download = ((Get-ChildItem -Path "$env:USERPROFILE\Downloads" | Where-Object { $_.Name -match "$ComparesUSBs" }).Name ) -replace ".7z", ""
                Write-Host "Found deploy: $Download"
            }
            else {
                $Download = ((Get-ChildItem -Path "D:\" | Where-Object { $_.Name -match "$ComparesUSBs" }).Name ) -replace ".7z", ""
            }
            Write-Host "Found deploy: $Download"

            Set-Location  "C:\Program Files\7-Zip"
            #Set-Location  "C:\Program Files (x86)\7-Zip" 
            if ($Download) {

                # clear content in USB
                Write-Host "Formatting both $DeployName and $BootName"
                Format-Volume -FriendlyName $DeployName -FileSystem NTFS -NewFileSystemLabel $DeployName
                Format-Volume -FriendlyName $BootName -FileSystem FAT32 -NewFileSystemLabel $BootName

                # Find and select driver for USB
                $getDepLetter = (Get-Volume | Where-Object { ($_.FileSystemLabel -eq "$DeployName") }).DriveLetter
                $getBootLetter = (Get-Volume | Where-Object { ($_.FileSystemLabel -eq "$BootName") }).DriveLetter
                Write-Host "Found following driver letters: $getDepLetter & $getBootLetter "

                # Select correct deploy from Download
                write-host "Unpacking..."
                if ($DeployName -match "Surface_Pro" -or $DeployName -match "ISeries" -or $ComparesUSBs -match "I-110-850_W" -or $ComparesUSBs -match "Surface_Pro_SP7") {
                    $unpack = .\7z.exe x $env:USERPROFILE\Downloads\"$Download".7z -o"$getDepLetter":\  -p5rd4c5vgcTvuKC -r
                }
                else {
                    $unpack = .\7z.exe x D:\"$Download".7z -o"$getDepLetter":\  -p5rd4c5vgcTvuKC -r
                }
        
                Write-Host "Unpack folder is $unpack"

                # Move files to its proper path
                $USBpath = "$getDepLetter':\'$Download" -replace "'", ""
                $newgetDepLetter = "$getDepLetter':\'" -replace "'", ""
                $newgetBootLetter = "$getBootLetter':\'" -replace "'", ""

                Get-ChildItem -Path "$USBpath\winpe" -Recurse | Move-Item -Destination $newgetBootLetter
                Get-ChildItem -Path "$USBpath\deploy" -Recurse | Move-Item -Destination $newgetDepLetter 
                Remove-Item -Path "$USBpath" -Force -Recurse
                Write-Host "Moving folders to its right path and cleaning"
            }
            else {
                Write-Host "No Deploy match $ComparesUSBs"
            }
        }
    }

}

#B23
Function LogCollector {
    $outputBox.clear()
    $outputbox.appendtext("Start `r`n")
    $fpath = Get-ChildItem -Path $PSScriptRoot -Filter "$LogCollectorTool" -Recurse -erroraction SilentlyContinue | Select-Object -expand Fullname | Split-Path
    if ($fpath.count -gt 0) {
        Set-Location $fpath
        powershell.exe "$fpath\$LogCollectorTool"
    }
    else { 
        $outputbox.appendtext("File $LogCollectorTool is missing!`r`n" )
    }
    $outputbox.appendtext("Done! `r`n")
}




#Windows forms
$Optionlist = @("Remove Progressive Sweet", "Remove PCEye5 Bundle", "Remove all ET SW", "Remove WC&GP Bundle", "Remove VC++", "Remove PCEye Package", "Remove Communicator", "Remove Compass", "Remove TGIS only", "Remove TGIS profile calibrations", "Remove all users C5", "Backup Gaze Interaction", "Copy License")
$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(600, 590)
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
$Button.Size = New-Object System.Drawing.Size(110, 30)
$Button.Text = "Start"
$Button.Font = New-Object System.Drawing.Font ("" , 12, [System.Drawing.FontStyle]::Regular)
$Form.Controls.Add($Button)
$Button.Add_Click{ selectedscript }

#B1 Button1 "List Tobii Software"
$Button1 = New-Object System.Windows.Forms.Button
$Button1.Location = New-Object System.Drawing.Size(420, 0)
$Button1.Size = New-Object System.Drawing.Size(150, 30)
$Button1.Text = "All versions"
$Button1.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button1)
$Button1.Add_Click{ ListApps }

#B2 Button2 "HW Info"
$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(420, 30)
$Button2.Size = New-Object System.Drawing.Size(150, 30)
$Button2.Text = "HW Info"
$Button2.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button2)
$Button2.Add_Click{ HWInfo }

#B3 Button3 "List active Tobii processes"
$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Size(420, 60)
$Button3.Size = New-Object System.Drawing.Size(150, 30)
$Button3.Text = "Get Services"
$Button3.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button3)
$Button3.Add_Click{ GetOtherInfo }

#B4 Button4 Restart Services
$Button4 = New-Object System.Windows.Forms.Button
$Button4.Location = New-Object System.Drawing.Size(420, 90)
$Button4.Size = New-Object System.Drawing.Size(150, 30)
$Button4.Text = "Restart Services"
$Button4.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button4)
$Button4.Add_Click{ RestartProcesses }

#B5 Button5 "ET fw"
$Button5 = New-Object System.Windows.Forms.Button
$Button5.Location = New-Object System.Drawing.Size(420, 120)
$Button5.Size = New-Object System.Drawing.Size(150, 30)
$Button5.Text = "Firmware v / Upgrade"
$Button5.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button5)
$Button5.Add_Click{ ETfw }

#B6 Button6 "Show Track status"
$Button6 = New-Object System.Windows.Forms.Button
$Button6.Location = New-Object System.Drawing.Size(420, 150)
$Button6.Size = New-Object System.Drawing.Size(150, 30)
$Button6.Text = "Show/hide Track Status"
$Button6.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button6)
$Button6.Add_Click{ TrackStatus }

#B7 Button7 "WCF"
$Button7 = New-Object System.Windows.Forms.Button
$Button7.Location = New-Object System.Drawing.Size(420, 180)
$Button7.Size = New-Object System.Drawing.Size(150, 30)
$Button7.Text = "WCF"
$Button7.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button7)
$Button7.Add_Click{ WCF }

#B8 Button8 "SMBios"
$Button8 = New-Object System.Windows.Forms.Button
$Button8.Location = New-Object System.Drawing.Size(420, 210)
$Button8.Size = New-Object System.Drawing.Size(150, 35)
$Button8.Text = "SMBIOS"
$Button8.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button8)
$Button8.Add_Click{ SMBios }

#B9 Button9 "Reset IS5 to bootloader"
$Button9 = New-Object System.Windows.Forms.Button
$Button9.Location = New-Object System.Drawing.Size(420, 245)
$Button9.Size = New-Object System.Drawing.Size(75, 35)
$Button9.Text = "Reset ET"
$Button9.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button9)
$Button9.Add_Click{ resetBOOT }

#B10 Button10 "RetrieveUnreleased"
$Button10 = New-Object System.Windows.Forms.Button
$Button10.Location = New-Object System.Drawing.Size(495, 245)
$Button10.Size = New-Object System.Drawing.Size(75, 35)
$Button10.Text = "RetrieveUN"
$Button10.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button10)
$Button10.Add_Click{ RetrieveUnreleased }

#B11 Button11 "Delete services"
$Button11 = New-Object System.Windows.Forms.Button
$Button11.Location = New-Object System.Drawing.Size(420, 280)
$Button11.Size = New-Object System.Drawing.Size(75, 35)
$Button11.Text = "Delete Services"
$Button11.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button11)
$Button11.Add_Click{ DeleteServices }

#B12 Button12 "Remove Drivers"
$Button12 = New-Object System.Windows.Forms.Button
$Button12.Location = New-Object System.Drawing.Size(495, 280)
$Button12.Size = New-Object System.Drawing.Size(75, 35)
$Button12.Text = "Delete Drivers"
$Button12.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button12)
$Button12.Add_Click{ RemoveDrivers }

#B13 Button13 "StreamEngineTest"
$Button13 = New-Object System.Windows.Forms.Button
$Button13.Location = New-Object System.Drawing.Size(420, 315)
$Button13.Size = New-Object System.Drawing.Size(75, 35)
$Button13.Text = "SE-Test"
$Button13.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button13)
$Button13.Add_Click{ SETest }

#B14 Button14 "InternalSE"
$Button14 = New-Object System.Windows.Forms.Button
$Button14.Location = New-Object System.Drawing.Size(495, 315)
$Button14.Size = New-Object System.Drawing.Size(75, 35)
$Button14.Text = "Internal SE"
$Button14.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button14)
$Button14.Add_Click{ InternalSE }

#B15 Button15 "SetDebugLogging"
$Button15 = New-Object System.Windows.Forms.Button
$Button15.Location = New-Object System.Drawing.Size(420, 350)
$Button15.Size = New-Object System.Drawing.Size(75, 35)
$Button15.Text = "DebugLog"
$Button15.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button15)
$Button15.Add_Click{ SetDebugLogging }

#B16 Button16 "Diagnostic"
$Button16 = New-Object System.Windows.Forms.Button
$Button16.Location = New-Object System.Drawing.Size(495, 350)
$Button16.Size = New-Object System.Drawing.Size(75, 35)
$Button16.Text = "RIDiagnostic"
$Button16.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button16)
$Button16.Add_Click{ Diagnostic }

#B17 Button17 "Check ET connection through Service"
$Button17 = New-Object System.Windows.Forms.Button
$Button17.Location = New-Object System.Drawing.Size(420, 385)
$Button17.Size = New-Object System.Drawing.Size(75, 35)
$Button17.Text = "ET con."
$Button17.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button17)
$Button17.Add_Click{ ETConnection }

#B18 Button18 "EAProfileCreation"
$Button18 = New-Object System.Windows.Forms.Button
$Button18.Location = New-Object System.Drawing.Size(495, 385)
$Button18.Size = New-Object System.Drawing.Size(75, 35)
$Button18.Text = "EA Profile"
$Button18.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button18)
$Button18.Add_Click{ EAProfileCreation }

#B19 Button19 "ETSamples"
$Button19 = New-Object System.Windows.Forms.Button
$Button19.Location = New-Object System.Drawing.Size(420, 420)
$Button19.Size = New-Object System.Drawing.Size(75, 35)
$Button19.Text = "RI Samples"
$Button19.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button19)
$Button19.Add_Click{ ETSamples }

#B20 Button20 "BatteryLog"
$Button20 = New-Object System.Windows.Forms.Button
$Button20.Location = New-Object System.Drawing.Size(495, 420)
$Button20.Size = New-Object System.Drawing.Size(75, 35)
$Button20.Text = "Battery Log"
$Button20.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button20)
$Button20.Add_Click{ BatteryLog }

#B21 Button21 "Sleeper"
$Button21 = New-Object System.Windows.Forms.Button
$Button21.Location = New-Object System.Drawing.Size(420, 455)
$Button21.Size = New-Object System.Drawing.Size(75, 35)
$Button21.Text = "Sleeper"
$Button21.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button21)
$Button21.Add_Click{ Sleeper }

#B21 Button21 "Deployment"
$Button21 = New-Object System.Windows.Forms.Button
$Button21.Location = New-Object System.Drawing.Size(495, 455)
$Button21.Size = New-Object System.Drawing.Size(75, 35)
$Button21.Text = "Deployment"
$Button21.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button21)
$Button21.Add_Click{ Deployment }

#B22 Button22 "LogCollector"
$Button22 = New-Object System.Windows.Forms.Button
$Button22.Location = New-Object System.Drawing.Size(420, 490)
$Button22.Size = New-Object System.Drawing.Size(75, 35)
$Button22.Text = "LogCollector"
$Button22.Font = New-Object System.Drawing.Font ("" , 8, [System.Drawing.FontStyle]::Regular)
$form.Controls.add($Button22)
$Button22.Add_Click{ LogCollector }

#Form name + activate form.
$Form.Text = $fileversion
$Form.Add_Shown( { $Form.Activate() })
$Form.ShowDialog()