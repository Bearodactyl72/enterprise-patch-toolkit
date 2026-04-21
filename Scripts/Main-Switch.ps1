# DOTS formatting comment

$scriptRoot    = $scriptPath
$patchRoot     = "M:\Share\VMT\Patches" # Patch folder goes here
$listPathRoot  = "$env:USERPROFILE\Desktop\Lists"   # List  folder goes here

$officeAppDir64  = "C:\Program Files\Microsoft Office\root\Office16"
$officeAppDir32  = "C:\Program Files (x86)\Microsoft Office\Office16"
$officeProdDir64 = "C:\Program Files\Common Files\Microsoft Shared"       #??
$officeProdDir32 = "C:\Program Files (x86)\Common Files\Microsoft Shared"
$officeDir64     = "$patchRoot\Office64bit"
$officeDir32     = "$patchRoot\Office32bit"
$installLine     = @()

$adobeCCProcesses = @("AcroCEF", "acrodist", "acrotray", "AcroRd32", "AdobeARM", "Adobe CEF Helper", "Adobe Desktop Service"
    "AdobeIPCBroker", "AdobeUpdateService", "CCLibrary", "Creative Cloud", "Creative Cloud Helper", "CCXProcess"
)
$autoDeskProcesses = @("acad","Revit","3dsmax","Inventor","Navisworks","AdskLicensingService","AdskAccessServiceHost","AdskLicensingAgent","AdSSO",
    "AdskIdentityManager","AdODIS-installer","AutodeskDesktopApp","AdAppMgrSvc","WorksharingMonitor","FNPLicensingService64"
)

switch ($targetSoftware) {


#region  --- Browsers  ---

"Edge" {
    $installTimeout = 15
    $listPath      = "$listPathRoot\Browsers\Microsoft_Edge.txt"
    $software      = "Microsoft Edge"
    $processName   = @("msedge","MicrosoftEdgeUpdate","MicrosoftEdgeCrashpad","MicrosoftEdgeElevatedService","setup")
    $compliantVer  = "146.0.3856.84"
    $patchPath     = "$patchRoot\Browsers\Microsoft_Edge_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Microsoft\Edge\Application\msedge.exe"
    $installLine   = "& cmd /c 'C:\Windows\System32\msiexec.exe /i C:\Temp\Microsoft_Edge_$compliantVer\MicrosoftEdgeEnterpriseX64.msi /qn'"
}

"EdgeRemove" {
    $installTimeout = 15
    $listPath      = "$listPathRoot\Browsers\Microsoft_Edge.txt"
    $software      = "Microsoft Edge"
    $processName   = @("msedge","MicrosoftEdgeUpdate","MicrosoftEdgeCrashpad","MicrosoftEdgeElevatedService","setup")
    $compliantVer  = "146.0.3856.84"
    $patchPath     = "$patchRoot\Browsers\Microsoft_Edge_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Edge.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Microsoft\Edge\Application\msedge.exe"
    $installLine   = $null
}

"Chrome" {
    $installTimeout = 15
    $listPath      = "$listPathRoot\Browsers\Google_Chrome.txt"
    $software      = "Google Chrome"
    $processName   = "chrome","GoogleUpdate","GoogleCrashHandler","GoogleCrashHandler64"
    $compliantVer  = "146.0.7680.178"
    $patchPath     = "$patchRoot\Browsers\Google_Chrome_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $installLine   = "& cmd /c 'C:\Windows\System32\msiexec.exe /i C:\Temp\Google_Chrome_$compliantVer\googlechromestandaloneenterprise64.msi /qn'"
                     # "PowerShell.exe -ExecutionPolicy Bypass -File 'C:\Temp\Google_Chrome_$compliantVer\installer.ps1'"
    $softwarePaths = "C:\Program Files*\Google\Chrome\Application\Chrome.exe"
}

"ChromeRemove" {
    $installTimeout = 15
    $listPath      = "$listPathRoot\Browsers\Google_Chrome.txt"
    $software      = "Google Chrome"
    $processName   = "chrome","GoogleUpdate","GoogleCrashHandler","GoogleCrashHandler64"
    $compliantVer  = "146.0.7680.178"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Chrome.ps1").ScriptBlock
    $installLine   = $null
    $softwarePaths = @("C:\Program Files*\Google\Chrome\Application\Chrome.exe"
                     "C:\Users\USER\AppData\Local\Google\Chrome\Application\Chrome.exe")
}

"Firefox" {
    $installTimeout = 15
    $listPath      = "$listPathRoot\Browsers\Mozilla_Firefox.txt"
    $software      = "Firefox"
    $processName   = "firefox"
    $compliantVer  = "140.9.0"
    $patchPath     = "$patchRoot\Browsers\Mozilla_Firefox_$compliantVer"
    $softwarePaths = "C:\Program Files*\Mozilla Firefox\Firefox.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $installLine   = @(
        "Set-Location C:\Temp\Mozilla_Firefox_$compliantVer;",
        "PowerShell.exe -ExecutionPolicy Bypass -File 'C:\Temp\Mozilla_Firefox_$compliantVer\install.ps1'"
    )
}

"FirefoxAppData" {
    $installTimeout = 15
    $listPath      = "$listPathRoot\Browsers\Mozilla_Firefox.txt"
    $software      = "Firefox"
    $processName   = "firefox"
    $compliantVer  = "999.9.9"
    $patchPath     = "$patchRoot\Browsers\Uninstall"
    $softwarePaths = "C:\Users\USER\AppData\Local\Mozilla Firefox\Firefox.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $firePath      = '"C:\Users\USER\AppData\Local\Mozilla Firefox\uninstall\helper.exe"'
    $installLine   = "(Start-Process $firePath -ArgumentList '/S' -Wait -PassThru).ExitCode"
}

"BraveRemove" {
    $installTimeout = 15
    $listPath      = "$listPathRoot\Browsers\Brave_Browser.txt"
    $software      = "Brave"
    $processName   = "brave","BraveUpdate","BraveCrashHandler","BraveCrashHandler64"
    $compliantVer  = "999.9.9"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Brave.ps1").ScriptBlock
    $installLine   = $null
    $softwarePaths = "C:\Users\USER\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe"
}

#endregion  --- Browsers ---


#region  ---  Adobe Software  ---

"Reader" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Reader.txt"
    $software      = "Adobe Acrobat Reader"
    $processName   = $adobeCCProcesses
    $compliantVer  = "26.001.21346"
    $patchPath     = "$patchRoot\Adobe\Reader_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Acrobat Reader*\Reader\AcroRd32.exe"
    $installLine   = "& cmd /c 'C:\Temp\Reader_$compliantVer\setup.exe'"
}

# Removes Adobe Acrobat Reader DC 2021
"ReaderRemove" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Reader.txt"
    $software      = "Adobe Acrobat Reader DC"
    $processName   = $adobeCCProcesses
    $compliantVer  = "26.001.21346"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Reader.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Acrobat Reader*\Reader\AcroRd32.exe"
    $installLine   = $null
}

# Ready - Installs Adobe Acrobat. Patching post install is required to install up-to-date version
"AcrobatInstall" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Acrobat.txt"
    $software      = "Adobe Acrobat DC"
    $processName   = @($adobeCCProcesses) + @("Acrobat")
    $compliantVer  = "23.003.20210"
    $patchPath     = "$patchRoot\Adobe\Acrobat_Install_Package"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    $installLine   = @(
        "Set-Location 'C:\Temp\Acrobat_Install_Package';"
        "Set-ExecutionPolicy Bypass -Scope Process -Force;"
        "PowerShell.exe -ExecutionPolicy Bypass -File 'C:\Temp\Acrobat_Install_Package\setup.exe'")
}

# Updates Adobe Acrobat Pro
"AcrobatUpdate" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Acrobat.txt"
    $software      = "Adobe Acrobat DC"
    $processName   = @($adobeCCProcesses) + @("Acrobat")
    $compliantVer  = "26.001.21346"
    $patchPath     = "$patchRoot\Adobe\AcrobatDC_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    $installLine   = "& cmd /c 'C:\Temp\AcrobatDC_$compliantVer\AcrobatDCUpd$($($compliantVer).replace('.','')).msp'"
}

# Removes Adobe Acrobat Online ???
"AcrobatRemove" {
    $installTimeout = 20
    $listPath      = "$listPathRoot\Adobe\Acrobat.txt"
    $software      = "Adobe Acrobat" # Updated for Adobe Acrobat Online   # "Adobe Acrobat DC"
    $processName   = @($adobeCCProcesses) + @("Acrobat")
    $compliantVer  = "99.999" # "22.001.201177"
    $patchPath     = "$patchRoot\Adobe\Acrobat_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Acrobat-Online.ps1").ScriptBlock
    $installLine   = $null
    $softwarePaths = "C:\Program Files*\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
}

"AdobeCC" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\AdobeCC.txt"
    $software      = "Adobe Creative Cloud"
    $processName   = $adobeCCProcesses
    $compliantVer  = "5.10.0"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $installLine   = "(Start-Process 'C:\Program Files (x86)\Adobe\Adobe Creative Cloud\Utils\Creative Cloud Uninstaller.exe' -ArgumentList '-uninstall' -PassThru -Wait).ExitCode ; Start-Sleep 90"
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Creative Cloud\ACC\Creative Cloud.exe"
}

# Only uninstalls Adobe After Effects
"AfterEffects" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\AfterEffects.txt"
    $software      = "Adobe After Effects"
    $processName   = @($adobeCCProcesses) + @("AfterFX")
    $compliantVer  = "999.9.9"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-AfterEffects.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe After Effects*\Support Files\AfterFX.exe"
    $installLine   = $null
}

# Only uninstalls Adobe Animate
"Animate" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Animate.txt"
    $software      = "Animate"
    $processName   = @($adobeCCProcesses) + @("Animate")
    $compliantVer  = "999.9.9"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Animate.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Animate*\Animate.exe"
    $installLine   = $null
}

# Only uninstalls Adobe Audition
"Audition" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Audition.txt"
    $software      = "Adobe Audition"
    $processName   = @($adobeCCProcesses) + @("Adobe Audition")
    $compliantVer  = "999.9.9"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Audition.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Audition*\Adobe Audition.exe"
    $installLine   = $null
}

# Only uninstalls Adobe Bridge
"Bridge" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Bridge.txt"
    $software      = "Bridge"
    $processName   = @($adobeCCProcesses) + @("Bridge")
    $compliantVer  = "999.9.9"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Bridge.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Bridge*\Bridge.exe"
    $installLine   = $null
}

# Ready - Installs version 16.0.5.1096
"Frame" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\FrameMaker.txt"
    $software      = "Adobe FrameMaker 2020"
    $processName   = @($adobeCCProcesses) + @("FrameMaker")
    $compliantVer  = "16.0.5.1096" # 16.0.5.1096
    $patchPath     = "$patchRoot\Adobe\FrameMaker_16.0.5.1096"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe FrameMaker*\FrameMaker.exe"
    $installLine   = "Set-Location 'C:\Temp\FrameMaker_$compliantVer' ; PowerShell.exe -ExecutionPolicy Bypass -File 'C:\Temp\FrameMaker_$compliantVer\install.ps1'"
}

# Ready 2/6/2024 - Version 16.0.5.1096
"FrameUninstall" {                          # NEED TO UPDATE THIS TO UNINSTALL EVERY VERSION
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\FrameMaker.txt"
    $software      = "Adobe FrameMaker 2020"
    $processName   = @($adobeCCProcesses) + @("FrameMaker")
    $compliantVer  = "16.0.5.1097" # 16.0.5.1096
    $patchPath     = "$patchRoot\Adobe\FrameMaker_16.0.5.1096"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe FrameMaker*\FrameMaker.exe"
    $installLine   = "Set-ExecutionPolicy Bypass -Scope Process -Force ; Set-Location 'C:\Temp\FrameMaker_16.0.5.1096' ; .\FramMU5_uninstall.ps1"
}

# Ready 2/6/2024 - Version 16.0.5.1096
"FrameRemove" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\FrameMaker.txt"
    $software      = "Adobe FrameMaker 2020"
    $processName   = @($adobeCCProcesses) + @("FrameMaker")
    $compliantVer  = "16.0.5.1097" # 16.0.5.1096
    $patchPath     = "$patchRoot\Adobe\Uninstall_Packages\FrameMaker"
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-FrameMaker.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe FrameMaker*\FrameMaker.exe"
    $installLine   = $null
}

"Illustrator" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Illustrator.txt"
    $software      = "Illustrator"
    $processName   = @($adobeCCProcesses) + @("Illustrator")
    $compliantVer  = "30.1"
    $patchPath     = "$patchRoot\Adobe\Illustrator_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Illustrator 202*\Support Files\Contents\Windows\Illustrator.exe"
    $installLine   = "Set-ExecutionPolicy Bypass -Scope Process -Force ; Set-Location 'C:\Temp\Illustrator_$compliantVer' ; .\install_Illustrator_29_7_1.ps1 ; Set-ExecutionPolicy Restricted -Scope Process -Force"
}

"IllustratorUninstall" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Illustrator.txt"
    $software      = "Illustrator"
    $processName   = @($adobeCCProcesses) + @("Illustrator")
    $compliantVer  = "30.1" # 29.8 required, not in SC yet
    $patchPath     = "$patchRoot\Adobe\Illustrator_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Illustrator 202*\Support Files\Contents\Windows\Illustrator.exe"
    $installLine   = "Set-ExecutionPolicy Bypass -Scope Process -Force ; Set-Location 'C:\Temp\Illustrator_$compliantVer' ; .\uninstall_Illustrator_29_7_1.ps1 ; Set-ExecutionPolicy Restricted -Scope Process -Force"
}

# Only uninstalls Adobe Lightroom Classic
"Lightroom" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Lightroom.txt"
    $software      = "Adobe Lightroom"
    $processName   = @($adobeCCProcesses) + @("Adobe Lightroom")
    $compliantVer  = "999.9.9"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptroot\Patching\AppUninstalls\Uninstall-Lightroom.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Lightroom*\Lightroom.exe"
    $installLine   = $null
}

# Only uninstalls Adobe LiveCycle Designer (AKA Forms Designer / Designer ES4)
"LiveCycle" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\LiveCycle.txt"
    $software      = "LiveCycle"
    $processName   = @($adobeCCProcesses) + @("LiveCycle", "Forms", "FormDesigner")
    $compliantVer  = "999.9.9"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-LiveCycle.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe LiveCycle*\FormDesigner.exe"
    $installLine   = $null
}

# Only uninstalls Adobe Media Encoder
"MediaEncoder" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\MediaEncoder.txt"
    $software      = "Adobe Media Encoder"
    $processName   = @($adobeCCProcesses) + @("Adobe Media Encoder")
    $compliantVer  = "999.9.9"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-MediaEncoder.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Media Encoder*\Adobe Media Encoder.exe"
    $installLine   = $null
}

"Photoshop" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Photoshop.txt"
    $software      = "Photoshop"
    $processName   = @($adobeCCProcesses) + @("Photoshop")
    $compliantVer  = "26.10"
    $patchPath     = "$patchRoot\Adobe\Photoshop_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Photoshop 202*\Photoshop.exe"
    $installLine   = "Set-ExecutionPolicy Bypass -Scope Process -Force ; Set-Location 'C:\Temp\Photoshop_$compliantVer' ; .\Install_Photoshop26_10.ps1 ; Set-ExecutionPolicy Restricted -Scope Process -Force"
}

"PhotoshopUninstall" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Photoshop.txt"
    $software      = "Photoshop"
    $processName   = @($adobeCCProcesses) + @("Photoshop")
    $compliantVer  = "26.10"
    $patchPath     = "$patchRoot\Adobe\Photoshop_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Photoshop 202*\Photoshop.exe"
    $installLine   = "Set-ExecutionPolicy Bypass -Scope Process -Force ; Set-Location 'C:\Temp\Photoshop_$compliantVer' ; .\Uninstall_Photoshop26_10.ps1 ; Set-ExecutionPolicy Restricted -Scope Process -Force"
}

# Only uninstalls Adobe Prelude
"Prelude" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Prelude.txt"
    $software      = "Adobe Prelude"
    $processName   = @($adobeCCProcesses) + @("Adobe Prelude")
    $compliantVer  = "999.9.9"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Prelude.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Prelude*\Adobe Prelude.exe"
    $installLine   = $null
}

# Only uninstalls Adobe Premiere Pro
"Premiere" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Adobe\Premiere.txt"
    $software      = "Adobe Premiere Pro"
    $processName   = @($adobeCCProcesses) + @("Adobe Premiere Pro")
    $compliantVer  = "999.9.9"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Premiere.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Adobe\Adobe Premiere Pro*\Adobe Premiere Pro.exe"
    $installLine   = $null
}


#endregion  ---  Adobe Software  ---


#region  ---  User Specific Software  ---

"Zoom" {
    $installTimeout = 15
    $listPath      = "$listPathRoot\Zoom.txt"
    $software      = "Zoom"
    $processName   = "Zoom"
    $compliantVer  = "999.9.9.9" # 6.2.5
    $patchPath     = "$patchRoot\Zoom"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Users\USER\AppData\Roaming\Zoom\bin\Zoom.exe"
    $installLine   = '& cmd /c "C:\Temp\Zoom\CleanZoom.exe /silent"'
}

"Teams" {
    $installTimeout = 15
    $listPath      = "$listPathRoot\Microsoft_Teams.txt"
    $software      = "Microsoft Teams"
    $processName   = "Teams"
    $compliantVer  = "1.6.0.26474" # 25163.3611.3774.6315
    $patchPath     = "$patchRoot\Microsoft_Teams\Teams_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $installLine   = ". C:\Temp\Teams_$compliantVer\installTeams-Oct2023.ps1" # . seems to work"
    $softwarePaths = "C:\Users\USER\AppData\Local\Microsoft\Teams\current\Teams.exe"
}

"TeamsNEW" { # WIP
    $installTimeout = 15
    $listPath      = "$listPathRoot\Microsoft_Teams.txt"
    $software      = "Microsoft Teams"
    $processName   = "Teams"
    $compliantVer  = "25163.3611.3774.6315"
    $patchPath     = "$patchRoot\Microsoft_Teams\Teams_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $installLine   = ". C:\Temp\Teams_$compliantVer\installTeams-Oct2023.ps1" # . seems to work"
    $softwarePaths = "C:\Program Files*\WindowsApps\MSTeams_*\ms-teams.exe"
}

# WIP
"TeamsAppData" {
    $installTimeout = 15
    $listPath      = "$listPathRoot\Microsoft_Teams.txt"
    $software      = "Microsoft Teams" # WIP
    $processName   = "Teams"
    $compliantVer  = "1.6.0.26474" # 1.10.54.0
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Teams.ps1").ScriptBlock
    $softwarePaths = "C:\Users\USER\AppData\Local\Microsoft\Teams\current\Teams.exe"
    $installLine   = '& cmd /c "C:\Users\USER\AppData\Local\Microsoft\Teams\Update.exe" --uninstall -s'

    # User uninstalls:
    # "cmd /c 'C:\Users\USER\AppData\Local\Microsoft\Teams\Update.exe' --uninstall -s"
    # (Start-Process 'C:\Users\USER\AppData\Local\Microsoft\Teams\Update.exe' -ArgumentList "--uninstall -s" -Wait -PassThru).ExitCode

    # System uninstalls:   ...for some reason
    # "& 'C:\Users\USER\AppData\Local\Microsoft\Teams\current\Teams.exe' --uninstall -s"
    # 'Start-Process "C:\Users\USER\AppData\Local\Microsoft\Teams\current\Teams.exe" -ArgumentList "--uninstall -s"'
}

"Grammarly" { # works if you remote in and run manually... sometimes. Otherwise use Uninstall-UserBasedSoftware.ps1
    $installTimeout = 15
    $listPath      = "$listPathRoot\Grammarly.txt"
    $software      = "Grammarly"
    $processName   = "Grammarly*"
    $compliantVer  = "9999.0.0.0"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Users\USER\AppData\Local\Grammarly\DesktopIntegrations\Grammarly*.exe"
    $installLine   = '& cmd /c "C:\Users\USER\AppData\Local\Grammarly\DesktopIntegrations\Uninstall.exe" /S'
}

#endregion ---  User Specific Software  ---


#region  ---  Misc. Software  ---

"AppGate" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\AppGate.txt"
    $software      = "AppGate"
    $processName   = "PROCESS"
    $compliantVer  = "6.2.7.0"
    $patchPath     = "$patchRoot\AppGate"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Appgate SDP\Appgate SDP.exe"
    $installLine   = @(
        "Set-ExecutionPolicy Bypass -Scope Process -Force;"
        "Set-Location 'C:\Temp\AppGate';"
        ".\AppGate627wAO.ps1")
}

# Ready
"Tanium" {
    $installTimeout = 30
    $tag           = "RegVersion"
    $listPath      = "$listPathRoot\Tanium_Client.txt"
    $software      = "Tanium Client 7"                                                               # Matches Tanium Client 7.X and avoids matching Tanium Client Installer
    $processName   = $null                                                                           # "TaniumClient", "TaniumCX", "TaniumDetectEngine", "TaniumDriverSvc", "TaniumTSDB"
    $compliantVer  =  "7.7.3.8198"
    $patchPath     = "$patchRoot\Tanium\Tanium_$compliantVer"
    $patchScript   =  (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock                            # (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock  # Default for old versions
    $softwarePaths = "Placeholder"                                                                   # "C:\Program Files (x86)\Tanium\Tanium Client\..." Don't even have read permission via the server :/
    $installLine   = @(
        "Copy-Item 'C:\Temp\Tanium_$compliantVer\*' -Destination 'C:\Program Files (x86)\Tanium\Tanium Client\' -Recurse;"
        "Set-Location 'C:\Temp\Tanium_$compliantVer';"
        ".\Install.bat"
    )
}

"Java" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Oracle_Java.txt"
    $software      = "Java 8"
    $processName   = "Java"
    $compliantVer  = "8.0.4810.25"
    $patchPath     = "$patchRoot\Oracle\Java_$compliantVer"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Java\*\bin\Java.exe"
    $installLine   = @(
        "Set-ExecutionPolicy Bypass -Scope Process -Force;",
        "Set-Location 'C:\Temp\Java_$compliantVer';",
        "PowerShell.exe -ExecutionPolicy Bypass -File  'C:\Temp\Java_$compliantVer\install_Oracle_JRE.ps1'")
        # (Start-Process "Powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File 'C:\Temp\Java_$compliantVer\install_Oracle_JRE.ps1'" -Wait -PassThru).ExitCode
}

"Lang" { #Restores Language Mode to 'Unrestricted" using the curl unblocker v2 script
    $installTimeout = 30
    $listPath      = "$listPathRoot\Language_Mode.txt"
    $software      = "Language Mode"
    $processName   = "PLACEHOLDER"
    $compliantVer  =  "9.9.9.9"
    $patchPath     = "$patchRoot\Curl"
    $patchScript   =  (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "Placeholder"
    $installLine   = @(
        "Set-ExecutionPolicy Bypass -Scope Process -Force;",
        "Set-Location 'C:\Temp\Curl';",
        ".\removeblockv2.ps1"
    )
}

# Updated 3/1/2022
"SQL" {
    $installTimeout = 30
    $software      = "SQL Server 2017" # WIP
    $listPath      = "$listPathRoot\Microsoft_SQL.txt"
    $compliantVer  = "2017.140.3445.2"
    $KB            = "KB5014553"
    $patchPath     = "$patchRoot\SQL\$KB"
    $patchName     = "SQLServer2017-$KB-x64.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\Microsoft SQL Server\MSSQL14.SPEED_DB\MSSQL\Binn\sqlservr.exe"
    $installLine   = "& cmd /c 'C:\Temp\$KB\$patchName /Q /IACCEPTSQLSERVERLICENSETERMS'"
}

"SQL16" {
    $installTimeout = 30
    $software      = "SQL Server 2016" # WIP
    $listPath      = "$listPathRoot\Microsoft_SQL.txt"
    $compliantVer  = "2017.140.3445.2"
    $KB            = "KB5014553"
    $patchPath     = "$patchRoot\SQL\$KB"
    $patchName     = "SQLServer2017-$KB-x64.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Windows\System32\msoledbsql.dll"
    $installLine   = "& cmd /c 'C:\Temp\$KB\$patchName /Q /IACCEPTSQLSERVERLICENSETERMS'"
}

"SQLdll" {
    $installTimeout = 30
    $software      = "SQL Server 2017" # WIP
    $listPath      = "$listPathRoot\Microsoft_SQL.txt"
    $compliantVer  = "17.10.3.1"
    $versionType   = "Product"
    $KB            = "KB5021127"
    $patchPath     = "$patchRoot\SQL\$KB"
    $patchName     = "SQLServer2017-$KB-x64.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Windows\System32\msodbcsql17.dll"
    $installLine   = "& cmd /c 'C:\Temp\$KB\$patchName /Q /IACCEPTSQLSERVERLICENSETERMS'"
}

"SQLmsi" {
    $installTimeout = 30
    $software      = "SQL Server 2017" # WIP
    $listPath      = "$listPathRoot\Microsoft_SQL.txt"
    $compliantVer  = "2017.1710.6" # 17.10.6 in ACAS
    $versionType   = "Product"
    $patchPath     = "$patchRoot\SQL\ODBC_Driver_SQL_64"
    $patchName     = "msodbcsql.msi"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Windows\System32\msodbcsql17.dll"
    $installLine   = "& cmd /c 'C:\Windows\System32\msiexec.exe /i C:\Temp\ODBC_Driver_SQL_64\msodbcsql.msi IACCEPTMSODBCSQLLICENSETERMS=YES /qn'"
    # "& cmd /c 'C:\Temp\ODBC_Driver_SQL_64\msodbcsqlx64.msi /Q /IACCEPTSQLSERVERLICENSETERMS'"
    # "& cmd /c "SQLServer2017-KB4052987-x64.exe /update /quiet /norestart"
    # MsiExec.exe /i msodbcsql.msi IACCEPTMSODBCSQLLICENSETERMS=YES /qn
}

"SQLole" {
    $installTimeout = 30
    $software      = "SQL Server 2017" # WIP
    $listPath      = "$listPathRoot\Microsoft_SQL.txt"
    $compliantVer  = "2018.187.2.0" # 18.7.2 in ACAS
    $versionType   = "Product"
    $patchPath     = "$patchRoot\SQL\OLE_DB_Driver_SQL_64"
    $patchName     = "msoledbsql.msi"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Windows\System32\msoledbsql.dll"
    $installLine   = "& cmd /c 'C:\Windows\System32\msiexec.exe /i C:\Temp\OLE_DB_Driver_SQL_64\msoledbsql.msi IACCEPTMSOLEDBSQLLICENSETERMS=YES /qn'"
}

"SQL19" {
    $installTimeout = 30
    $software      = "SQL Server 2019" # WIP
    $listPath      = "$listPathRoot\Microsoft_SQL.txt"
    $compliantVer  = "2019.150.4390.2"
    $KB            = "KB5042749"
    $patchPath     = "$patchRoot\SQL\$KB"
    $patchName     = "sqlserver2019-$KB-x64.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\Microsoft SQL Server\MSSQL15.TPSDB\MSSQL\Binn\sqlservr.exe"
    $installLine   = "& cmd /c 'C:\Temp\$KB\$patchName /Q /IACCEPTSQLSERVERLICENSETERMS'"
}

# WIP - Uninstalls VMWare Horizon Client
"VMware" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\VMware_Horizon.txt"
    $software      = "VMware Horizon Client"
    $processName   = "PROCESS"
    $compliantVer  = "5.5.3"
    $patchPath     = "$patchRoot\Placeholder"
    $softwarePaths = "C:\Program Files*\VMware\VMware Horizon View Client\vmware-view.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $installLine   = '(Start-Process "C:\ProgramData\Package Cache\{b202e5f3-ac14-4e31-848c-78b7cae95e33}\VMware-Horizon-Client-5.5.0-16975072.exe" -ArgumentList "/uninstall" -Wait -PassThru).ExitCode'
}

# Works if you remote in and run manually >:(
"Pulse" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Pulse.txt"
    $software      = "Pulse Secure Desktop"
    $processName   = "Pulse"
    $compliantVer  = "9.1.9.4983" #?
    $patchPath     = "$patchRoot\Pulse91R9"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Common Files\Pulse Secure\JamUI\Pulse.exe"
    $installLine   = "powershell.exe -ExecutionPolicy Bypass -File 'C:\Temp\$(Split-Path $patchPath -Leaf)\install.ps1'"
}

"DellDriver" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Dell.txt"
    $software      = "Dell dbutil Driver"
    $processName   = "Process"
    $compliantVer  = "9.9.9.9" # No Version
    $patchPath     = "$patchRoot\Dell_cbutil_Driver"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Users\Administrator\AppData\Local\Temp" # No real path
    $installLine   = "powershell.exe -ExecutionPolicy Bypass -File 'C:\Temp\$(Split-Path $patchPath -Leaf)\install.ps1'"
}

"DellControlVault" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Dell_ControlVault.txt"
    $software      = "Dell ControlVault3"
    $processName   = "Process"
    $compliantVer  = "5.15.10.14"
    $patchPath     = "$patchRoot\Dell_ControlVault3"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\system32\Drivers\cvusbdrv.sys"
    $installLine   = @(
        "Set-Location 'C:\Temp\Dell_ControlVault3\';"
        "& cmd /c '`"C:\Temp\Dell_ControlVault3\Dell-ControlVault3_5.15.10.14.exe`" /s'")

    # Works locally!!:
    # ----------------
    # cd C:\temp\Dell_ControlVault3
    # .\Dell-ControlVault3_5.15.10.14.exe /s

    # Can't get to work remotely!!:
    # -----------------------------

    # '& "C:\Temp\To\Dell-ControlVault3_5.15.10.14.exe" /s' <--- next thing to try

    # @('.\Dell-ControlVault3_5.15.10.14.exe /s')
    #
    #


    # @(
    #     "Set-Location 'C:\Temp\Dell_ControlVault3'";
    #     ".\Dell_ControlVault3_5.15.10.14.exe /s"
    # )
    # "(Start-Process 'C:\Temp\Dell_ControlVault3\Dell-ControlVault3_$compliantVer.exe' -ArgumentList '/s' -PassThru -Wait).ExitCode"
    # "& cmd /c 'C:\Temp\Dell_ControlVault3\Dell-ControlVault3_$compliantVer.exe /s'"
    # "& cmd /c 'C:\Temp\Dell_ControlVault3\Dell-ControlVault3_$compliantVer.exe' /s"
    # ('& cmd /c "C:\Temp\Dell_ControlVault3\Dell-ControlVault3_') + ($compliantVer) +('.exe" /s')
}

# Uninstalls Trellix aka McAfee # WIP
"Trellix" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Trellix.txt"
    $software      = "Trellix"
    $compliantVer  = "999.0.0.0"
    $patchPath     = $null
    $patchName     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\McAfee\Agent\cmdagent.exe"
    $installLine   = 'Start-procesS "C:\Program Files\McAfee\Agent\x86\frminst.exe" -ArgumentList "/forceuninstall /SILENT" -wait'
}

"McAfee" {
    $installTimeout = 30
    $tag           = "PsExec"
    $listPath      = "$listPathRoot\McAfee_Agent.txt"
    $software      = "McAfee Agent"
    $compliantVer  = "5.8.3.622"
    $patchPath     = "$patchRoot\McAfee_Agents"
    $patchName     = "MAgent_$compliantVer.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\McAfee\Agent\cmdagent.exe"
}

# Uninstalls McAfee Agent
"McAfeeMA" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\McAfee_MA.txt"
    $software      = "McAfee MA"
    $processName   = "Process"
    $compliantVer  = "5.7.8.262"
    $patchPath     = "$patchRoot\Trellix_EPR"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\McAfee\Agent\cmdagent.exe"
    $installLine   = '& cmd /c "C:\Temp\Trellix_EPR\EndpointProductRemoval_23.2.0.64.exe --accepteula --MA --T=30"'
    # '(Start-Process "C:\Temp\McAfee_EPR\McAfeeEndpointProductRemoval_22.5.0.54.exe" -ArgumentList "--accepteula --MA --T30" -Wait -PassThru).ExitCode'
}

# Uninstalls McAfee Policy Auditor
"McAfeePA" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\McAfee_PA.txt"
    $software      = "McAfee PA"
    $processName   = "Process"
    $compliantVer  = "6.5.2"
    $patchPath     = "$patchRoot\Trellix_EPR"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\McAfee\Policy Auditor Agent\AuditManagerService.exe"
    $installLine   = '& cmd /c "C:\Temp\Trellix_EPR\EndpointProductRemoval_23.2.0.64.exe --accepteula --PA --T=30"'
    # '(Start-Process "C:\Temp\McAfee_EPR\McAfeeEndpointProductRemoval_22.5.0.54.exe" -ArgumentList "--accepteula --PA --T30" -Wait -PassThru).ExitCode'
}

# Uninstalls McAfee Endpoint Security
"McAfeeENS" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\McAfee_ENS.txt"
    $software      = "McAfee ENS"
    $processName   = "Process"
    $compliantVer  = "10.7.0.3255"
    $patchPath     = "$patchRoot\Trellix_EPR"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\McAfee\Endpoint Security\Endpoint Security Platform\mfeesp.exe"
    $installLine   = '& cmd /c "C:\Temp\Trellix_EPR\EndpointProductRemoval_23.2.0.64.exe --accepteula --ENS --T=30"'
    # '(Start-Process "C:\Temp\McAfee_EPR\McAfeeEndpointProductRemoval_22.5.0.54.exe" -ArgumentList "--accepteula --ENS --T30" -Wait -PassThru).ExitCode'
}

# Uninstalls McAfee Data Loss Prevention
"McAfeeDLP" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\McAfee_DLP.txt"
    $software      = "McAfee DLP"
    $processName   = "Process"
    $compliantVer  = "11.6.600.212"
    $patchPath     = "$patchRoot\Trellix_EPR"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\McAfee\DLP\Agent\fcagd.exe"
    $installLine   = '& cmd /c "C:\Temp\Trellix_EPR\EndpointProductRemoval_23.2.0.64.exe --accepteula --DLP --T=30"'
    # '(Start-Process "C:\Temp\McAfee_EPR\McAfeeEndpointProductRemoval_22.5.0.54.exe" -ArgumentList "--accepteula --DLP --T30" -Wait -PassThru).ExitCode'
}

# Only uninstalls
"VSCode" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\VS_Code.txt"
    $software      = "Microsoft VS Code"
    $processName   = "Code"
    $compliantVer  = "1.105.1"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = @("C:\Program Files*\Microsoft VS Code\Code.exe"
                       "C:\Users\USER\AppData\Local\Programs\Microsoft VS Code\Code.exe"
    )
    $installLine   = @('& "C:\Program Files\Microsoft VS Code\unins000.exe" /SILENT',  # Doesn't wait for uninstall to finish, does work though
                       '& "C:\users\USER\Microsoft VS Code\unins000.exe" /SILENT'      # Works only locally
    )
}

# POAMed so DO NOT USE
"Sybase" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Sybase.txt"
    $software      = "Sybase"
    $processName   = "PROCESS"
    $compliantVer  = "2.16.0"
    $patchPath     = "$patchRoot\PlaceholderPath" # No patch
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Sybase.ps1").ScriptBlock
    $softwarePaths = @("C:\Program Files*\Sybase\SCC-3_3\common\lib\log4j-1.2.*.jar"
                       "C:\Program Files*\Sybase\SybCentral\lib\log4j-1.2.*.jar")
    $installLine   = $null # "Placeholder Line"
}

# Only uninstalls 7.1.6.1
"ICODES" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\ICODES.txt"
    $software      = "ICODES"
    $processName   = "PROCESS"
    $compliantVer  = "999.9.9.9"
    $patchPath     = "$patchRoot\PlaceholderPath" # No patch.
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = @("C:\ProgramData\ICODES\uninstall_5686_ICODES_Desktop_7.1.6.1.exe")
    $installLine   = '& "C:\ProgramData\ICODES\uninstall_5686_ICODES_Desktop_7.1.6.1.exe" /S'
}

# WIP
"Cognos" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\IBM_Cognos.txt"
    $software      = "IBM Cognos"
    $processName   = "PROCESS"
    $compliantVer  = "99.99.99.99"
    $patchPath     = "$patchRoot\PlaceholderPath" # No patch
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = @("C:\Program Files\ibm\cognos\tm1_64\coginsight\plugins\com.ibm.cognos.fmeng.libraries_10.2.6102.19\lib\log4j-1.2.17.jar")
    $installLine   = 'Placeholder Line'
}

# Uninstalls only - Need to change get-version module to allow for product version priority
"ForeScout" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\ForeScout.txt"
    $software      = "ForeScout"
    $processName   = "PROCESS"
    $compliantVer  = "8.2.3.0" # 8.1.4
    $patchPath     = "$patchRoot\PlaceholderPath" # No patch
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = @("C:\Program Files*\ForeScout\GuiManager\current\Forescout Console.exe"
                       "C:\Users\USER\ForeScout\GuiManager*\current\Forescout Console.exe")
    $installLine   = @("& cmd /c '`"C:\Program Files (x86)\ForeScout\Uninstall\Uninstall.exe`" --mode unattended'"
                       "& cmd /c '`"C:\Users\USER\Forescout Console 8.1.3\Uninstall\Uninstall.exe`" --mode unattended'"
                       "& cmd /c '`"C:\Users\USER\Forescout Console 8.1.4\Uninstall\Uninstall.exe`" --mode unattended'"
                       "& cmd /c '`"C:\Users\USER\Forescout Console 8.2.2\Uninstall\Uninstall.exe`" --mode unattended'")
}

"IntelWireless" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Intel_Wireless.txt"
    $software      = "Intel Wireless"
    $processName   = "Process"
    $compliantVer  = "9.9.9.9" # No real version
    $patchPath     = "M:\Regional\VMT\Patches\Intel"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = @("C:\WINDOWS\INF\netwtw04.inf"  # No real path
                       "C:\WINDOWS\INF\netwtw06.inf"
                       "C:\WINDOWS\INF\netwtw08.inf"
                       "C:\WINDOWS\INF\netwtw6e.inf")
    $installLine   = @("pnputil.exe /add-driver C:\Temp\Intel\netwtw04.inf /install;"
                       "pnputil.exe /add-driver C:\Temp\Intel\netwtw06.inf /install;"
                       "pnputil.exe /add-driver C:\Temp\Intel\netwtw08.inf /install;"
                       "pnputil.exe /add-driver C:\Temp\Intel\netwtw6e.inf /install;")
}

# Untested
"IntelWirelessRemove" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Intel.txt"
    $software      = "Intel Wireless"
    $processName   = "Process"
    $compliantVer  = "9.9.9.9" # No real version
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-IntelWireless.ps1").ScriptBlock
    $softwarePaths = @("C:\WINDOWS\INF\netwtw04.inf"
                       "C:\WINDOWS\INF\netwtw06.inf"
                       "C:\WINDOWS\INF\netwtw08.inf"
                       "C:\WINDOWS\INF\netwtw6e.inf")
    $installLine   = $null
}

# Updated to 4.18.2302.7
"Defender" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Windows_Defender.txt"
    $software      = "Windows Defender"
    $processName   = "PROCESS"
    $compliantVer  = "4.18.2302.7"
    $patchPath     = "$patchRoot\Defender\KB4052623"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\Windows Defender\MpCmdRun.exe"
    $installLine   = @(
        "Set-Location 'C:\Temp\KB4052623\$compliantVer';"
        "& cmd /c '.\Installer.cmd'")
}

"Spark" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Cisco_Spark.txt"
    $software      = "Cisco Spark"
    $processName   = "CiscoCollabHost", "CiscoCollabHostCef", "ciscowebexstart", "CiscoWebexWebService", "atmgr", "WebexHost", "Cisco Webex Meetings", "Webex Teams", "Cisco Webex Meetings Desktop"
    $compliantVer  = "42.7"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-Spark.ps1").ScriptBlock
    $softwarePaths = "C:\Users\USER\AppData\Local\Programs\Cisco Spark\dependencies\CiscoCollabHostCef.exe"
    $installLine   = '(Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList "/X {8B6CF205-9F89-4715-8DEE-D6C5D02E3E98} /quiet /norestart" -Wait -PassThru).ExitCode'
}

"Jabber" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Cisco_Jabber.txt"
    $software      = "Cisco Jabber"
    $processName   = "PROCESS"
    $compliantVer  = "14.1.3"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Cisco Systems\Cisco Jabber\CiscoJabber.exe"
    $installLine   = $null #'(Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList "/X {8B6CF205-9F89-4715-8DEE-D6C5D02E3E98} /quiet /norestart" -Wait -PassThru).ExitCode'
}

"Webex" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Cisco_Webex.txt"
    $software      = "Webex"
    $processName   = "CiscoCollabHost", "CiscoCollabHostCef", "ciscowebexstart", "CiscoWebexWebService", "atmgr", "WebexHost", "Cisco Webex Meetings", "Webex Teams", "Cisco Webex Meetings Desktop"
    $compliantVer  = "99999.9.9.9"
    $patchPath     = "$patchRoot\Cisco_Webex_Tool"
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-WebEx.ps1").ScriptBlock
    $softwarePaths = @(
        "C:\Program Files*\WebEx\Webex.exe"
        "C:\Users\USER\AppData\Local\WebEx\WebexHost.exe")
    $installLine   = @(
        '(Start-Process "C:\Temp\Cisco_Webex_Tool\CiscoWebexRemoveTool.exe" -ArgumentList "/s" -PassThru -Wait).ExitCode ;'
        '(Start-Process "C:\Users\USER\AppData\Local\WebEx\atcliun.exe" -ArgumentList "/x MEETINGS LANGUAGE=EN" -Wait -PassThru).ExitCode')
}

"AutoCAD2024Uninstall" { # Works but takes very long time, will timeout
    $installTimeout = 90
    $listPath      = "$listPathRoot\AutoCAD.txt"
    $software      = "AutoCAD 2024"
    $processName   = $autoDeskProcesses
    $compliantVer  = "30.3.222.0"
    # $versionType   = "Product"
    $patchPath     = "$patchRoot\Autodesk\AutoCAD2026"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Autodesk\AutoCAD*\acad.exe"
    $installLine   = "Set-ExecutionPolicy Bypass -Force ; Set-Location 'C:\Temp\AutoCAD2026' ; .\uninstallAutocad2024.ps1 ; Set-ExecutionPolicy Restricted -Force"
}

"AutoCAD2026" {
    $installTimeout = 90
    $listPath      = "$listPathRoot\AutoCAD.txt"
    $software      = "AutoCAD 2026" # "AutoDesk"
    $processName   =  $autoDeskProcesses
    $compliantVer  = "31.1.122.0"
    # $versionType   = "Product"
    $patchPath     = "$patchRoot\Autodesk\AutoCAD2026"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Autodesk\AutoCAD*\acad.exe"
    $installLine   = "Set-ExecutionPolicy Bypass -Force ; Set-Location 'C:\Temp\AutoCAD2026' ; .\installAutocad2026.ps1 ; Set-ExecutionPolicy Restricted -Force"
}

"AutodeskRevit" { # WIP... Just use Invoke-SoftwareUninstall bth...
    $installTimeout = 30
    $listPath      = "$listPathRoot\Autodesk.txt"
    $software      = "Autodesk Revit 2024"
    $processName   = $autoDeskProcesses
    $compliantVer  = "24.3.3.0" # "24.2.10.64"
    # $versionType   = "Product"
    $patchPath     = $null # "$patchRoot\Autodesk\AutoCAD2024_Package" WIP
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Autodesk\Revit 2024\Revit.exe"
    $installLine   = '(Start-Process "C:\Program Files\Autodesk\AdODIS\V1\Installer.exe" -ArgumentList "-i uninstall --trigger_point system -m C:\ProgramData\Autodesk\ODIS\metadata\{F9013D08-6F9F-3F9B-8360-93C40ABE4C1B}\bundleManifest.xml -x C:\ProgramData\Autodesk\ODIS\metadata\{F9013D08-6F9F-3F9B-8360-93C40ABE4C1B}\SetupRes\manifest.xsd" -Wait -PassThru).ExitCode'
}

"AdODISupdate" { #WIP
    $installTimeout = 30
    $listPath      = "$listPathRoot\AdODIS.txt"
    $software      = "AutoDesk ODIS" # "AutoDesk"
    $processName   = $autoDeskProcesses
    $compliantVer  = "2.19.0.5"
    # $versionType   = "Product"
    $patchPath     = "$patchRoot\Autodesk\AdODIS"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Autodesk\AdODIS\V1\Installer.exe"
    $installLine   = @(
        "Set-ExecutionPolicy Bypass -Force ; Set-Location 'C:\Temp\AdODIS' ; .\installAdODIS.ps1 ; " +
        # "Set-Location 'C:\Program Files\Autodesk\AdODIS\V1' ; & cmd /c '.\RemoveODIS.exe --mode unattended' ; " +
        # "Set-Location 'C:\Program Files\Autodesk\AdODIS\V1\Acccess' ; & cmd /c '.\RemoveAccess.exe --mode unattended' ; " +
        "Set-ExecutionPolicy Restricted -Force"
    )
}

"AdODISuninstall" { #WIP
    $installTimeout = 30
    $listPath      = "$listPathRoot\AdODIS.txt"
    $software      = "AutoDesk ODIS" # "AutoDesk"
    $processName   = $autoDeskProcesses
    $compliantVer  = "2.19.0.5"
    # $versionType   = "Product"
    $patchPath     = "$patchRoot\Autodesk\AdODIS"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Autodesk\AdODIS\V1\Installer.exe"
    $installLine   = @(
        "Set-Location 'C:\Program Files\Autodesk\AdODIS\V1' ; & cmd /c '.\RemoveODIS.exe --mode unattended' ; " +
        "Set-Location 'C:\Program Files\Autodesk\AdODIS\V1\Access' ; & cmd /c '.\RemoveAccess.exe --mode unattended'"
    )
}

# WIP
"FireEye" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\FireEye.txt"
    $software      = "FireEye"
    $processName   = "FireEye"
    $compliantVer  = "999.99.99" # "23.1.1.6"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\AppUninstalls\Uninstall-FireEye.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files (x86)\FireEye\xagt\xagt.exe"
    $installLine   = $null
}

"DCI" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\DCI.txt"
    $software      = "DCI"
    $processName   = "DCI_Wrapper"
    $compliantVer  = "999.9.9.9" # "23.1.1.6"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files (x86)\DCI\DCI_GUI\DCIClient.exe"
    $installLine   = $null
}

"Citrix" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Citrix.txt"
    $software      = "Citrix"
    $processName   = "PROCESS"
    $compliantVer  = "19.12.7001" # "19.12.6000"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files (x86)\Citrix\Citrix Workspace 1912\InstallHelper.exe" # "C:\Program Files (x86)\Citrix\Citrix Workspace 1912\InstallerHelper.exe"
    $installLine   = '& "C:\Program Files (x86)\Citrix\Citrix Workspace 1912\TrolleyExpress.exe" /uninstall /cleanup /silent'
}

"Git" { # WIP - incomplete
    $installTimeout = 30
    $listPath      = "$listPathRoot\Git.txt"
    $software      = "Git" # also matches against other software too, so be careful
    $processName   = "PROCESS"
    $compliantVer  = "2.43.0"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\Git\git-bash.exe"
    $installLine   = $null
}

"VLC" { # WIP
    $installTimeout = 30
    $listPath      = "$listPathRoot\VLC.txt"
    $software      = "VLC"
    $processName   = "VLC"
    $compliantVer  = "3.0.23"
    $patchPath     = "$patchRoot\VLC_$compliantVer" # No patch
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\VideoLAN\VLC\vlc.exe"
    $installLine   = @(
        'Start-Process -FilePath "C:\Program Files (x86)\VideoLAN\VLC\uninstall.exe" -ArgumentList "/S" -Wait ; Sleep 10 ;'
        "Set-ExecutionPolicy Bypass -Force ; Set-Location 'C:\Temp\VLC_$compliantVer' ; .\Install_VLC3023.ps1 ; Set-ExecutionPolicy Restricted -Force"
    )
}

"VLCuninstall" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\VLC.txt"
    $software      = "VLC"
    $processName   = "VLC"
    $compliantVer  = "3.0.23.0"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files*\VideoLAN\VLC\vlc.exe"
    $installLine   = 'Start-Process -FilePath "C:\Program Files (x86)\VideoLAN\VLC\uninstall.exe" -ArgumentList "/S" -Wait ; Sleep 10'
}

"WinRAR" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\WinRAR.txt"
    $software      = "WinRAR"
    $processName   = "WinRAR"
    $compliantVer  = "7.11"
    $patchPath     = $null
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\WinRAR\WinRAR.exe"
    $installLine   = '& "C:\Program Files\WinRAR\uninstall.exe" /S ; Sleep 10'
}
"MDER" { #WIP
    $installTimeout = 30
    $listPath      = "$listPathRoot\MDER.txt"
    $software      = "MDER Sensor Package"
    $processName   = "PROCESS"
    $compliantVer  = "0.0.0.0" # "2.0.13" # Set to 0.0.0.0 to simply check if the file exists, since no version info
    $patchPath     = "$patchRoot\MDER_2.0.13\MDER-Sensor-Package-Setup-2.0.13 1.msi"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\MDER Sensor Package\tagScripts\mder.ps1"
    $installLine   = '(Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList "/i `"C:\Temp\MDER_2.0.13\MDER-Sensor-Package-Setup-2.0.13 1.msi`" /qn" -Wait -PassThru).ExitCode'
}

"RSAT" { # Works. Can only check if it's there, no version info on the .msc. For version info, check the registry at HKEY_LOCAL_MACHINE\SOFTWARE\MCEDS > RSATInstalled > 22H2
    $installTimeout = 30
    $listPath      = "$listPathRoot\RSAT.txt"
    $software      = "RSAT"
    $processName   = "PROCESS"
    $compliantVer  = "0.0.0.0"
    $patchPath     = "$patchRoot\RSAT"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "C:\Windows\System32\dsa.msc"
    $installLine   = "Set-ExecutionPolicy Bypass -Force ; Set-Location 'C:\Temp\RSAT' ; .\RSATinstall.ps1 ; Set-ExecutionPolicy Restricted -Force"
}


"Notepad" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Notepad.txt"
    $software      = "Notepad"
    $processName   = "Notepad"
    $compliantVer  = "11.2510.14.0"
    $patchPath     = "$patchRoot\OS Updates\Notepad"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $softwarePaths = "Placeholder" # No single path for Windows Notepad
    $installLine   = @(
        "Set-Location C:\Temp\Notepad;",
        "PowerShell.exe -ExecutionPolicy Bypass -File 'C:\Temp\Notepad\install.ps1'"
    )
}


#endregion  ---  Misc. Software  ---


#region  ---  Office Software (.msp) ---

# Microsoft 365 Apps / Office Click-to-Run
"C2R" {
    $installTimeout = 30
    #Jan 2026
    $compliantVer  = "16.0.19127.0"
    $KB            = "C2R"
    $software      = "C2R"
    $listPath      = "$listPathRoot\Office\$software.txt"
    $patchPath     = "$officeDir64\C2R.May"
    $patchName     = "OfficeSetup.exe"
    $softwarePaths = "$officeAppDir64\excel.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $installLine   = "& cmd /c 'c:\temp\C2R.May\OfficeSetup.exe'"
}
"word" {
    $installTimeout = 30
    #May 2024
    $compliantVer  = "16.0.17928.20512"
    $KB            = "C2R"
    $software      = "word"
    $listPath      = "$listPathRoot\Office\$software.txt"
    $patchPath     = "$officeDir64\C2R.Mar"
    $patchName     = "install.cmd"
    $softwarePaths = "$officeAppDir64\winword.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    #$installLine   = "cmd /c 'c:\temp\C2R.Mar\install.cmd'"
}
"outlookc2r" {
    $installTimeout = 30
    #Feb 2024 - C2R Template
    $compliantVer  = "16.0.16731.20550"
    $KB            = "C2R"
    $software      = "Outlook"
    $listPath      = "$listPathRoot\Office\$software.txt"
    $patchPath     = "$officeDir64\C2R.feb"
    $patchName     = "install.cmd"
    $softwarePaths = "$officeAppDir64\outlook.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $installLine   = '& cmd /c "c:\temp\C2R.feb\install.cmd"'
}
"VisioC2R" {
    $installTimeout = 30
    #feb2024
    $compliantVer  = "16.0.16731.20550"
    $KB            = "C2R"
    $software      = "Visio"
    $listPath      = "$listPathRoot\Office\$software.txt"
    $patchPath     = "$officeDir64\C2R.feb"
    $patchName     = "install.cmd"
    $softwarePaths = "$officeAppDir64\visio.exe"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-NoUninstall.ps1").ScriptBlock
    $installLine   = '& cmd /c "c:\temp\C2R.feb\install.cmd"'
}

#endregion  ---  Office Software  ---


#region  ---  Microsoft Update Packages (.msu)  ---
# Everything in this region requires the line  $tag = "PsExec"  to be in the switch

#region  ---  Windows Kernel  ---

"win1124h2" {
    $installTimeout = 60
    #feb 2025
    $tag           = "PsExec"
    $listPath      = "$listPathRoot\Microsoft_Krnl.txt"
    $software      = "Win 11 Feb 2026"
    $compliantVer  = "10.0.26100.7824" #oldversion
    $patchPath     = "$patchRoot\OS Updates\kb5077181"#kb5072033"<-December
    $patchName     = "kb5077181.msu"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\system32\ntoskrnl.exe"
}
"win11" {
    $installTimeout = 60
    #Feb 2025
    $tag           = "PsExec"
    $listPath      = "$listPathRoot\Microsoft_Krnl.txt"
    $software      = "Win 11 Feb 2026"
    $compliantVer  = "10.0.22621.6630"
    $patchPath     = "$patchRoot\OS Updates\kb5075941"
    $patchName     = "kb5075941.msu"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\system32\ntoskrnl.exe"
}
"win11B" {
    $installTimeout = 60
    #Jan 2025
    $tag           = "PsExec"
    $listPath      = "$listPathRoot\Microsoft_Krnl.txt"
    $software      = "Win 11 Jan 2025"
    $compliantVer  = "10.0.26100.7463" #oldversion
    $patchPath     = "$patchRoot\OS Updates\kb5074109"#kb5072033"<-December
    $patchName     = "kb5074109.msu"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\system32\ntoskrnl.exe"
}
"win11Jan" {
    $installTimeout = 60
    #Jan 2025
    $tag           = "PsExec"
    $listPath      = "$listPathRoot\Microsoft_Krnl.txt"
    $software      = "Win 11 Jan 2025"
    $compliantVer  = "10.0.22621.6489"
    $patchPath     = "$patchRoot\OS Updates\kb5073455"
    $patchName     = "kb5073455.msu"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\system32\ntoskrnl.exe"
}

"win1124h2dec" {
    $installTimeout = 60
    #Dec 2025
    $tag           = "PsExec"
    $listPath      = "$listPathRoot\Microsoft_Krnl.txt"
    $software      = "Win 11 Jan 2025"
    $compliantVer  = "10.0.26100.7463" #oldversion
    $patchPath     = "$patchRoot\OS Updates\kb5074109"#kb5072033"<-December
    $patchName     = "kb5074109.msu"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\system32\ntoskrnl.exe"
}
"win11dec" {
    $installTimeout = 60
    #Dec 2025
    $tag           = "PsExec"
    $listPath      = "$listPathRoot\Microsoft_Krnl.txt"
    $software      = "Win 11 Dec 2025"
    $compliantVer  = "10.0.22621.6345" #oldversion
    $patchPath     = "$patchRoot\OS Updates\kb5071417"
    $patchName     = "kb5071417.msu"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\system32\ntoskrnl.exe"
}



"win1122h2" {
    $installTimeout = 60
    #Feb 2024
    $tag           = "PsExec"
    $listPath      = "$listPathRoot\Microsoft_Krnl.txt"
    $software      = "Win 11 22H2 Feb 2024"
    $compliantVer  = "10.0.22621.3155"
    $patchPath     = "$patchRoot\OS Updates\kb5034765"
    $patchName     = "kb5034765.msu"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\system32\ntoskrnl.exe"
}
"win1121h2" {
    $installTimeout = 60
    #Feb 2024
    $tag           = "PsExec"
    $listPath      = "$listPathRoot\Microsoft_Krnl.txt"
    $software      = "Win 11 21H2 Feb 2024"
    $compliantVer  = "10.0.22000.2777"
    $patchPath     = "$patchRoot\OS Updates\kb5034766"
    $patchName     = "kb5034766.msu"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\system32\ntoskrnl.exe"
}
"win10" {
    $installTimeout = 60
    #Feb 2024
    $tag           = "PsExec"
    $listPath      = "$listPathRoot\Microsoft_Krnl.txt"
    $software      = "Win 10 Feb 2024"
    $compliantVer  = "10.0.19041.4046"
    $patchPath     = "$patchRoot\OS Updates\kb5034763"
    $patchName     = "kb5034763.msu"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\system32\ntoskrnl.exe"
}
"win11old" {
    $installTimeout = 60
    #Jan 2023
    $tag           = "PsExec"
    $listPath      = "$listPathRoot\Microsoft_Krnl.txt"
    $software      = "Win 11 21H2 Jan 2023"
    $compliantVer  = "10.0.22000.1335"
    $patchPath     = "$patchRoot\OS Updates\KB5022287"
    $patchName     = "KB5022287.msu"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\system32\ntoskrnl.exe"
}

#endregion  ----  Windows Kernel  ----




#region  ---  DotNet Framework  ---

"Dotnettest" {
    $installTimeout = 30
    $listPath      = "$listPathRoot\Microsoft_dotNet.txt"
    $software      = "dotNet"
    $processName   = "PROCESS"
    $compliantVer  = "8.0.2300"
    $patchPath     = "$patchRoot\OS Updates\NetTest" # No patch
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default.ps1").ScriptBlock
    $softwarePaths = "C:\Program Files\dotnet\shared\Microsoft.NETCore.App\8.0.11\mscorlib.dll"
    #$softwarePaths = "C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App\6.0.36\Microsoft.AspNetCore.Identity.dll"
    $installLine   = @(
        #"Set-Location 'C:\Temp\NetTest';"
        '& cmd /c "c:\temp\nettest\aspnetcore-runtime-8.0.20.exe" /install /quiet;'
        '& cmd /c "c:\temp\nettest\aspnetcore-runtime-8.0.20.exe" /uninstall /quiet;'  )
        #'& cmd /c "c:\temp\nettest\aspnetcore-runtime-6.0.36.exe" /uninstall /quiet')
}
"dotNetjan1" {
    $installTimeout = 30
    #Jan 2024, .Net Plugin
    $tag           = "PsExec"
    $KB            = "KB5033909"
    $software      = "dotNet Jan 2024"
    $listPath      = "$listPathRoot\Microsoft_dotNET.txt"
    $compliantVer  = "4.8.4690.0"
    $patchPath     = "$patchRoot\OS Updates\$KB"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $patchName     = "KB5033909.msu"
    $softwarePaths = "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\system.web.dll"
}
"dotNetjan2" {
    $installTimeout = 30
    #Jan 2024, .Net Plugin
    $tag           = "PsExec"
    $KB            = "KB5033920"
    $software      = "dotNet Jan 2024"
    $listPath      = "$listPathRoot\Microsoft_dotNET.txt"
    $compliantVer  = "4.8.9214.0"
    $patchPath     = "$patchRoot\OS Updates\$KB"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $patchName     = "KB5033920.msu"
    $softwarePaths = "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\system.web.dll"
}
"dotNetjan3" {
    $installTimeout = 30
    #Jan 2024, .Net Plugin
    $tag           = "PsExec"
    $KB            = "KB5033912"
    $software      = "dotNet Jan 2024"
    $listPath      = "$listPathRoot\Microsoft_dotNET.txt"
    $compliantVer  = "4.8.4690.0"
    $patchPath     = "$patchRoot\OS Updates\$KB"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $patchName     = "KB5033912.msu"
    $softwarePaths = "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\system.web.dll"
}
"dotNetfeb1" {
    $installTimeout = 30
    #Feb 2024, Win10
    $tag           = "PsExec"
    $KB            = "KB5034468"
    $software      = "dotNet Feb 2024"
    $listPath      = "$listPathRoot\Microsoft_dotNET.txt"
    $compliantVer  = "4.8.4690.0"
    $patchPath     = "$patchRoot\OS Updates\$KB"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $patchName     = "KB5034468.msu"
    $softwarePaths = "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\system.web.dll"
}
"dotNetfeb2" {
    $installTimeout = 30
    #Feb 2024, Win11
    $tag           = "PsExec"
    $KB            = "KB5034467"
    $software      = "dotNet Feb 2024"
    $listPath      = "$listPathRoot\Microsoft_dotNET.txt"
    $compliantVer  = "4.8.4690.0"
    $patchPath     = "$patchRoot\OS Updates\$KB"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $patchName     = "KB5034467.msu"
    $softwarePaths = "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\system.web.dll"
}
"dotNetcore" {
    $installTimeout = 30
    #June 2023
    $tag           = "PsExec"
    $KB            = "kb5032007"
    $listPath      = "$listPathRoot\dotNetcore.txt"
    $software      = "dotNetcore"
    $compliantVer  = "4.8.9215.0" # 4.8.9214.0
    $patchPath     = "$patchRoot\OS Updates\$KB"
    $patchName     = "$KB.msu"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $softwarePaths = "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\system.core.dll"
}
"dotNetmscor" {
    $installTimeout = 30
    #Jan 2025
    $tag           = "PsExec"
    $KB            = "KB5049622"
    $software      = "dotNet April 2024"
    $listPath      = "$listPathRoot\Microsoft_dotNET.txt"
    $compliantVer  = "4.8.9290.0"
    $patchPath     = "$patchRoot\OS Updates\$KB"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $patchName     = "KB5049622.msu"
    $softwarePaths = "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\mscorlib.dll"
}

"dotNetApril" {
    $installTimeout = 30
    #Feb 2024, Win11
    $tag           = "PsExec"
    $KB            = "KB5036620"
    $software      = "dotNet April 2024"
    $listPath      = "$listPathRoot\Microsoft_dotNET.txt"
    $compliantVer  = "4.8.9236.0"
    $patchPath     = "$patchRoot\OS Updates\$KB"
    $patchScript   = (Get-Command "$scriptRoot\Patching\Default-PSExec.ps1").ScriptBlock
    $patchName     = "KB5036620.msu"
    $softwarePaths = "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\system.runtime.serialization.dll"
}

#endregion  ---  DotNet Framework  ---

#endregion  ---  Microsoft Update Packages (.msu)  ---

} # End switch


# Any key and value created above must also be added here
$parameters = @{
    Tag           = $tag
    Software      = $software
    ListPath      = $listPath
    CompliantVer  = $compliantVer
    VersionType   = $versionType
    PatchPath     = $patchPath
    PatchName     = $patchName
    ProcessName   = $processName
    PatchScript   = $patchScript
    SoftwarePaths = $softwarePaths
    RegistryKey   = $registryKey
    InstallLine      = $installLine
    InstallTimeout   = $installTimeout
    KB               = $KB
}

Return $parameters
