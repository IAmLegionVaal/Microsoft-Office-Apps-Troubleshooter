#requires -Version 5.1
<#
.SYNOPSIS
    Menu-driven Microsoft Office and Microsoft app troubleshooting toolkit.

.DESCRIPTION
    Basic support script for Microsoft 365 Apps, OneDrive, Teams, Microsoft Store apps,
    Windows app repair, SFC/DISM, and Microsoft 365 connectivity checks.

.NOTES
    Run from an elevated PowerShell window on Windows 10/11.
    Some actions close apps, clear caches, or reset local app state.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdministrator)) {
    try {
        Write-Host 'Administrator rights are recommended. Requesting elevation...' -ForegroundColor Yellow
        Start-Process powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',"`"$PSCommandPath`"") -Verb RunAs
        exit
    }
    catch {
        Write-Warning "Could not elevate automatically: $($_.Exception.Message)"
    }
}

$DesktopPath = [Environment]::GetFolderPath('Desktop')
$LogFolder = Join-Path $DesktopPath 'MS_Apps_Troubleshooter_Logs'
New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
$LogFile = Join-Path $LogFolder ("Troubleshooter_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

function Write-Log {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')] [string]$Level = 'INFO'
    )

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $LogFile -Value $line -Encoding UTF8

    switch ($Level) {
        'WARN'    { Write-Host $Message -ForegroundColor Yellow }
        'ERROR'   { Write-Host $Message -ForegroundColor Red }
        'SUCCESS' { Write-Host $Message -ForegroundColor Green }
        default   { Write-Host $Message }
    }
}

function Pause-Menu {
    Write-Host
    [void](Read-Host 'Press Enter to return to the menu')
}

function Confirm-Action {
    param([Parameter(Mandatory)] [string]$Message)
    $answer = Read-Host "$Message Type YES to continue"
    return ($answer -eq 'YES')
}

function Show-Header {
    Clear-Host
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host '   MICROSOFT OFFICE & APPS TROUBLESHOOTER' -ForegroundColor Cyan
    Write-Host "   Version $ScriptVersion" -ForegroundColor DarkCyan
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ("   Computer : {0}" -f $env:COMPUTERNAME)
    Write-Host ("   User     : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    Write-Host ("   Admin    : {0}" -f (Test-IsAdministrator))
    Write-Host ("   Log      : {0}" -f $LogFile)
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host
}

function Get-FirstExistingPath {
    param([Parameter(Mandatory)] [string[]]$Paths)
    foreach ($path in $Paths) {
        if ($path -and (Test-Path -LiteralPath $path)) { return $path }
    }
    return $null
}

function Get-OfficeC2RClientPath {
    $paths = @((Join-Path $env:ProgramFiles 'Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe'))
    if (${env:ProgramFiles(x86)}) {
        $paths += Join-Path ${env:ProgramFiles(x86)} 'Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe'
    }
    return Get-FirstExistingPath -Paths $paths
}

function Get-OneDrivePath {
    $paths = @(
        (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft OneDrive\OneDrive.exe')
    )
    if (${env:ProgramFiles(x86)}) {
        $paths += Join-Path ${env:ProgramFiles(x86)} 'Microsoft OneDrive\OneDrive.exe'
    }
    return Get-FirstExistingPath -Paths $paths
}

function Get-OfficeAppPath {
    param([Parameter(Mandatory)] [string]$Executable)

    $command = Get-Command $Executable -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    $folders = @(
        (Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16'),
        (Join-Path $env:ProgramFiles 'Microsoft Office\Office16')
    )
    if (${env:ProgramFiles(x86)}) {
        $folders += Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16'
        $folders += Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\Office16'
    }

    foreach ($folder in $folders) {
        $candidate = Join-Path $folder $Executable
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    return $null
}

function Stop-MicrosoftAppProcesses {
    param([switch]$IncludeTeams)

    $processNames = @('WINWORD','EXCEL','POWERPNT','OUTLOOK','ONENOTE','MSACCESS','MSPUB','VISIO','LYNC')
    if ($IncludeTeams) { $processNames += @('Teams','ms-teams') }

    $stopped = 0
    foreach ($name in $processNames) {
        Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Stop-Process -Id $_.Id -Force -ErrorAction Stop
                Write-Log "Stopped process $($_.ProcessName) PID $($_.Id)."
                $stopped++
            }
            catch {
                Write-Log "Could not stop $($_.ProcessName): $($_.Exception.Message)" 'WARN'
            }
        }
    }

    if ($stopped -eq 0) { Write-Log 'No matching Office or Teams processes were running.' }
    else { Write-Log "Stopped $stopped process(es)." 'SUCCESS' }
}

function Get-MicrosoftDiagnostics {
    Show-Header
    Write-Host '[1] Collecting diagnostics...' -ForegroundColor Cyan

    $reportPath = Join-Path $LogFolder ("Diagnostics_{0}.txt" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $report = New-Object System.Collections.Generic.List[string]

    $report.Add('MICROSOFT OFFICE & APPS DIAGNOSTIC REPORT')
    $report.Add(('Generated: {0}' -f (Get-Date)))
    $report.Add(('Computer: {0}' -f $env:COMPUTERNAME))
    $report.Add(('User: {0}\{1}' -f $env:USERDOMAIN, $env:USERNAME))
    $report.Add(('Administrator: {0}' -f (Test-IsAdministrator)))
    $report.Add(('PowerShell: {0}' -f $PSVersionTable.PSVersion))
    $report.Add('')

    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $report.Add('WINDOWS')
        $report.Add(('Name: {0}' -f $os.Caption))
        $report.Add(('Version: {0}' -f $os.Version))
        $report.Add(('Build: {0}' -f $os.BuildNumber))
        $report.Add(('Last boot: {0}' -f $os.LastBootUpTime))
        $report.Add('')
    }
    catch { $report.Add(("Windows information error: {0}" -f $_.Exception.Message)) }

    try {
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($env:SystemDrive)'"
        $report.Add('DISK')
        $report.Add(('{0}: {1} GB free of {2} GB' -f $env:SystemDrive, [math]::Round($disk.FreeSpace / 1GB,2), [math]::Round($disk.Size / 1GB,2)))
        $report.Add('')
    }
    catch { $report.Add(("Disk information error: {0}" -f $_.Exception.Message)) }

    $report.Add('MICROSOFT 365 / OFFICE')
    $officeRegPath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path $officeRegPath) {
        try {
            $office = Get-ItemProperty $officeRegPath
            foreach ($property in @('ProductReleaseIds','VersionToReport','ClientCulture','Platform','UpdateChannel','CDNBaseUrl')) {
                $value = $office.$property
                if ($null -ne $value -and "$value" -ne '') { $report.Add(('{0}: {1}' -f $property, $value)) }
            }
        }
        catch { $report.Add(("Office registry read error: {0}" -f $_.Exception.Message)) }
    }
    else { $report.Add('Click-to-Run configuration was not found.') }

    $officeService = Get-Service -Name ClickToRunSvc -ErrorAction SilentlyContinue
    if ($officeService) { $report.Add(('ClickToRunSvc: {0} / StartType {1}' -f $officeService.Status, $officeService.StartType)) }
    else { $report.Add('ClickToRunSvc: Not found') }

    foreach ($app in @('WINWORD.EXE','EXCEL.EXE','POWERPNT.EXE','OUTLOOK.EXE')) {
        $path = Get-OfficeAppPath -Executable $app
        if ($path) { $report.Add(('{0}: {1} ({2})' -f $app, (Get-Item $path).VersionInfo.FileVersion, $path)) }
    }
    $report.Add('')

    $report.Add('ONEDRIVE')
    $oneDrivePath = Get-OneDrivePath
    if ($oneDrivePath) {
        $report.Add(('Path: {0}' -f $oneDrivePath))
        $report.Add(('Version: {0}' -f (Get-Item $oneDrivePath).VersionInfo.FileVersion))
        $report.Add(('Running: {0}' -f [bool](Get-Process -Name OneDrive -ErrorAction SilentlyContinue)))
    }
    else { $report.Add('OneDrive executable was not found.') }
    $report.Add('')

    $report.Add('TEAMS AND STORE PACKAGES')
    foreach ($packageName in @('MSTeams','Microsoft.WindowsStore')) {
        try {
            $package = Get-AppxPackage -Name $packageName -ErrorAction SilentlyContinue
            if ($package) { $report.Add(('{0}: Version {1}, Status {2}' -f $package.Name, $package.Version, $package.Status)) }
            else { $report.Add(('{0}: Not found' -f $packageName)) }
        }
        catch { $report.Add(('{0}: Query error - {1}' -f $packageName, $_.Exception.Message)) }
    }
    $report.Add('')

    $report.Add('SERVICES')
    foreach ($serviceName in @('ClickToRunSvc','AppXSvc','ClipSVC','InstallService','BITS','wuauserv')) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($service) { $report.Add(('{0}: {1} / StartType {2}' -f $serviceName, $service.Status, $service.StartType)) }
        else { $report.Add(('{0}: Not found' -f $serviceName)) }
    }
    $report.Add('')

    $report.Add('WINHTTP PROXY')
    try { (& netsh.exe winhttp show proxy 2>&1) | ForEach-Object { $report.Add([string]$_) } }
    catch { $report.Add(("Proxy query error: {0}" -f $_.Exception.Message)) }

    $report | Set-Content -Path $reportPath -Encoding UTF8
    Write-Log "Diagnostic report created: $reportPath" 'SUCCESS'
    Start-Process notepad.exe -ArgumentList "`"$reportPath`""
    Pause-Menu
}

function Close-OfficeAndTeams {
    Show-Header
    Write-Host '[2] Close Office and Teams processes' -ForegroundColor Cyan
    Write-Host 'WARNING: Unsaved work in open applications will be lost.' -ForegroundColor Yellow
    if (Confirm-Action 'Close Microsoft Office and Teams now?') { Stop-MicrosoftAppProcesses -IncludeTeams }
    else { Write-Log 'Operation cancelled.' }
    Pause-Menu
}

function Restart-OfficeClickToRun {
    Show-Header
    Write-Host '[3] Restart Office Click-to-Run service' -ForegroundColor Cyan
    $service = Get-Service -Name ClickToRunSvc -ErrorAction SilentlyContinue
    if (-not $service) {
        Write-Log 'Office Click-to-Run service was not found. Office may use an MSI installation.' 'WARN'
        Pause-Menu
        return
    }
    try {
        if ($service.Status -eq 'Running') { Restart-Service -Name ClickToRunSvc -Force }
        else { Start-Service -Name ClickToRunSvc }
        Write-Log 'Office Click-to-Run service restart command completed.' 'SUCCESS'
    }
    catch { Write-Log "Could not restart ClickToRunSvc: $($_.Exception.Message)" 'ERROR' }
    Pause-Menu
}

function Open-OfficeRepair {
    Show-Header
    Write-Host '[4] Open Microsoft Office repair' -ForegroundColor Cyan
    Write-Host 'Quick Repair is faster and works offline. Online Repair is more thorough.' -ForegroundColor Yellow
    if (-not (Confirm-Action 'Open Windows Installed Apps so you can repair Office?')) {
        Write-Log 'Operation cancelled.'
        Pause-Menu
        return
    }
    try {
        Start-Process 'ms-settings:appsfeatures'
        Write-Log 'Opened Installed Apps. Search for Microsoft 365 or Office, choose Modify, then Quick Repair or Online Repair.' 'SUCCESS'
    }
    catch { Write-Log "Could not open Installed Apps: $($_.Exception.Message)" 'ERROR' }
    Pause-Menu
}

function Update-Microsoft365 {
    Show-Header
    Write-Host '[5] Check for Microsoft 365 Apps updates' -ForegroundColor Cyan
    $clientPath = Get-OfficeC2RClientPath
    if (-not $clientPath) {
        Write-Log 'OfficeC2RClient.exe was not found. This option requires Click-to-Run Office.' 'WARN'
        Pause-Menu
        return
    }
    Write-Host 'Office apps may need to close before an update can complete.' -ForegroundColor Yellow
    if (Confirm-Action 'Start Microsoft 365 update check?') {
        try {
            Start-Process -FilePath $clientPath -ArgumentList '/update user displaylevel=true forceappshutdown=false'
            Write-Log 'Microsoft 365 update check started.' 'SUCCESS'
        }
        catch { Write-Log "Could not start update check: $($_.Exception.Message)" 'ERROR' }
    }
    else { Write-Log 'Operation cancelled.' }
    Pause-Menu
}

function Start-OfficeSafeMode {
    Show-Header
    Write-Host '[6] Start an Office app in Safe Mode' -ForegroundColor Cyan
    Write-Host '  1. Microsoft Word'
    Write-Host '  2. Microsoft Excel'
    Write-Host '  3. Microsoft PowerPoint'
    Write-Host '  4. Microsoft Outlook'
    Write-Host '  0. Back'
    Write-Host
    $choice = Read-Host 'Select an app'
    $selection = switch ($choice) {
        '1' { @{ Name='Word';       Exe='WINWORD.EXE';  Args='/safe' } }
        '2' { @{ Name='Excel';      Exe='EXCEL.EXE';    Args='/safe' } }
        '3' { @{ Name='PowerPoint'; Exe='POWERPNT.EXE'; Args='/safe' } }
        '4' { @{ Name='Outlook';    Exe='OUTLOOK.EXE';  Args='/safe' } }
        default { $null }
    }
    if (-not $selection) { return }
    $path = Get-OfficeAppPath -Executable $selection.Exe
    if (-not $path) { Write-Log "$($selection.Name) executable was not found." 'WARN' }
    else {
        try {
            Start-Process -FilePath $path -ArgumentList $selection.Args
            Write-Log "$($selection.Name) started in Safe Mode." 'SUCCESS'
        }
        catch { Write-Log "Could not start $($selection.Name): $($_.Exception.Message)" 'ERROR' }
    }
    Pause-Menu
}

function Reset-OneDriveClient {
    Show-Header
    Write-Host '[7] Reset OneDrive' -ForegroundColor Cyan
    Write-Host 'This does not delete cloud files, but OneDrive will rebuild its sync state.' -ForegroundColor Yellow
    $oneDrivePath = Get-OneDrivePath
    if (-not $oneDrivePath) {
        Write-Log 'OneDrive.exe was not found.' 'WARN'
        Pause-Menu
        return
    }
    if (-not (Confirm-Action 'Reset OneDrive now?')) {
        Write-Log 'Operation cancelled.'
        Pause-Menu
        return
    }
    try {
        Get-Process OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Process -FilePath $oneDrivePath -ArgumentList '/reset'
        Write-Log 'OneDrive reset command started.'
        Start-Sleep -Seconds 8
        if (-not (Get-Process OneDrive -ErrorAction SilentlyContinue)) {
            Start-Process -FilePath $oneDrivePath
            Write-Log 'OneDrive was started manually after the reset.'
        }
        Write-Log 'OneDrive reset completed or is continuing.' 'SUCCESS'
    }
    catch { Write-Log "OneDrive reset failed: $($_.Exception.Message)" 'ERROR' }
    Pause-Menu
}

function Clear-TeamsCache {
    Show-Header
    Write-Host '[8] Clear Microsoft Teams cache' -ForegroundColor Cyan
    Write-Host 'Teams will rebuild its cache when reopened. You may need to sign in again.' -ForegroundColor Yellow
    if (-not (Confirm-Action 'Close Teams and clear its cache?')) {
        Write-Log 'Operation cancelled.'
        Pause-Menu
        return
    }
    try {
        Get-Process -Name 'Teams','ms-teams' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        $cachePaths = @(
            (Join-Path $env:APPDATA 'Microsoft\Teams'),
            (Join-Path $env:LOCALAPPDATA 'Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams')
        )
        $cleared = 0
        foreach ($cachePath in $cachePaths) {
            if (Test-Path -LiteralPath $cachePath) {
                Get-ChildItem -LiteralPath $cachePath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction Stop
                Write-Log "Cleared Teams cache: $cachePath"
                $cleared++
            }
        }
        if ($cleared -eq 0) { Write-Log 'No supported Teams cache folder was found.' 'WARN' }
        else { Write-Log 'Teams cache cleared. Reopen Teams manually.' 'SUCCESS' }
    }
    catch { Write-Log "Teams cache clear failed: $($_.Exception.Message)" 'ERROR' }
    Pause-Menu
}

function Reset-MicrosoftStoreCache {
    Show-Header
    Write-Host '[9] Reset Microsoft Store cache' -ForegroundColor Cyan
    try {
        Start-Process -FilePath 'wsreset.exe'
        Write-Log 'Microsoft Store cache reset started.' 'SUCCESS'
    }
    catch { Write-Log "Could not start WSReset.exe: $($_.Exception.Message)" 'ERROR' }
    Pause-Menu
}

function Reset-SelectedStoreApp {
    Show-Header
    Write-Host '[10] Reset a selected Windows Store app' -ForegroundColor Cyan
    Write-Host 'WARNING: Resetting an app can remove local settings and sign-in state.' -ForegroundColor Yellow
    Write-Host
    try {
        $packages = Get-AppxPackage |
            Where-Object { $_.IsFramework -eq $false -and $_.IsResourcePackage -eq $false -and $_.NonRemovable -eq $false } |
            Sort-Object Name |
            Select-Object -First 60

        if (-not $packages) {
            Write-Log 'No resettable app packages were found.' 'WARN'
            Pause-Menu
            return
        }

        for ($index = 0; $index -lt $packages.Count; $index++) {
            Write-Host ('{0,3}. {1}' -f ($index + 1), $packages[$index].Name)
        }
        Write-Host
        $selection = Read-Host 'Enter the app number, or 0 to cancel'
        $number = 0
        if (-not [int]::TryParse($selection, [ref]$number)) {
            Write-Log 'Invalid selection.' 'WARN'
            Pause-Menu
            return
        }
        if ($number -eq 0) { return }
        if ($number -lt 1 -or $number -gt $packages.Count) {
            Write-Log 'Selection is outside the valid range.' 'WARN'
            Pause-Menu
            return
        }

        $package = $packages[$number - 1]
        if (-not (Confirm-Action "Reset $($package.Name)?")) {
            Write-Log 'Operation cancelled.'
            Pause-Menu
            return
        }

        if (Get-Command Reset-AppxPackage -ErrorAction SilentlyContinue) {
            Reset-AppxPackage -Package $package.PackageFullName
            Write-Log "Reset app package: $($package.Name)" 'SUCCESS'
        }
        else {
            $manifestPath = Join-Path $package.InstallLocation 'AppxManifest.xml'
            if (-not (Test-Path -LiteralPath $manifestPath)) { throw "App manifest was not found for $($package.Name)." }
            Add-AppxPackage -DisableDevelopmentMode -Register $manifestPath
            Write-Log "Reset-AppxPackage is unavailable; re-registered $($package.Name) instead." 'SUCCESS'
        }
    }
    catch { Write-Log "App reset failed: $($_.Exception.Message)" 'ERROR' }
    Pause-Menu
}

function Run-SfcScan {
    Show-Header
    Write-Host '[11] Run System File Checker' -ForegroundColor Cyan
    Write-Host 'This checks and repairs protected Windows system files.' -ForegroundColor Yellow
    if (-not (Confirm-Action 'Run SFC /scannow now?')) {
        Write-Log 'Operation cancelled.'
        Pause-Menu
        return
    }
    try {
        $process = Start-Process -FilePath 'sfc.exe' -ArgumentList '/scannow' -Wait -PassThru -NoNewWindow
        Write-Log "SFC finished with exit code $($process.ExitCode)." 'SUCCESS'
    }
    catch { Write-Log "SFC failed: $($_.Exception.Message)" 'ERROR' }
    Pause-Menu
}

function Run-DismRestoreHealth {
    Show-Header
    Write-Host '[12] Run DISM RestoreHealth' -ForegroundColor Cyan
    Write-Host 'This repairs the Windows component store used by system repairs.' -ForegroundColor Yellow
    Write-Host 'An internet connection may be required.' -ForegroundColor Yellow
    if (-not (Confirm-Action 'Run DISM /RestoreHealth now?')) {
        Write-Log 'Operation cancelled.'
        Pause-Menu
        return
    }
    try {
        $process = Start-Process -FilePath 'dism.exe' -ArgumentList @('/Online','/Cleanup-Image','/RestoreHealth') -Wait -PassThru -NoNewWindow
        Write-Log "DISM finished with exit code $($process.ExitCode)." 'SUCCESS'
    }
    catch { Write-Log "DISM failed: $($_.Exception.Message)" 'ERROR' }
    Pause-Menu
}

function Test-MicrosoftConnectivity {
    Show-Header
    Write-Host '[13] Test basic Microsoft 365 connectivity' -ForegroundColor Cyan
    Write-Host 'This checks DNS resolution and TCP port 443 only.' -ForegroundColor Yellow
    Write-Host
    $targets = @('login.microsoftonline.com','www.office.com','graph.microsoft.com','onedrive.com','teams.microsoft.com')
    foreach ($target in $targets) {
        $dnsStatus = 'Failed'
        $httpsStatus = 'Failed'
        try { [void][System.Net.Dns]::GetHostAddresses($target); $dnsStatus = 'OK' } catch { $dnsStatus = 'Failed' }
        try {
            if (Test-NetConnection -ComputerName $target -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue) { $httpsStatus = 'OK' }
        }
        catch { $httpsStatus = 'Failed' }
        $line = '{0,-32} DNS: {1,-6} HTTPS/443: {2}' -f $target, $dnsStatus, $httpsStatus
        Write-Host $line
        Write-Log $line
    }
    Write-Host
    Write-Host 'Current WinHTTP proxy:' -ForegroundColor Cyan
    & netsh.exe winhttp show proxy
    Pause-Menu
}

function Open-WindowsAppSettings {
    Show-Header
    Write-Host '[14] Open Windows app repair settings' -ForegroundColor Cyan
    try {
        Start-Process 'ms-settings:appsfeatures'
        Write-Log 'Opened Installed Apps settings.' 'SUCCESS'
    }
    catch { Write-Log "Could not open Installed Apps settings: $($_.Exception.Message)" 'ERROR' }
    Pause-Menu
}

Write-Log "Microsoft Office & Apps Troubleshooter $ScriptVersion started."
Write-Log "Administrator: $(Test-IsAdministrator)"

do {
    Show-Header
    Write-Host '  1. Collect Office and app diagnostics'
    Write-Host '  2. Close Office and Teams processes'
    Write-Host '  3. Restart Office Click-to-Run service'
    Write-Host '  4. Open Microsoft Office repair'
    Write-Host '  5. Check for Microsoft 365 Apps updates'
    Write-Host '  6. Start an Office app in Safe Mode'
    Write-Host '  7. Reset OneDrive'
    Write-Host '  8. Clear Microsoft Teams cache'
    Write-Host '  9. Reset Microsoft Store cache'
    Write-Host ' 10. Reset a selected Windows Store app'
    Write-Host ' 11. Run System File Checker (SFC)'
    Write-Host ' 12. Run DISM RestoreHealth'
    Write-Host ' 13. Test Microsoft 365 connectivity'
    Write-Host ' 14. Open Windows Installed Apps settings'
    Write-Host
    Write-Host '  0. Exit'
    Write-Host

    $menuChoice = Read-Host 'Select an option'

    switch ($menuChoice) {
        '1'  { Get-MicrosoftDiagnostics }
        '2'  { Close-OfficeAndTeams }
        '3'  { Restart-OfficeClickToRun }
        '4'  { Open-OfficeRepair }
        '5'  { Update-Microsoft365 }
        '6'  { Start-OfficeSafeMode }
        '7'  { Reset-OneDriveClient }
        '8'  { Clear-TeamsCache }
        '9'  { Reset-MicrosoftStoreCache }
        '10' { Reset-SelectedStoreApp }
        '11' { Run-SfcScan }
        '12' { Run-DismRestoreHealth }
        '13' { Test-MicrosoftConnectivity }
        '14' { Open-WindowsAppSettings }
        '0'  {
            Write-Log 'Troubleshooter closed by the user.'
            Write-Host 'Goodbye.' -ForegroundColor Green
        }
        default {
            Write-Host 'Invalid selection.' -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }
    }
}
while ($menuChoice -ne '0')
