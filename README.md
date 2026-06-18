# Microsoft Office & Apps Troubleshooter

A menu-driven PowerShell toolkit for basic Microsoft Office, Microsoft 365 Apps, OneDrive, Teams, Microsoft Store, and Windows app troubleshooting.

## What it does

The script includes options to:

- Collect Office and Microsoft app diagnostics
- Close stuck Office and Teams processes
- Restart the Office Click-to-Run service
- Open Microsoft Office repair settings
- Check for Microsoft 365 Apps updates
- Start Word, Excel, PowerPoint, or Outlook in Safe Mode
- Reset OneDrive
- Clear Microsoft Teams cache
- Reset Microsoft Store cache
- Reset a selected Windows Store app
- Run System File Checker with `sfc /scannow`
- Run DISM RestoreHealth
- Test basic Microsoft 365 connectivity
- Open Windows Installed Apps settings

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1 or later
- Administrator rights recommended

## How to run

Open PowerShell as Administrator, go to the folder where the script is saved, then run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Microsoft_Office_Apps_Troubleshooter.ps1
```

The script attempts to relaunch itself as Administrator if it is not already elevated.

## Logging

Logs are saved on the desktop in:

```text
MS_Apps_Troubleshooter_Logs
```

Each run creates a timestamped log file.

## Safety notes

Some options can close apps, clear cache files, or reset local app settings. The script asks for confirmation before running actions that may affect user sessions or local app data.

OneDrive reset does not delete cloud files, but it can rebuild the local sync state and may take time to resync.

Teams cache clearing can require the user to sign in again.

## File

Main script:

```text
Microsoft_Office_Apps_Troubleshooter.ps1
```
