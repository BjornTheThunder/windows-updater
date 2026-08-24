## Windows Updater Utility
Simple CLI utility written in PowerShell to update Windows and the various programs

## Disclaimer
Still a work in progress and the update process need to be tested more to be sure they work.

## How to use
Run this command in the terminal to run the script directly without downloading it from GitHub.

This command execute a simple scan for updates:
```PowerShell
irm https://raw.githubusercontent.com/BjornTheThunder/windows-updater/refs/heads/main/windows_updater.ps1 | iex
```

Add the -Install parameter to also apply the updates:
```PowerShell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/BjornTheThunder/windows-updater/refs/heads/main/windows_updater.ps1))) -Install
```
