## Windows Updater Utility
Simple CLI utility written in PowerShell to update Windows and the various programs

## Disclaimer
Still a work in progress and the update process need to be tested more, to be sure they work.

## How to use
Run the terminal as Administrator then run the command below, because there is some bug (that i still need to fix) that prevent the elevation to Admin.

Run this command in the terminal to run the script directly without downloading it from GitHub:
```PowerShell
irm https://raw.githubusercontent.com/BjornTheThunder/windows-updater/refs/heads/main/windows_updater.ps1 | iex
```
