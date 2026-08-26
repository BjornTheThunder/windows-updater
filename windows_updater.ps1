# ==============================================================================
# Windows Multi-Manager Update & Inventory Script
# ==============================================================================

[CmdletBinding()]
param(
    [switch]$Detailed,
    [switch]$ExportCSV,
    [switch]$Install
)

# Force UTF-8 Encoding for PowerShell console & external CLI tools output
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "Windows Multi-Manager Update and Inventory"

# ------------------------------------------------------------------------------
# CORE VARIABLE INITIALIZATION
# ------------------------------------------------------------------------------
$activityTitle = "Windows Multi-Manager Update and Inventory"
$totalSteps = 6
$updateResults = [System.Collections.Generic.List[PSCustomObject]]::new()

# State trackers for installation phase
$pendingWinUpdates  = $null
$hasWinGetUpdates   = $false
$hasChocoUpdates    = $false
$hasScoopUpdates    = $false

# ------------------------------------------------------------------------------
# UI HELPER FUNCTIONS
# ------------------------------------------------------------------------------
function Write-Header {
    param([string]$Text)
    Write-Host ""
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host "==========================================================================" -ForegroundColor Cyan
}

function Write-SectionHeader {
    param([int]$Step, [string]$Title)
    Write-Host ""
    Write-Host "[$Step/$totalSteps] $Title" -ForegroundColor Yellow
    Write-Host ("-" * ($Title.Length + 7)) -ForegroundColor DarkGray
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Debug')]$Type = 'Info'
    )
    switch ($Type) {
        'Success' { Write-Host " [v] $Message" -ForegroundColor Green }
        'Warning' { Write-Host " [!] $Message" -ForegroundColor Yellow }
        'Error'   { Write-Host " [x] $Message" -ForegroundColor Red }
        'Debug'   { Write-Host " [.] $Message" -ForegroundColor DarkGray }
        Default   { Write-Host " [i] $Message" -ForegroundColor Gray }
    }
}

# ------------------------------------------------------------------------------
# 0. PRIVILEGE & ELEVATION CHECK
# ------------------------------------------------------------------------------
Write-Header "WINDOWS MULTI-MANAGER UPDATE AND INVENTORY"
Write-Log "Scan Started : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Type Info

$identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]$identity
$isAdmin   = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Log "Administrative privileges required. Requesting elevation..." -Type Warning

    # Define the URL where this script is hosted so it can re-call itself in-memory
    # IMPORTANT: Replace this with your actual GitHub RAW link!
    $ScriptUrl = "https://raw.githubusercontent.com/BjornTheThunder/windows-updater/refs/heads/main/windows_updater.ps1"

    # Reconstruct any passed parameters (-Install, -Detailed, -ExportCSV)
    $paramList = @()
    foreach ($key in $PSBoundParameters.Keys) {
        $val = $PSBoundParameters[$key]
        if ($val -is [switch] -and $val.IsPresent) {
            $paramList += "-$key"
        } elseif ($val -isnot [switch]) {
            $paramList += "-$key `"$val`""
        }
    }
    $paramsString = $paramList -join ' '
    
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        # 1. RUNNING IN-MEMORY (irm | iex)
        if ([string]::IsNullOrWhiteSpace($ScriptUrl) -or $ScriptUrl -eq "https://raw.githubusercontent.com/YourUsername/YourRepo/main/YourScript.ps1") {
            Write-Log "Running in-memory, but `$ScriptUrl is not set. Cannot auto-elevate." -Type Error
            Write-Host "Press any key to exit..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            return 
        }

        Write-Log "Re-launching script from web in an Elevated prompt..." -Type Info
        # Convert the web response directly into a scriptblock so we can pass arguments to it
        $elevateCmd = "& ([scriptblock]::Create((Invoke-RestMethod '$ScriptUrl'))) $paramsString"
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "`"$elevateCmd`""
        return 
        
    } else {
        # 2. RUNNING FROM A LOCAL FILE (.ps1)
        $argsToPass = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"") + $paramList
        Start-Process powershell.exe -Verb RunAs -ArgumentList ($argsToPass -join ' ')
        return 
    }
}

# ==============================================================================
# PHASE 1: SCANNING FOR UPDATES
# ==============================================================================

# ------------------------------------------------------------------------------
# 1. WINDOWS UPDATES (COM API)
# ------------------------------------------------------------------------------
$step = 1
Write-Progress -Activity $activityTitle -Status "Step $step of $totalSteps - Checking Windows Updates..." -PercentComplete (($step / $totalSteps) * 100) -Id 1
Write-SectionHeader -Step $step -Title "Windows Updates (COM API)"

try {
    Write-Log "Initializing Windows Update COM session..." -Type Debug
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()

    Write-Log "Querying Windows Update service..." -Type Info
    $searchResult = $updateSearcher.Search("IsInstalled=0 and IsHidden=0")
    $pendingWinUpdates = $searchResult.Updates

    if ($null -ne $pendingWinUpdates -and $pendingWinUpdates.Count -gt 0) {
        Write-Log "Found $($pendingWinUpdates.Count) pending Windows update(s):" -Type Success
        foreach ($update in $pendingWinUpdates) {
            $updateResults.Add([PSCustomObject]@{
                Source         = "Windows Update"
                Title          = $update.Title
                IsInstalled    = $update.IsInstalled
                IsMandatory    = $update.IsMandatory
                RebootRequired = ($update.RebootBehavior -eq 1)
            })
            Write-Host "    * $($update.Title)" -ForegroundColor Green
        }
    } else {
        Write-Log "Windows operating system is up to date." -Type Success
    }
} catch {
    Write-Log "Error processing Windows Updates: $($_.Exception.Message)" -Type Error
}

# ------------------------------------------------------------------------------
# 2. WINGET PACKAGES (Includes MS Store via WinGet)
# ------------------------------------------------------------------------------
$step = 2
Write-Progress -Activity $activityTitle -Status "Step $step of $totalSteps - Checking WinGet packages..." -PercentComplete (($step / $totalSteps) * 100) -Id 1
Write-SectionHeader -Step $step -Title "WinGet Packages"

if (Get-Command winget -ErrorAction SilentlyContinue) {
    try {
        Write-Log "Querying WinGet package repository..." -Type Info
        # WinGet might output a progress bar natively, we suppress errors
        $wingetRaw = winget upgrade --include-unknown --accept-source-agreements 2>$null
        $lines = $wingetRaw -split "`r?\n"

        $dividerIndex = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^-{5,}$') {
                $dividerIndex = $i
                break
            }
        }

        $wingetLines = @()
        if ($dividerIndex -ge 0 -and ($dividerIndex + 1) -lt $lines.Count) {
            $wingetLines = $lines[($dividerIndex + 1)..($lines.Count - 1)] | Where-Object {
                $_.Trim() -ne '' -and $_ -match '\s{2,}' -and $_ -notmatch '^\d+\s+'
            }
        }

        if ($wingetLines.Count -gt 0) {
            $hasWinGetUpdates = $true
            Write-Log "Found $($wingetLines.Count) update(s) via WinGet:" -Type Success

            foreach ($line in $wingetLines) {
                $trimmed = $line.Trim() -replace '\s{2,}', ' | '
                $updateResults.Add([PSCustomObject]@{
                    Source         = "WinGet"
                    Title          = $trimmed
                    IsInstalled    = $true
                    IsMandatory    = $false
                    RebootRequired = $false
                })
                Write-Host "    * $trimmed" -ForegroundColor Green
            }
        } else {
            Write-Log "No WinGet package updates available." -Type Success
        }
    } catch {
        Write-Log "Error processing WinGet: $($_.Exception.Message)" -Type Error
    }
} else {
    Write-Log "WinGet is not installed or available in PATH." -Type Debug
}

# ------------------------------------------------------------------------------
# 3. CHOCOLATEY PACKAGES
# ------------------------------------------------------------------------------
$step = 3
Write-Progress -Activity $activityTitle -Status "Step $step of $totalSteps - Checking Chocolatey packages..." -PercentComplete (($step / $totalSteps) * 100) -Id 1
Write-SectionHeader -Step $step -Title "Chocolatey Packages"

if (Get-Command choco -ErrorAction SilentlyContinue) {
    try {
        Write-Log "Querying Chocolatey repository..." -Type Info
        $chocoUpdates = choco outdated -r 2>$null

        if ($chocoUpdates) {
            $validUpdates = $chocoUpdates | Where-Object { $_ -match '^[^|]+\|[^|]+\|[^|]+\|' }

            if ($validUpdates) {
                $hasChocoUpdates = $true
                Write-Log "Found $($validUpdates.Count) Chocolatey update(s):" -Type Success

                foreach ($line in $validUpdates) {
                    $parts = $line.Split('|').Trim()
                    $updateResults.Add([PSCustomObject]@{
                        Source         = "Chocolatey"
                        Title          = "$($parts[0]) ($($parts[1]) -> $($parts[2]))"
                        IsInstalled    = $true
                        IsMandatory    = $false
                        RebootRequired = $false
                    })
                    Write-Host "    * $($parts[0]): $($parts[1]) to $($parts[2])" -ForegroundColor Green
                }
            } else {
                Write-Log "No Chocolatey updates available." -Type Success
            }
        } else {
            Write-Log "No Chocolatey updates available." -Type Success
        }
    } catch {
        Write-Log "Error processing Chocolatey: $($_.Exception.Message)" -Type Error
    }
} else {
    Write-Log "Chocolatey is not installed." -Type Debug
}

# ------------------------------------------------------------------------------
# 4. SCOOP PACKAGES
# ------------------------------------------------------------------------------
$step = 4
Write-Progress -Activity $activityTitle -Status "Step $step of $totalSteps - Checking Scoop packages..." -PercentComplete (($step / $totalSteps) * 100) -Id 1
Write-SectionHeader -Step $step -Title "Scoop Packages"

if (Get-Command scoop -ErrorAction SilentlyContinue) {
    # Scoop natively blocks execution when running as an Administrator.
    Write-Log "Skipped Scoop check: Scoop actively blocks running under an elevated Administrator context." -Type Warning
} else {
    Write-Log "Scoop is not installed." -Type Debug
}

# ------------------------------------------------------------------------------
# 5. INSTALLED PROGRAMS (Registry Inventory)
# ------------------------------------------------------------------------------
$step = 5
Write-Progress -Activity $activityTitle -Status "Step $step of $totalSteps - Scanning Registry Inventory..." -PercentComplete (($step / $totalSteps) * 100) -Id 1
Write-SectionHeader -Step $step -Title "Installed Programs Inventory"

try {
    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $installedApps = [System.Collections.Generic.List[PSCustomObject]]::new()
    $regCount = 0

    foreach ($path in $registryPaths) {
        $regCount++
        Write-Progress -Activity "Scanning Registry Keys" -Status "Path $regCount of $($registryPaths.Count)" -PercentComplete (($regCount / $registryPaths.Count) * 100) -ParentId 1 -Id 2

        if (Test-Path $path) {
            $items = Get-ItemProperty "$path\*" -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }
            foreach ($item in $items) {
                $installedApps.Add([PSCustomObject]@{
                    DisplayName    = $item.DisplayName
                    DisplayVersion = $item.DisplayVersion
                    Publisher      = $item.Publisher
                })
            }
        }
    }
    Write-Progress -Activity "Scanning Registry Keys" -Completed -Id 2

    Write-Log "Indexed $($installedApps.Count) installed application(s)." -Type Success

    if ($Detailed) {
        Write-Host ""
        $installedApps | Sort-Object DisplayName | Select-Object -First 50 | Format-Table -AutoSize
        if ($installedApps.Count -gt 50) {
            Write-Log "Truncated list at 50 items. Pass -ExportCSV for full inventory output." -Type Info
        }
    }
} catch {
    Write-Log "Error checking installed programs: $($_.Exception.Message)" -Type Error
}

# ------------------------------------------------------------------------------
# 6. MICROSOFT STORE APPS
# ------------------------------------------------------------------------------
$step = 6
Write-Progress -Activity $activityTitle -Status "Step $step of $totalSteps - Checking Store Packages..." -PercentComplete 100 -Id 1
Write-SectionHeader -Step $step -Title "Microsoft Store Apps"

Write-Log "Microsoft Store App updates are now evaluated dynamically by WinGet." -Type Debug

if ($Detailed) {
    Write-Log "Inventorying installed Microsoft Store packages..." -Type Info
    $storeApps = Get-AppxPackage -ErrorAction SilentlyContinue |
        Where-Object { $_.NonRemovable -ne $true -and $_.SignatureKind -eq 'Store' } |
        Select-Object Name, Version | Sort-Object Name

    if ($storeApps) {
        Write-Host ""
        $storeApps | Select-Object -First 25 | Format-Table -AutoSize
    }
}

Write-Progress -Activity $activityTitle -Completed -Id 1

# ==============================================================================
# PHASE 2: SUMMARY & PROMPT
# ==============================================================================
Write-Header "EXECUTION SUMMARY"

if ($updateResults.Count -gt 0) {
    Write-Log "Total Pending Updates Identified: $($updateResults.Count)" -Type Warning
    Write-Host ""
    $updateResults | Format-Table -Property Source, Title, RebootRequired -AutoSize
} else {
    Write-Log "No pending updates detected across any package manager." -Type Success
}

if ($ExportCSV) {
    $desktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
    $csvPath = Join-Path -Path $desktopPath -ChildPath "UpdateReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $updateResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Log "Report saved to: $csvPath" -Type Success
}

# Determine whether installation should run
$shouldInstall = $false

if ($updateResults.Count -gt 0) {
    if ($Install) {
        $shouldInstall = $true
    } else {
        Write-Host ""
        $userInput = Read-Host "Do you want to apply these updates now? [Y/n]"
        if ([string]::IsNullOrWhiteSpace($userInput) -or $userInput -match '^[Yy](es)?$') {
            $shouldInstall = $true
        } else {
            Write-Log "Update installation cancelled by user." -Type Info
        }
    }
}

# ==============================================================================
# PHASE 3: INSTALLATION
# ==============================================================================
if ($shouldInstall) {
    Write-Header "APPLYING UPDATES"

    # 1. Windows Updates
    if ($null -ne $pendingWinUpdates -and $pendingWinUpdates.Count -gt 0) {
        try {
            Write-Log "Preparing $($pendingWinUpdates.Count) Windows update(s) for installation..." -Type Warning

            # Must wrap array in a COM UpdateCollection for download/install methods
            $updatesCollection = New-Object -ComObject Microsoft.Update.UpdateColl
            foreach ($update in $pendingWinUpdates) {
                if (-not $update.EulaAccepted) { $update.AcceptEula() }
                $updatesCollection.Add($update) | Out-Null
            }

            Write-Progress -Activity "Downloading Windows Updates" -Status "Progressing download..." -PercentComplete 50 -Id 1
            $downloader = $updateSession.CreateUpdateDownloader()
            $downloader.Updates = $updatesCollection
            $downloader.Download()
            Write-Progress -Activity "Downloading Windows Updates" -Completed -Id 1

            Write-Log "Installing Windows updates..." -Type Warning
            Write-Progress -Activity "Installing Windows Updates" -Status "Applying updates..." -PercentComplete 75 -Id 1
            $installer = $updateSession.CreateUpdateInstaller()
            $installer.Updates = $updatesCollection
            $installResult = $installer.Install()
            Write-Progress -Activity "Installing Windows Updates" -Completed -Id 1

            Write-Log "Windows Update completed with Result Code: $($installResult.ResultCode)" -Type Success
        } catch {
            Write-Log "Error installing Windows Updates: $($_.Exception.Message)" -Type Error
        }
    }

    # 2. WinGet Packages
    if ($hasWinGetUpdates) {
        try {
            Write-Log "Upgrading all WinGet packages silently..." -Type Warning
            Write-Progress -Activity "WinGet Upgrade" -Status "Upgrading packages..." -Id 1
            winget upgrade --all --silent --accept-source-agreements --accept-package-agreements
            Write-Progress -Activity "WinGet Upgrade" -Completed -Id 1
            Write-Log "WinGet updates applied successfully." -Type Success
        } catch {
            Write-Log "Error upgrading WinGet packages: $($_.Exception.Message)" -Type Error
        }
    }

    # 3. Chocolatey Packages
    if ($hasChocoUpdates) {
        try {
            Write-Log "Upgrading all Chocolatey packages..." -Type Warning
            Write-Progress -Activity "Chocolatey Upgrade" -Status "Upgrading packages..." -Id 1
            choco upgrade all -y
            Write-Progress -Activity "Chocolatey Upgrade" -Completed -Id 1
            Write-Log "Chocolatey updates applied successfully." -Type Success
        } catch {
            Write-Log "Error upgrading Chocolatey packages: $($_.Exception.Message)" -Type Error
        }
    }

    # 4. Background MS Store Trigger (Native WMI)
    try {
        Write-Log "Triggering background native MS Store update check..." -Type Warning
        $namespaceName = "Root\cimv2\mdm\dmmap"
        $className = "MDM_EnterpriseModernAppManagement_AppManagement01"
        Get-CimInstance -Namespace $namespaceName -ClassName $className -ErrorAction SilentlyContinue | Invoke-CimMethod -MethodName UpdateScanMethod -ErrorAction SilentlyContinue
        Write-Log "Microsoft Store background updates triggered." -Type Success
    } catch {
        Write-Log "Error triggering Microsoft Store updates: $($_.Exception.Message)" -Type Debug
    }
}

Write-Log "Process finished at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Type Info

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
