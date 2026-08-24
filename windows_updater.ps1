# ==============================================================================
# Windows Multi-Manager Update & Inventory Script (Fixed Encoding & Store Filtering)
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
    Write-Host "[$Step/6] $Title" -ForegroundColor Yellow
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
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]$identity
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($Install -and -not $isAdmin) {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        Write-Host " [!] Administrator privileges are required when using -Install." -ForegroundColor Red
        Write-Host " [!] Please re-open PowerShell AS ADMINISTRATOR and run the command again." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Press any key to exit..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        exit
    } else {
        Write-Log "Re-launching script with Administrator privileges..." -Type Warning
        $boundArgs = $PSBoundParameters.Keys | ForEach-Object { "-$_" }
        $argsToPass = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$PSCommandPath`"") + $boundArgs
        Start-Process powershell.exe -Verb RunAs -ArgumentList ($argsToPass -join ' ')
        exit
    }
}

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
    $windowsUpdates = $searchResult.Updates

    if ($windowsUpdates.Count -gt 0) {
        Write-Log "Found $($windowsUpdates.Count) pending Windows update(s):" -Type Success

        foreach ($update in $windowsUpdates) {
            $updateResults.Add([PSCustomObject]@{
                Source         = "Windows Update"
                Title          = $update.Title
                IsInstalled    = $update.IsInstalled
                IsMandatory    = $update.IsMandatory
                RebootRequired = ($update.RebootBehavior -eq 1)
            })
            Write-Host "    * $($update.Title)" -ForegroundColor Green
        }

        if ($Install) {
            foreach ($update in $windowsUpdates) {
                if (-not $update.EulaAccepted) { $update.AcceptEula() }
            }

            Write-Log "Downloading $($windowsUpdates.Count) update(s)..." -Type Warning
            Write-Progress -Activity "Downloading Windows Updates" -Status "Progressing download..." -PercentComplete 50 -ParentId 1 -Id 2
            $downloader = $updateSession.CreateUpdateDownloader()
            $downloader.Updates = $windowsUpdates
            $downloader.Download()
            Write-Progress -Activity "Downloading Windows Updates" -Completed -Id 2

            Write-Log "Installing Windows updates..." -Type Warning
            Write-Progress -Activity "Installing Windows Updates" -Status "Applying updates..." -PercentComplete 75 -ParentId 1 -Id 2
            $installer = $updateSession.CreateUpdateInstaller()
            $installer.Updates = $windowsUpdates
            $installResult = $installer.Install()
            Write-Progress -Activity "Installing Windows Updates" -Completed -Id 2

            Write-Log "Windows Update completed with Result Code: $($installResult.ResultCode)" -Type Success
        }
    } else {
        Write-Log "Windows operating system is up to date." -Type Success
    }
} catch {
    Write-Log "Error processing Windows Updates: $($_.Exception.Message)" -Type Error
}

# ------------------------------------------------------------------------------
# 2. WINGET PACKAGES
# ------------------------------------------------------------------------------
$step = 2
Write-Progress -Activity $activityTitle -Status "Step $step of $totalSteps - Checking WinGet packages..." -PercentComplete (($step / $totalSteps) * 100) -Id 1
Write-SectionHeader -Step $step -Title "WinGet Packages"

if (Get-Command winget -ErrorAction SilentlyContinue) {
    try {
        Write-Log "Querying WinGet package repository..." -Type Info
        $wingetRaw = winget upgrade --include-unknown --accept-source-agreements 2>$null | Out-String
        $wingetLines = $wingetRaw -split "`r?\n" | Where-Object {
            $_ -match '\s{2,}' -and
            $_ -notmatch '^(Name|---|Have you|No updates|Upgrades available|The following)'
        }

        if ($wingetLines) {
            Write-Log "Found $($wingetLines.Count) update(s) via WinGet:" -Type Success

            foreach ($line in $wingetLines) {
                $trimmed = $line.Trim()
                if ($trimmed) {
                    $updateResults.Add([PSCustomObject]@{
                        Source         = "WinGet"
                        Title          = $trimmed
                        IsInstalled    = $true
                        IsMandatory    = $false
                        RebootRequired = $false
                    })
                    Write-Host "    * $trimmed" -ForegroundColor Green
                }
            }

            if ($Install) {
                Write-Log "Upgrading all WinGet packages silently..." -Type Warning
                Write-Progress -Activity "WinGet Upgrade" -Status "Upgrading packages..." -ParentId 1 -Id 2
                winget upgrade --all --silent --accept-source-agreements --accept-package-agreements
                Write-Progress -Activity "WinGet Upgrade" -Completed -Id 2
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
                Write-Log "Found $($validUpdates.Count) Chocolatey update(s):" -Type Success

                foreach ($line in $validUpdates) {
                    $parts = $line.Split('|').Trim()
                    $updateResults.Add([PSCustomObject]@{
                        Source         = "Chocolatey"
                        Title          = "$($parts[0]) ($($parts[1]) to $($parts[2]))"
                        IsInstalled    = $true
                        IsMandatory    = $false
                        RebootRequired = $false
                    })
                    Write-Host "    * $($parts[0]): $($parts[1]) to $($parts[2])" -ForegroundColor Green
                }

                if ($Install) {
                    Write-Log "Upgrading all Chocolatey packages..." -Type Warning
                    Write-Progress -Activity "Chocolatey Upgrade" -Status "Upgrading packages..." -ParentId 1 -Id 2
                    choco upgrade all -y
                    Write-Progress -Activity "Chocolatey Upgrade" -Completed -Id 2
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
    if ($isAdmin) {
        Write-Log "Skipped Scoop: Scoop runs in user context (cannot run as Administrator)." -Type Warning
    } else {
        try {
            if ($Install) {
                Write-Log "Updating Scoop manifests..." -Type Info
                scoop update 2>$null
            }

            Write-Log "Querying Scoop app status..." -Type Info
            $scoopStatus = scoop status 2>$null | Where-Object {
                $_ -match '^\s*\S+' -and $_ -notmatch 'Scoop is up to date' -and $_ -notmatch '---'
            }

            if ($scoopStatus) {
                Write-Log "Found Scoop app updates:" -Type Success

                foreach ($line in $scoopStatus) {
                    $cleanLine = $line.Trim()
                    $updateResults.Add([PSCustomObject]@{
                        Source         = "Scoop"
                        Title          = $cleanLine
                        IsInstalled    = $true
                        IsMandatory    = $false
                        RebootRequired = $false
                    })
                    Write-Host "    * $cleanLine" -ForegroundColor Green
                }

                if ($Install) {
                    Write-Log "Upgrading all Scoop apps..." -Type Warning
                    Write-Progress -Activity "Scoop Upgrade" -Status "Upgrading apps..." -ParentId 1 -Id 2
                    scoop update *
                    Write-Progress -Activity "Scoop Upgrade" -Completed -Id 2
                }
            } else {
                Write-Log "No Scoop app updates available." -Type Success
            }
        } catch {
            Write-Log "Error processing Scoop: $($_.Exception.Message)" -Type Error
        }
    }
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
        $installedApps | Select-Object -First 50 | Format-Table -AutoSize
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

if (Get-Command store -ErrorAction SilentlyContinue) {
    try {
        Write-Log "Querying Microsoft Store updates using 'store' CLI..." -Type Info

        # Query store updates and filter out status output lines
        $storeRaw = store updates 2>$null | Out-String
        $storeLines = $storeRaw -split "`r?\n" | Where-Object {
            $t = $_.Trim()
            $t -ne '' -and
            $t -notmatch 'Checking for updates' -and
            $t -notmatch 'No updates found' -and
            $t -notmatch '^[-=]{2,}$' -and
            $t -notmatch '^(Name|Package|ID)\s+'
        }

        if ($storeLines) {
            Write-Log "Found $($storeLines.Count) Microsoft Store app update(s):" -Type Success

            foreach ($line in $storeLines) {
                $trimmed = $line.Trim()
                $updateResults.Add([PSCustomObject]@{
                    Source         = "MS Store"
                    Title          = $trimmed
                    IsInstalled    = $true
                    IsMandatory    = $false
                    RebootRequired = $false
                })
                Write-Host "    * $trimmed" -ForegroundColor Green
            }

            if ($Install) {
                Write-Log "Applying Microsoft Store updates ('store updates --apply')..." -Type Warning
                Write-Progress -Activity "Store Apps Upgrade" -Status "Applying Store updates..." -ParentId 1 -Id 2
                store updates --apply
                Write-Progress -Activity "Store Apps Upgrade" -Completed -Id 2
                Write-Log "Microsoft Store updates applied successfully." -Type Success
            }
        } else {
            Write-Log "All Microsoft Store apps are up to date." -Type Success
        }
    } catch {
        Write-Log "Error checking Microsoft Store updates: $($_.Exception.Message)" -Type Error
    }
} else {
    Write-Log "'store' utility is not installed or not available in PATH." -Type Debug
}

# Inventory output when running with -Detailed
if ($Detailed) {
    Write-Log "Inventorying installed Microsoft Store packages..." -Type Info
    $storeApps = Get-AppxPackage -ErrorAction SilentlyContinue |
        Where-Object { $_.NonRemovable -ne $true -and $_.SignatureKind -eq 'Store' } |
        Select-Object Name, Version

    if ($storeApps) {
        Write-Host ""
        $storeApps | Select-Object -First 25 | Format-Table -AutoSize
    }
}

# ------------------------------------------------------------------------------
# EXECUTIVE SUMMARY & REPORT EXPORT
# ------------------------------------------------------------------------------
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

Write-Log "Process finished at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Type Info

# PAUSE AT END
Write-Host ""
if (-not $Install) {
 Write-Host "Please, to apply updates rerun with option -Install"
}
Write-Host "Press any key to exit..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
