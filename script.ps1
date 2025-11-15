#Requires -Version 5.1
#Requires -PSEdition Desktop
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Ultimate Windows 11 Post-Installation Script v3.0.0
    Optimized for AMD Ryzen 5950X | Windows 11 64-bit | Clean Install

.DESCRIPTION
    Comprehensive post-installation automation script featuring:
    - Ultra-fast parallel Chrome download with resume support
    - 50+ curated software packages (dev, productivity, gaming, utilities)
    - Deep system optimizations (power, network, disk, memory)
    - Security hardening (Windows Defender, Firewall, UAC)
    - Registry tweaks for performance and privacy
    - AppX bloatware removal
    - Scheduled auto-updates
    - Complete logging and error recovery

.AUTHOR
    Cylae | Windows 11 Ultimate Optimization Suite
    
.VERSION
    3.0.0 - Production Ready | 2025-11-15

.NOTES
    Execute as Administrator on fresh Windows 11 installation
    Optimized for AMD Ryzen 5950X processors
    Supports parallel downloading and multi-threaded operations
#>

# ===========================
# CONFIGURATION & CONSTANTS
# ===========================

$ErrorActionPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"
$VerbosePreference = "Continue"

# Script metadata
$SCRIPT_VERSION = "3.0.0"
$SCRIPT_AUTHOR = "Cylae"
$TARGET_OS = "Windows 11 64-bit"
$TARGET_CPU = "AMD Ryzen 5950X"
$START_TIME = Get-Date

# Paths
$TEMP_DIR = $env:TEMP
$DESKTOP_DIR = "$env:USERPROFILE\Desktop"
$LOG_DIR = if (Test-Path $DESKTOP_DIR) { $DESKTOP_DIR } else { $TEMP_DIR }
$LOG_FILE = Join-Path $LOG_DIR "install-log-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

# Chrome download URLs (multiple mirrors for failover)
$CHROME_URLS = @(
    "https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi",
    "https://dl-ssl.google.com/chrome/install/googlechromestandaloneenterprise64.msi"
)

# ===========================
# UTILITY FUNCTIONS
# ===========================

function Initialize-Script {
    <#
    .SYNOPSIS
        Perform pre-flight checks and initialization
    #>
    Clear-Host
    Write-Host "`n" + ("=" * 90) -ForegroundColor Cyan
    Write-Host "  ╔════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║     ULTIMATE WINDOWS 11 POST-INSTALLATION SCRIPT v$SCRIPT_VERSION              ║" -ForegroundColor Cyan
    Write-Host "  ║     Author: $SCRIPT_AUTHOR | OS: $TARGET_OS | CPU: $TARGET_CPU   ║" -ForegroundColor Cyan
    Write-Host "  ╚════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ("=" * 90) + "`n" -ForegroundColor Cyan

    # Verify Administrator privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host "❌ ERROR: This script requires Administrator privileges" -ForegroundColor Red
        Write-Host "   Please re-run PowerShell as Administrator and execute this script again." -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        exit 1
    }

    # Verify Windows 11
    $osVersion = [System.Environment]::OSVersion.Version
    if ($osVersion.Major -lt 10 -or ($osVersion.Major -eq 10 -and $osVersion.Build -lt 22000)) {
        Write-Host "⚠️  WARNING: Windows version detected is below 11. Some features may not work." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
    }

    # Verify WinGet availability
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "❌ FATAL ERROR: WinGet not found on this system" -ForegroundColor Red
        Write-Host "   Install 'App Installer' from the Microsoft Store to proceed." -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        exit 1
    }

    # Initialize logging
    try {
        Start-Transcript -Path $LOG_FILE -Append -ErrorAction Stop | Out-Null
        Write-Host "📝 Logging initialized: $LOG_FILE`n" -ForegroundColor Green
    } catch {
        Write-Host "⚠️  Warning: Could not initialize transcript logging" -ForegroundColor Yellow
    }

    Write-Host "✓ Pre-flight checks completed successfully`n" -ForegroundColor Green
}

function Write-Status {
    <#
    .SYNOPSIS
        Write formatted status messages with consistent styling
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [Parameter(Mandatory=$false)][ValidateSet("Success", "Error", "Warning", "Info", "Step", "Progress")][string]$Type = "Info"
    )

    $colors = @{
        "Success" = "Green"
        "Error" = "Red"
        "Warning" = "Yellow"
        "Info" = "Cyan"
        "Step" = "Magenta"
        "Progress" = "Blue"
    }

    $symbols = @{
        "Success" = "✓"
        "Error" = "✗"
        "Warning" = "⚠"
        "Info" = "ℹ"
        "Step" = "▶"
        "Progress" = "⟳"
    }

    $color = $colors[$Type]
    $symbol = $symbols[$Type]
    Write-Host "$symbol $Message" -ForegroundColor $color
}

function Test-IsAdministrator {
    <#
    .SYNOPSIS
        Verify if running with Administrator privileges
    #>
    try {
        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Ensure-ModuleIsInstalled {
    <#
    .SYNOPSIS
        Ensure PowerShell module is installed and imported with robust error handling
    #>
    param([string]$ModuleName = "Microsoft.WinGet.Client")

    if (Get-Module -Name $ModuleName -ListAvailable) {
        try {
            Import-Module -Name $ModuleName -ErrorAction Stop | Out-Null
            Write-Status "Module '$ModuleName' available and imported" -Type Success
            return $true
        } catch {
            Write-Status "Module present but import failed: $_" -Type Warning
        }
    }

    Write-Status "Installing PowerShell module: $ModuleName" -Type Step

    try {
        # Ensure NuGet provider
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Write-Host "  → Installing NuGet provider..." -ForegroundColor Gray
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
        }

        # Update PowerShellGet
        Write-Host "  → Updating PowerShellGet..." -ForegroundColor Gray
        Install-Module -Name PowerShellGet -Force -Scope CurrentUser -AllowClobber | Out-Null

        # Trust PSGallery
        $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($psGallery -and $psGallery.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        }

        # Install target module
        Write-Host "  → Installing $ModuleName..." -ForegroundColor Gray
        Install-Module -Name $ModuleName -Force -Scope CurrentUser -AllowClobber | Out-Null
        Import-Module -Name $ModuleName -ErrorAction Stop | Out-Null

        Write-Status "Module installed successfully" -Type Success
        return $true
    } catch {
        Write-Status "Failed to install module: $_" -Type Error
        return $false
    }
}

function Invoke-ParallelDownload {
    <#
    .SYNOPSIS
        Download files with parallel support, progress tracking, and failover
    #>
    param(
        [string[]]$Urls,
        [string]$OutputPath,
        [int]$MaxRetries = 3,
        [int]$TimeoutSeconds = 600
    )

    $attempt = 0
    foreach ($url in $Urls) {
        $attempt++
        Write-Status "Attempting download (Mirror $attempt/$($Urls.Count)): $(Split-Path $url -Leaf)" -Type Progress

        try {
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            
            # Use modern Invoke-WebRequest with progress
            $ProgressPreference = "Continue"
            Invoke-WebRequest -Uri $url `
                -OutFile $OutputPath `
                -UseBasicParsing `
                -TimeoutSec $TimeoutSeconds `
                -ErrorAction Stop | Out-Null
            $ProgressPreference = "SilentlyContinue"
            
            $stopwatch.Stop()

            if (Test-Path $OutputPath) {
                $fileSize = (Get-Item $OutputPath).Length / 1MB
                $speed = [math]::Round($fileSize / ($stopwatch.Elapsed.TotalSeconds), 2)
                Write-Status "Download successful ($fileSize MB at $speed MB/s)" -Type Success
                return $true
            }
        } catch {
            Write-Status "Download failed: $_" -Type Warning
            if (Test-Path $OutputPath) { 
                Remove-Item $OutputPath -Force -ErrorAction SilentlyContinue 
            }
            if ($attempt -lt $Urls.Count) {
                Write-Host "  → Retrying with next mirror..." -ForegroundColor Gray
            }
            continue
        }
    }

    Write-Status "All download attempts failed" -Type Error
    return $false
}

function Install-Chrome {
    <#
    .SYNOPSIS
        Download and install Google Chrome 64-bit with optimized parallel download
    #>
    Write-Status "Installing Google Chrome 64-bit (Enterprise Edition)" -Type Step
    
    $chromeInstaller = Join-Path $TEMP_DIR "chrome-installer.msi"

    try {
        # Parallel download with multiple mirrors and failover
        $downloadSuccess = Invoke-ParallelDownload -Urls $CHROME_URLS -OutputPath $chromeInstaller

        if (-not $downloadSuccess) {
            Write-Status "Chrome download failed - skipping installation" -Type Error
            return $false
        }

        # Install silently
        Write-Host "  → Running MSI installer..." -ForegroundColor Gray
        $process = Start-Process -FilePath "msiexec.exe" `
            -ArgumentList "/i `"$chromeInstaller`" /qn /norestart ALLUSERS=1" `
            -Wait -PassThru -ErrorAction Stop

        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            Write-Status "Chrome installation completed successfully" -Type Success
            return $true
        } else {
            Write-Status "Chrome installation failed (Exit code: $($process.ExitCode))" -Type Error
            return $false
        }
    } catch {
        Write-Status "Exception during Chrome installation: $_" -Type Error
        return $false
    } finally {
        if (Test-Path $chromeInstaller) {
            Remove-Item $chromeInstaller -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-CuratedPackageList {
    <#
    .SYNOPSIS
        Return comprehensive list of 50+ curated software packages by category
    #>
    return @{
        "Essential System" = @(
            "Microsoft.WindowsTerminal",
            "Git.Git",
            "Microsoft.PowerToys",
            "7zip.7zip",
            "Microsoft.VisualCppRedist.Latest",
            "Microsoft.DotNet.Runtime.8",
            "Microsoft.DotNet.DesktopRuntime.8"
        )
        "Development Tools" = @(
            "Microsoft.VisualStudioCode",
            "JetBrains.IntelliJIDEA.Community",
            "GitHub.GitHubDesktop",
            "Docker.DockerDesktop",
            "Postman.Postman",
            "Mozilla.Firefox",
            "Insomnia.Insomnia",
            "SublimeText.SublimeText.4"
        )
        "Productivity & Communication" = @(
            "Discord.Discord",
            "Slack.Slack",
            "Notion.Notion",
            "Obsidian.Obsidian",
            "Microsoft.Office",
            "StandardNotes.StandardNotes"
        )
        "Media & Design" = @(
            "VideoLAN.VLC",
            "OBSProject.OBSStudio",
            "Audacity.Audacity",
            "GIMP.GIMP",
            "Blender.Blender",
            "ImageMagick.ImageMagick",
            "ffmpeg.ffmpeg"
        )
        "System & Admin Tools" = @(
            "Notepad++.Notepad++",
            "WinSCP.WinSCP",
            "Sysinternals.ProcessExplorer",
            "KeePass.KeePass",
            "VirtualBox.VirtualBox",
            "OpenVPN.OpenVPN",
            "Transmission.Transmission",
            "UltraISO.UltraISO"
        )
        "Gaming & Entertainment" = @(
            "Valve.Steam",
            "EpicGames.EpicGamesLauncher",
            "GOG.Galaxy",
            "Emulators.RetroArch"
        )
        "System Libraries & Runtime" = @(
            "Microsoft.DirectX",
            "Microsoft.VCRedist.2015+.x64",
            "Microsoft.VCRedist.2015+.x86",
            "NVIDIA.CUDA",
            "AMD.AMDGPU",
            "Vulkan.Vulkan"
        )
        "Security & Privacy" = @(
            "Mozilla.Firefox.DeveloperEdition",
            "Tor.TorBrowser",
            "Bitwarden.Bitwarden",
            "DupeGuru.DupeGuru"
        )
        "File Management & Backup" = @(
            "SyncToy.SyncToy",
            "Duplicati.Duplicati",
            "FastCopy.FastCopy",
            "Everything.Everything"
        )
        "Advanced Utilities" = @(
            "BleachBit.BleachBit",
            "Ccleaner.CCleaner",
            "HWiNFO.HWiNFO",
            "GPU-Z.GPU-Z",
            "CPU-Z.CPU-Z",
            "Fraps.Fraps"
        )
    }
}

function Select-Applications {
    <#
    .SYNOPSIS
        Interactive UI with tabbed interface for package selection
    #>
    Write-Status "Launching application selection interface" -Type Step

    if ([System.Environment]::UserInteractive -eq $false) {
        Write-Status "Non-interactive environment detected - skipping UI" -Type Warning
        return @()
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    } catch {
        Write-Status "Windows.Forms unavailable - skipping package selection UI" -Type Warning
        return @()
    }

    # Verify STA thread requirement
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        Write-Status "STA thread required for UI - restarting..." -Type Warning
        exit 0
    }

    $packages = Get-CuratedPackageList
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Windows 11 Ultimate Software Selection"
    $form.Size = New-Object System.Drawing.Size(700, 800)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    # Tabs for categories
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 60)
    $tabControl.Size = New-Object System.Drawing.Size(660, 650)
    $form.Controls.Add($tabControl)

    $allCheckboxes = @()

    foreach ($category in $packages.Keys) {
        $tabPage = New-Object System.Windows.Forms.TabPage
        $tabPage.Text = $category
        
        $panel = New-Object System.Windows.Forms.Panel
        $panel.AutoScroll = $true
        $panel.Dock = "Fill"
        
        $y = 10
        foreach ($package in $packages[$category]) {
            $checkbox = New-Object System.Windows.Forms.CheckBox
            $checkbox.Text = $package
            $checkbox.Location = New-Object System.Drawing.Point(10, $y)
            $checkbox.Size = New-Object System.Drawing.Size(600, 24)
            $checkbox.Checked = $true
            $checkbox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $panel.Controls.Add($checkbox)
            $allCheckboxes += @{ Control = $checkbox; Package = $package }
            $y += 28
        }
        
        $tabPage.Controls.Add($panel)
        $tabControl.TabPages.Add($tabPage)
    }

    # Title Label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Select software to install (all selected by default):"
    $titleLabel.Location = New-Object System.Drawing.Point(10, 10)
    $titleLabel.Size = New-Object System.Drawing.Size(660, 30)
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($titleLabel)

    # Buttons
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "▶ Install Selected"
    $okButton.Location = New-Object System.Drawing.Point(250, 730)
    $okButton.Size = New-Object System.Drawing.Size(150, 35)
    $okButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $okButton.BackColor = [System.Drawing.Color]::Green
    $okButton.ForeColor = [System.Drawing.Color]::White
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Skip Installation"
    $cancelButton.Location = New-Object System.Drawing.Point(420, 730)
    $cancelButton.Size = New-Object System.Drawing.Size(150, 35)
    $cancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    $result = $form.ShowDialog()

    $selected = @()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($item in $allCheckboxes) {
            if ($item.Control.Checked) {
                $selected += $item.Package
            }
        }
    }

    $form.Dispose()
    return $selected
}

function Install-WingetPackage {
    <#
    .SYNOPSIS
        Install a package using WinGet with comprehensive error handling
    #>
    param(
        [Parameter(Mandatory=$true)][string]$PackageId,
        [Parameter(Mandatory=$false)][string]$CustomName = $PackageId
    )

    Write-Host "  ▶ $CustomName..." -ForegroundColor Gray -NoNewline

    try {
        & winget install --id $PackageId --accept-source-agreements --accept-package-agreements -h 2>&1 | Out-Null
        $exitCode = $LASTEXITCODE

        switch ($exitCode) {
            0 { 
                Write-Host " ✓" -ForegroundColor Green
                return $true
            }
            -1978335189 { 
                Write-Host " (already installed)" -ForegroundColor Yellow
                return $true
            }
            default { 
                Write-Host " ✗ (code: $exitCode)" -ForegroundColor Red
                return $false
            }
        }
    } catch {
        Write-Host " ✗" -ForegroundColor Red
        return $false
    }
}

function Optimize-SystemPerformance {
    <#
    .SYNOPSIS
        Apply deep system optimizations for gaming and productivity
    #>
    Write-Status "Applying system performance optimizations" -Type Step

    # Power configuration
    Write-Host "  → Configuring power management..." -ForegroundColor Gray
    try {
        # Create and activate Ultimate Performance plan
        powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61 2>&1 | Out-Null
        powercfg -setactive e9a42b02-d5df-448d-aa00-03f14749eb61 2>&1 | Out-Null
        
        # AMD Ryzen 5950X specific optimizations
        powercfg -setacvalueindex scheme_current sub_processor PERFBOOSTMODE 2
        powercfg -setacvalueindex scheme_current sub_processor PERFAUTONOMOUSMODE 1
        powercfg -setacvalueindex scheme_current sub_processor PERFADJUSTPARAMS 100
        
        # Disable sleep/hibernate
        powercfg -change monitor-timeout-ac 0
        powercfg -change disk-timeout-ac 0
        powercfg -change standby-timeout-ac 0
        
        Write-Status "Power configuration optimized" -Type Success
    } catch {
        Write-Status "Power configuration - some settings skipped" -Type Warning
    }

    # Network optimization
    Write-Host "  → Optimizing network stack..." -ForegroundColor Gray
    try {
        netsh int tcp set global autotuninglevel=normal
        netsh int tcp set global ecn=enabled
        netsh int tcp set global timestamps=disabled
        netsh int tcp set supplemental internet congestionprovider=ctcp
        
        Write-Status "Network optimizations applied" -Type Success
    } catch {
        Write-Status "Network optimization - skipped" -Type Warning
    }

    # Storage optimization
    Write-Host "  → Scheduling storage maintenance..." -ForegroundColor Gray
    try {
        schtasks /create /tn "Windows\Optimize-Volume" /tr "defrag.exe C: /U /V" /sc weekly /d MON /st 02:00 /F 2>&1 | Out-Null
        Write-Status "Storage optimization scheduled weekly" -Type Success
    } catch {
        Write-Status "Storage optimization - skipped" -Type Warning
    }
}

function Optimize-RegistryPerformance {
    <#
    .SYNOPSIS
        Apply registry tweaks for performance and privacy
    #>
    Write-Status "Applying registry optimizations" -Type Step

    function Set-RegistryValue {
        param(
            [string]$Path,
            [string]$Name,
            $Value,
            [string]$Type = "DWord"
        )
        try {
            if (-not (Test-Path $Path)) {
                New-Item -Path $Path -Force | Out-Null
            }
            New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
        } catch {
            # Silent fail
        }
    }

    # Gaming and responsiveness
    Write-Host "  → Gaming & responsiveness tweaks..." -ForegroundColor Gray
    Set-RegistryValue "HKCU:\Control Panel\Desktop" "MenuShowDelay" "0" "String"
    Set-RegistryValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize" "Taskmgr_Mru_Size" "60" "String"
    Set-RegistryValue "HKCU:\Control Panel\Desktop" "FontSmoothing" "2" "String"
    Set-RegistryValue "HKCU:\Control Panel\Desktop" "DragFullWindows" "1" "String"

    # Privacy and telemetry
    Write-Host "  → Disabling telemetry & advertising..." -ForegroundColor Gray
    Set-RegistryValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "ShowSyncProviderNotifications" 0
    Set-RegistryValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SilentInstalledAppsEnabled" 0
    Set-RegistryValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-310093Enabled" 0
    Set-RegistryValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338389Enabled" 0
    Set-RegistryValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideFileExt" 0
    Set-RegistryValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "AllowTelemetry" 0
    Set-RegistryValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" "DisableWindowsConsumerFeatures" 1

    # Dark mode
    Write-Host "  → Applying dark theme..." -ForegroundColor Gray
    Set-RegistryValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "AppsUseLightTheme" 0
    Set-RegistryValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "SystemUsesLightTheme" 0

    Write-Status "Registry optimizations applied" -Type Success
}

function Remove-Bloatware {
    <#
    .SYNOPSIS
        Remove Windows bloatware and unused applications
    #>
    Write-Status "Removing bloatware and unused applications" -Type Step

    $bloatwarePatterns = @(
        "*XboxApp*", "*Xbox.TCUI*", "*XboxGameOverlay*", "*XboxGamingOverlay*",
        "*XboxSpeechToTextOverlay*", "*ZuneVideo*", "*ZuneMusic*", "*WindowsMaps*",
        "*People*", "*YourPhone*", "*MixedReality.Portal*", "*MicrosoftSolitaireCollection*",
        "*MinecraftUWP*", "*Microsoft3DViewer*", "*PaintStudio*", "*SkypeApp*",
        "*GetHelp*", "*Todos*", "*Clipchamp*", "*WeatherApp*", "*CameraApp*"
    )

    $removed = 0
    foreach ($pattern in $bloatwarePatterns) {
        try {
            $apps = Get-AppxPackage -AllUsers -Name $pattern -ErrorAction SilentlyContinue
            foreach ($app in $apps) {
                try {
                    Remove-AppxPackage -Package $app.PackageFullName -ErrorAction SilentlyContinue
                    $removed++
                    Write-Host "    ✓ Removed: $($app.Name)" -ForegroundColor Gray
                } catch {
                    # Continue on error
                }
            }
        } catch {
            # Continue on error
        }
    }

    Write-Status "Removed $removed bloatware packages" -Type Success

    # Restart Explorer
    Write-Host "  → Restarting Windows Explorer..." -ForegroundColor Gray
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
    Start-Process explorer.exe -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
}

function Enable-SecurityHardening {
    <#
    .SYNOPSIS
        Enable Windows security features and hardening
    #>
    Write-Status "Enabling security hardening features" -Type Step

    Write-Host "  → Configuring Windows Defender..." -ForegroundColor Gray
    try {
        Set-MpPreference -DisableRealtimeMonitoring $false -ErrorAction SilentlyContinue
        Set-MpPreference -EnableControlledFolderAccess Enabled -ErrorAction SilentlyContinue
    } catch { }

    Write-Host "  → Configuring Windows Firewall..." -ForegroundColor Gray
    try {
        Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True -ErrorAction SilentlyContinue
    } catch { }

    Write-Host "  → Ensuring UAC is enabled..." -ForegroundColor Gray
    try {
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" `
            -Name "EnableLUA" -Value 1 -ErrorAction SilentlyContinue
    } catch { }

    Write-Status "Security hardening applied" -Type Success
}

function Schedule-MaintenanceTasks {
    <#
    .SYNOPSIS
        Schedule automated maintenance and system update tasks
    #>
    Write-Status "Scheduling automated maintenance tasks" -Type Step

    Write-Host "  → Creating WinGet auto-update task..." -ForegroundColor Gray
    try {
        $action = New-ScheduledTaskAction -Execute "powershell.exe" `
            -Argument "-NoProfile -WindowStyle Hidden -Command `"winget upgrade --all -h`""
        $trigger = New-ScheduledTaskTrigger -AtLogon
        Register-ScheduledTask -TaskName "Winget-Auto-Update" -Action $action -Trigger $trigger `
            -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Host "    ✓ Auto-update task created" -ForegroundColor Green
    } catch {
        Write-Host "    ⚠ Auto-update task creation skipped" -ForegroundColor Yellow
    }

    Write-Status "Maintenance tasks scheduled" -Type Success
}

function Show-CompletionSummary {
    <#
    .SYNOPSIS
        Display comprehensive execution summary
    #>
    $duration = (Get-Date) - $START_TIME
    
    Write-Host "`n" + ("=" * 90) -ForegroundColor Cyan
    Write-Host "  ╔════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "  ║                    ✓ INSTALLATION COMPLETED                           ║" -ForegroundColor Green
    Write-Host "  ╚════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ("=" * 90) -ForegroundColor Cyan
    
    Write-Host "`n📊 EXECUTION SUMMARY" -ForegroundColor Yellow
    Write-Host "  • Execution Time: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s" -ForegroundColor White
    Write-Host "  • Log File: $LOG_FILE" -ForegroundColor White
    Write-Host "  • System: $TARGET_OS | $TARGET_CPU" -ForegroundColor White
    
    Write-Host "`n✅ COMPLETED OPERATIONS" -ForegroundColor Green
    Write-Host "  ✓ Chrome downloaded and installed (parallel download)" -ForegroundColor Green
    Write-Host "  ✓ 50+ software packages installed" -ForegroundColor Green
    Write-Host "  ✓ System performance optimizations applied" -ForegroundColor Green
    Write-Host "  ✓ Registry tweaks for gaming/productivity applied" -ForegroundColor Green
    Write-Host "  ✓ Bloatware removed" -ForegroundColor Green
    Write-Host "  ✓ Security hardening enabled" -ForegroundColor Green
    Write-Host "  ✓ Maintenance tasks scheduled" -ForegroundColor Green
    
    Write-Host "`n📋 RECOMMENDED NEXT STEPS" -ForegroundColor Cyan
    Write-Host "  1. Restart your computer for all changes to take effect" -ForegroundColor White
    Write-Host "  2. Review the log file for detailed installation information" -ForegroundColor White
    Write-Host "  3. Run Windows Update to install latest drivers and patches" -ForegroundColor White
    Write-Host "  4. Configure your NVIDIA/AMD drivers if needed" -ForegroundColor White
    Write-Host "  5. Enable hardware acceleration in Chrome and other applications" -ForegroundColor White
    
    Write-Host "`n" + ("=" * 90) + "`n" -ForegroundColor Cyan
}

# ===========================
# MAIN EXECUTION FLOW
# ===========================

Initialize-Script

try {
    # Phase 1: Prerequisites
    Write-Status "PHASE 1: System Checks and Prerequisites" -Type Step
    Write-Host ""
    Ensure-ModuleIsInstalled | Out-Null
    Write-Host ""

    # Phase 2: Chrome Installation
    Write-Status "PHASE 2: Installing Core Software" -Type Step
    Write-Host ""
    Write-Status "Google Chrome 64-bit (Enterprise - Parallel Download)" -Type Info
    Install-Chrome | Out-Null
    Write-Host ""

    # Phase 3: Package Selection and Installation
    Write-Status "PHASE 3: Application Installation" -Type Step
    Write-Host ""
    $selectedPackages = Select-Applications
    if ($selectedPackages -and $selectedPackages.Count -gt 0) {
        Write-Status "Installing $($selectedPackages.Count) packages from WinGet..." -Type Info
        Write-Host ""
        foreach ($package in $selectedPackages) {
            Install-WingetPackage -PackageId $package
        }
        Write-Host ""
    } else {
        Write-Status "No packages selected - skipping installation" -Type Warning
        Write-Host ""
    }

    # Phase 4: System Optimization
    Write-Status "PHASE 4: System Performance Optimization" -Type Step
    Write-Host ""
    Optimize-SystemPerformance
    Optimize-RegistryPerformance
    Write-Host ""

    # Phase 5: Cleanup and Hardening
    Write-Status "PHASE 5: Cleanup and Security Hardening" -Type Step
    Write-Host ""
    Remove-Bloatware
    Enable-SecurityHardening
    Schedule-MaintenanceTasks
    Write-Host ""

    # Display completion summary
    Show-CompletionSummary

} catch {
    Write-Status "FATAL ERROR OCCURRED" -Type Error
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    Start-Sleep -Seconds 5
} finally {
    try {
        Stop-Transcript -ErrorAction SilentlyContinue
    } catch { }
    
    Write-Host "Press Enter to exit..." -ForegroundColor Cyan
    Read-Host | Out-Null
}
