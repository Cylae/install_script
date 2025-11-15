#Requires -Version 5.1
#Requires -PSEdition Desktop

<#
.SYNOPSIS
    A comprehensive post-installation and optimization script for Windows 11.

.DESCRIPTION
    This script automates the post-installation process for Windows 11 by providing a suite of powerful features, including:
    - A user-friendly graphical interface for selecting and installing a curated list of popular and essential applications.
    - Automated installation of Google Chrome.
    - System optimizations for enhanced performance and privacy, including power plan adjustments and service management.
    - Registry tweaks to improve the user experience and disable telemetry.
    - Removal of bloatware and pre-installed applications to free up system resources.
    - The ability to install local applications from .exe or .msi files.
    - Detailed logging of all actions for easy troubleshooting.

.AUTHOR
    Cylae

.VERSION
    3.1.0
#>

# =====================================================================================================================
# Configuration and Constants
# =====================================================================================================================

$ErrorActionPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"

# Script metadata
$SCRIPT_VERSION = "3.1.0"
$SCRIPT_AUTHOR = "Cylae"
$TARGET_OS = "Windows 11 24H2"
$START_TIME = Get-Date

# Paths
$DESKTOP_DIR = "$env:USERPROFILE\Desktop"
$TEMP_DIR = $env:TEMP
$LOG_DIR = if (Test-Path $DESKTOP_DIR) { $DESKTOP_DIR } else { $TEMP_DIR }
$LOG_FILE = Join-Path $LOG_DIR "log-$($START_TIME.ToString('yyyy-MM-dd_HH-mm-ss')).log"

# =====================================================================================================================
# Utility Functions
# =====================================================================================================================

function Initialize-Script {
    <#
    .SYNOPSIS
        Initializes the script and performs pre-flight checks.
    #>
    Clear-Host

    Write-Host "`n" + ("=" * 80)
    Write-Host "╔════════════════════════════════════════════════════════════════════════════╗"
    Write-Host "║             WINDOWS 11 POST-INSTALLATION SCRIPT v$SCRIPT_VERSION             ║"
    Write-Host "║                  Author: $SCRIPT_AUTHOR | Target OS: $TARGET_OS                  ║"
    Write-Host "╚════════════════════════════════════════════════════════════════════════════╝"
    Write-Host ("=" * 80) + "`n"

    # Start logging
    try {
        Start-Transcript -Path $LOG_FILE -Append -Force | Out-Null
    }
    catch {
        Write-Status -Message "Failed to initialize logging." -Type Error
    }

    # Verify Administrator privileges
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Status -Message "This script requires Administrator privileges. Please re-run PowerShell as an Administrator." -Type Error
        Start-Sleep -Seconds 5
        exit 1
    }

    # Verify Windows 11
    $osVersion = (Get-CimInstance -ClassName Win32_OperatingSystem).Version
    if ($osVersion -notmatch "^10\.0\.(22|23|24|25|26)\d{3}$") {
        Write-Status -Message "This script is optimized for Windows 11. Some features may not work as expected on other versions." -Type Warning
    }

    # Verify STA mode for GUI
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        Write-Status -Message "This script requires Single-Threaded Apartment (STA) mode to display the GUI. Restarting in STA mode..." -Type Warning
        Start-Sleep -Seconds 3
        powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File $PSCommandPath
        exit
    }
}

function Write-Status {
    <#
    .SYNOPSIS
        Writes a formatted status message to the console.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet("Success", "Error", "Warning", "Info", "Step")]
        [string]$Type = "Info"
    )

    $colors = @{
        "Success" = "Green"
        "Error"   = "Red"
        "Warning" = "Yellow"
        "Info"    = "Cyan"
        "Step"    = "Magenta"
    }

    $symbols = @{
        "Success" = "✓"
        "Error"   = "✗"
        "Warning" = "⚠"
        "Info"    = "ℹ"
        "Step"    = "▶"
    }

    Write-Host "$($symbols[$Type]) $Message" -ForegroundColor $colors[$Type]
}

function Ensure-ModuleIsInstalled {
    <#
    .SYNOPSIS
        Ensures that a PowerShell module is installed and imported.
    #>
    param(
        [string]$ModuleName = "Microsoft.WinGet.Client"
    )

    if (Get-Module -Name $ModuleName -ListAvailable) {
        Write-Status -Message "Module '$ModuleName' is already installed." -Type Success
        return
    }

    Write-Status -Message "Installing PowerShell module: $ModuleName" -Type Step

    try {
        # Ensure NuGet provider is installed
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers | Out-Null
        }

        # Trust the PSGallery repository
        if ((Get-PSRepository -Name "PSGallery").InstallationPolicy -ne "Trusted") {
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -Scope AllUsers
        }

        # Install the module
        Install-Module -Name $ModuleName -Force -Scope AllUsers -AllowClobber -ErrorAction Stop | Out-Null
        Import-Module -Name $ModuleName -ErrorAction Stop | Out-Null

        Write-Status -Message "Module '$ModuleName' installed successfully." -Type Success
    }
    catch {
        Write-Status -Message "Failed to install module '$ModuleName'. Please install it manually and re-run the script." -Type Error
        Start-Sleep -Seconds 5
        exit 1
    }
}

function Get-CuratedPackageList {
    <#
    .SYNOPSIS
        Return comprehensive list of 50+ curated software packages by category
    #>
    return @{
        "Essential System" = @(
            "Google.Chrome",
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
        Displays a graphical user interface for selecting applications to install.
    #>
    Write-Status -Message "Launching application selection interface..." -Type Step

    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
    }
    catch {
        Write-Status -Message "Failed to load Windows Forms assemblies. Please ensure that .NET Framework is installed." -Type Error
        return @()
    }

    $packages = Get-CuratedPackageList
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Application Selection"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 40)
    $tabControl.Size = New-Object System.Drawing.Size(760, 450)
    $form.Controls.Add($tabControl)

    $allCheckboxes = @()

    foreach ($category in $packages.Keys) {
        $tabPage = New-Object System.Windows.Forms.TabPage
        $tabPage.Text = $category
        $tabControl.Controls.Add($tabPage)

        $panel = New-Object System.Windows.Forms.FlowLayoutPanel
        $panel.Dock = "Fill"
        $panel.AutoScroll = $true
        $tabPage.Controls.Add($panel)

        foreach ($package in $packages[$category]) {
            $checkbox = New-Object System.Windows.Forms.CheckBox
            $checkbox.Text = $package
            $checkbox.Checked = $true
            $checkbox.Width = 350
            $panel.Controls.Add($checkbox)
            $allCheckboxes += $checkbox
        }
    }

    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = "Search:"
    $searchLabel.Location = New-Object System.Drawing.Point(10, 15)
    $form.Controls.Add($searchLabel)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(70, 12)
    $searchBox.Size = New-Object System.Drawing.Size(200, 20)
    $searchBox.add_TextChanged({
        $searchText = $searchBox.Text.ToLower()
        foreach ($checkbox in $allCheckboxes) {
            $checkbox.Visible = $checkbox.Text.ToLower().Contains($searchText)
        }
    })
    $form.Controls.Add($searchBox)

    $selectAllButton = New-Object System.Windows.Forms.Button
    $selectAllButton.Text = "Select All"
    $selectAllButton.Location = New-Object System.Drawing.Point(550, 12)
    $selectAllButton.add_Click({
        foreach ($checkbox in $allCheckboxes) {
            $checkbox.Checked = $true
        }
    })
    $form.Controls.Add($selectAllButton)

    $deselectAllButton = New-Object System.Windows.Forms.Button
    $deselectAllButton.Text = "Deselect All"
    $deselectAllButton.Location = New-Object System.Drawing.Point(650, 12)
    $deselectAllButton.add_Click({
        foreach ($checkbox in $allCheckboxes) {
            $checkbox.Checked = $false
        }
    })
    $form.Controls.Add($deselectAllButton)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Install"
    $okButton.Location = New-Object System.Drawing.Point(300, 520)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(410, 520)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $allCheckboxes | Where-Object { $_.Checked -and $_.Visible } | ForEach-Object { $_.Text }
    }
    else {
        return @()
    }
}

function Install-WingetPackage {
    <#
    .SYNOPSIS
        Installs a package using WinGet.
    #>
    param(
        [string]$PackageId
    )

    Write-Status -Message "Installing '$PackageId'..." -Type Info

    try {
        $process = Start-Process -FilePath "winget" -ArgumentList "install --id $PackageId --accept-source-agreements --accept-package-agreements" -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop

        if ($process.ExitCode -eq 0) {
            Write-Status -Message "'$PackageId' installed successfully." -Type Success
        }
        elseif ($process.ExitCode -eq -1978335189) {
            Write-Status -Message "'$PackageId' is already installed." -Type Info
        }
        else {
            Write-Status -Message "Failed to install '$PackageId'. Exit code: $($process.ExitCode)" -Type Error
        }
    }
    catch {
        Write-Status -Message "An error occurred while installing '$PackageId'." -Type Error
    }
}

function Optimize-System {
    <#
    .SYNOPSIS
        Applies system optimizations for performance and privacy.
    #>
    Write-Status -Message "Applying system optimizations..." -Type Step

    # Set power plan
    $powerPlan = powercfg -list | Where-Object { $_ -match "Ultimate Performance" }
    if ($powerPlan) {
        $guid = ($powerPlan -split " ")[3]
        powercfg -setactive $guid
        Write-Status -Message "Power plan set to 'Ultimate Performance'." -Type Success
    }
    else {
        $powerPlan = powercfg -list | Where-Object { $_ -match "High Performance" }
        if ($powerPlan) {
            $guid = ($powerPlan -split " ")[3]
            powercfg -setactive $guid
            Write-Status -Message "Power plan set to 'High Performance'." -Type Success
        }
        else {
            Write-Status -Message "Could not find 'Ultimate Performance' or 'High Performance' power plan." -Type Warning
        }
    }

    # Disable sleep/hibernate
    powercfg -change monitor-timeout-ac 0
    powercfg -change disk-timeout-ac 0
    powercfg -change standby-timeout-ac 0
    Write-Status -Message "Disabled sleep and hibernate timeouts." -Type Success

    # Network optimization
    netsh int tcp set global autotuninglevel=normal
    netsh int tcp set global ecn=enabled
    netsh int tcp set global timestamps=disabled
    netsh int tcp set supplemental internet congestionprovider=ctcp
    Write-Status -Message "Network optimizations applied." -Type Success

    # Disable unnecessary services
    $servicesToDisable = @(
        "SysMain", # Superfetch
        "diagtrack" # Connected User Experiences and Telemetry
    )

    foreach ($service in $servicesToDisable) {
        try {
            Set-Service -Name $service -StartupType Disabled -ErrorAction Stop
            Write-Status -Message "Disabled service '$service'." -Type Success
        }
        catch {
            Write-Status -Message "Failed to disable service '$service'." -Type Warning
        }
    }

    # Registry optimizations
    function Set-RegistryValue {
        param(
            [string]$Path,
            [string]$Name,
            $Value,
            [string]$Type = "DWord"
        )
        try {
            if (-not (Test-Path $Path)) {
                New-Item -Path $Path -Force -ErrorAction Stop | Out-Null
            }
            Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force -ErrorAction Stop | Out-Null
            Write-Status -Message "Set registry value '$Name' at '$Path'." -Type Success
        }
        catch {
            Write-Status -Message "Failed to set registry value '$Name' at '$Path'." -Type Warning
        }
    }

    # UI and Responsiveness
    Set-RegistryValue -Path "HKCU:\Control Panel\Desktop" -Name "MenuShowDelay" -Value "0" -Type "String"
    Set-RegistryValue -Path "HKCU:\Control Panel\Desktop" -Name "DragFullWindows" -Value "1" -Type "String"

    # Privacy and Telemetry
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -Value 0
    Set-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Value 0
    Set-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowSyncProviderNotifications" -Value 0
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -Value 1

    # Dark Mode
    Set-RegistryValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "AppsUseLightTheme" 0
    Set-RegistryValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "SystemUsesLightTheme" 0
}

function Enable-SecurityHardening {
    <#
    .SYNOPSIS
        Enable Windows security features and hardening
    #>
    Write-Status "Enabling security hardening features" -Type Step

    Write-Host "  → Configuring Windows Defender..."
    try {
        Set-MpPreference -DisableRealtimeMonitoring $false
        Set-MpPreference -EnableControlledFolderAccess Enabled
        Write-Status "Windows Defender settings applied." -Type Success
    } catch {
        Write-Status "Failed to apply some Windows Defender settings." -Type Warning
    }

    Write-Host "  → Configuring Windows Firewall..."
    try {
        Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True
        Write-Status "Windows Firewall enabled for all profiles." -Type Success
    } catch {
        Write-Status "Failed to enable Windows Firewall." -Type Warning
    }

    Write-Host "  → Ensuring UAC is enabled..."
    try {
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -Value 1
        Write-Status "User Account Control (UAC) is enabled." -Type Success
    } catch {
        Write-Status "Failed to ensure UAC is enabled." -Type Warning
    }
}

function Remove-Bloatware {
    <#
    .SYNOPSIS
        Removes bloatware and pre-installed applications.
    #>
    Write-Status -Message "Removing bloatware..." -Type Step

    $bloatwarePatterns = @(
        "*XboxApp*", "*Xbox.TCUI*", "*XboxGameOverlay*", "*XboxGamingOverlay*",
        "*XboxSpeechToTextOverlay*", "*ZuneVideo*", "*ZuneMusic*", "*WindowsMaps*",
        "*People*", "*YourPhone*", "*MixedReality.Portal*", "*MicrosoftSolitaireCollection*",
        "*MinecraftUWP*", "*Microsoft3DViewer*", "*PaintStudio*", "*SkypeApp*",
        "*GetHelp*", "*Todos*", "*Clipchamp*", "*WeatherApp*", "*CameraApp*"
    )

    $removedCount = 0
    foreach ($pattern in $bloatwarePatterns) {
        try {
            $apps = Get-AppxPackage -AllUsers -Name $pattern -ErrorAction SilentlyContinue
            foreach ($app in $apps) {
                try {
                    Remove-AppxPackage -Package $app.PackageFullName -AllUsers -ErrorAction Stop
                    Write-Status -Message "Removed: $($app.Name)" -Type Success
                    $removedCount++
                } catch {
                    Write-Status -Message "Failed to remove: $($app.Name)" -Type Warning
                }
            }
        } catch {
            # Ignore errors finding packages
        }
    }

    Write-Status -Message "Removed $removedCount bloatware packages." -Type Info

    # Restart Explorer to apply changes
    Write-Status -Message "Restarting Windows Explorer..." -Type Info
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Start-Process explorer.exe -ErrorAction SilentlyContinue
}

function Show-CompletionSummary {
    <#
    .SYNOPSIS
        Displays a summary of the script's execution.
    #>
    $duration = (Get-Date) - $START_TIME

    Write-Host "`n" + ("=" * 80)
    Write-Host "╔════════════════════════════════════════════════════════════════════════════╗"
    Write-Host "║                            SCRIPT EXECUTION SUMMARY                          ║"
    Write-Host "╚════════════════════════════════════════════════════════════════════════════╝"
    Write-Host ("=" * 80) + "`n"

    Write-Status -Message "Execution Time: $($duration.Hours)h $($duration.Minutes)m $($duration.Seconds)s" -Type Info
    Write-Status -Message "Log File: $LOG_FILE" -Type Info

    Write-Host "`n" + ("=" * 80)
    Write-Host "Please restart your computer for all changes to take effect."
    Write-Host ("=" * 80) + "`n"
}

function Schedule-MaintenanceTasks {
    <#
    .SYNOPSIS
        Schedule automated maintenance and system update tasks
    #>
    Write-Status "Scheduling automated maintenance tasks" -Type Step

    Write-Host "  → Creating WinGet auto-update task..."
    try {
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -Command `"winget upgrade --all --accept-package-agreements --accept-source-agreements`""
        $trigger = New-ScheduledTaskTrigger -AtLogon
        Register-ScheduledTask -TaskName "Winget-Auto-Update" -Action $action -Trigger $trigger -Force -ErrorAction Stop | Out-Null
        Write-Status "WinGet auto-update task created successfully." -Type Success
    } catch {
        Write-Status "Failed to create WinGet auto-update task." -Type Warning
    }
}

# =====================================================================================================================
# Main Execution
# =====================================================================================================================

Initialize-Script

try {
    # Phase 1: Prerequisites
    Write-Status -Message "PHASE 1: System Checks and Prerequisites" -Type Step
    Ensure-ModuleIsInstalled

    # Phase 2: Application Installation
    Write-Status -Message "PHASE 2: Application Installation" -Type Step
    $selectedPackages = Select-Applications
    if ($selectedPackages) {
        foreach ($package in $selectedPackages) {
            Install-WingetPackage -PackageId $package
        }
    }
    else {
        Write-Status -Message "No applications selected. Skipping installation." -Type Info
    }

    # Phase 3: System Optimization & Hardening
    Write-Status -Message "PHASE 3: System Optimization & Hardening" -Type Step
    Optimize-System
    Enable-SecurityHardening

    # Phase 6: Cleanup
    Write-Status -Message "PHASE 6: Cleanup" -Type Step
    Remove-Bloatware

    # Phase 7: Maintenance
    Write-Status -Message "PHASE 7: Maintenance" -Type Step
    Schedule-MaintenanceTasks

    # Phase 8: Completion Summary
    Write-Status -Message "PHASE 8: Completion Summary" -Type Step
    Show-CompletionSummary
}
catch {
    Write-Status -Message "An unexpected error occurred. Please check the log file for more details." -Type Error
}
finally {
    Write-Status -Message "Script execution finished. Press Enter to exit." -Type Info
    Read-Host | Out-Null
    Stop-Transcript | Out-Null
}
