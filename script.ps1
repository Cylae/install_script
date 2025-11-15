#Requires -Version 5.1
#Requires -PSEdition Desktop

<#
.SYNOPSIS
    A comprehensive post-installation and optimization script for Windows.

.DESCRIPTION
    This script automates the post-installation process for Windows by providing a suite of powerful features, including:
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

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# Script metadata
$SCRIPT_VERSION = "3.1.0"
$SCRIPT_AUTHOR = "Cylae"
$TARGET_OS = "Windows 10 & 11"
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
    Write-Host "║              WINDOWS POST-INSTALLATION SCRIPT v$SCRIPT_VERSION              ║"
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

    # Verify Windows version
    $osVersion = [System.Environment]::OSVersion.Version
    if ($osVersion.Major -ne 10 -or $osVersion.Build -lt 19041) {
        Write-Status -Message "This script is designed for Windows 10 (version 2004+) and Windows 11. Running on an unsupported OS version may lead to unexpected results." -Type Warning
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
        Write-Status "Module '$ModuleName' is already installed." -Type Success
        return
    }

    Write-Status "Installing PowerShell module: $ModuleName..." -Type Step

    try {
        Write-Host "  → Ensuring PowerShellGet is up-to-date..."
        Install-Module -Name PowerShellGet -Force -Scope AllUsers -AllowClobber -ErrorAction Stop | Out-Null

        Write-Host "  → Ensuring NuGet provider is installed..."
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction Stop | Out-Null

        Write-Host "  → Trusting the PSGallery repository..."
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted

        Write-Host "  → Installing module '$ModuleName'..."
        Install-Module -Name $ModuleName -Force -Scope AllUsers -AllowClobber -ErrorAction Stop | Out-Null

        Write-Status "Module '$ModuleName' installed and imported successfully." -Type Success
    }
    catch {
        Write-Status "An error occurred during the installation of module '$ModuleName'." -Type Error
        Write-Status "Please try to install it manually by running 'Install-Module -Name $ModuleName -Force' in an Administrator PowerShell window and then re-run the script." -Type Info
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
            "Piriform.CCleaner",
            "HWiNFO.HWiNFO",
            "TechPowerUp.GPU-Z",
            "CPUID.CPU-Z",
            "Fraps.Fraps"
        )
        "Local Installers" = @(
            "Google Chrome",
            "NVIDIA App"
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

    $allCheckboxes = @()

    # Main panel for checkboxes
    $mainPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $mainPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainPanel.AutoScroll = $true
    $mainPanel.Padding = New-Object System.Windows.Forms.Padding(10)

    # Top panel for controls
    $topPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $topPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $topPanel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 0)
    $topPanel.Height = 40

    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = "Search:"
    $searchLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Size = New-Object System.Drawing.Size(200, 20)
    $searchBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $searchBox.add_TextChanged({
        $searchText = $searchBox.Text.ToLower()
        $mainPanel.SuspendLayout() # Pause layout updates for performance
        foreach ($checkbox in $allCheckboxes) {
            $checkbox.Visible = $checkbox.Text.ToLower().Contains($searchText)
        }
        $mainPanel.ResumeLayout() # Resume layout updates
    })

    $selectAllButton = New-Object System.Windows.Forms.Button
    $selectAllButton.Text = "Select All"
    $selectAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $selectAllButton.add_Click({ foreach ($checkbox in $allCheckboxes) { $checkbox.Checked = $true } })

    $deselectAllButton = New-Object System.Windows.Forms.Button
    $deselectAllButton.Text = "Deselect All"
    $deselectAllButton.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $deselectAllButton.add_Click({ foreach ($checkbox in $allCheckboxes) { $checkbox.Checked = $false } })

    $topPanel.Controls.AddRange(@($searchLabel, $searchBox, $selectAllButton, $deselectAllButton))

    # Bottom panel for buttons
    $bottomPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $bottomPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $bottomPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $bottomPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    $bottomPanel.Height = 50

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Install"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton

    $bottomPanel.Controls.AddRange(@($cancelButton, $okButton))

    foreach ($category in $packages.Keys.GetEnumerator() | Sort-Object) {
        $categoryLabel = New-Object System.Windows.Forms.Label
        $categoryLabel.Text = $category
        $categoryLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        $categoryLabel.AutoSize = $true
        $categoryLabel.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 5)
        $mainPanel.SetFlowBreak($categoryLabel, $true)
        $mainPanel.Controls.Add($categoryLabel)

        foreach ($package in $packages[$category]) {
            $checkbox = New-Object System.Windows.Forms.CheckBox
            $checkbox.Text = $package
            $checkbox.Checked = $true
            $checkbox.Width = 350
            $checkbox.Margin = New-Object System.Windows.Forms.Padding(20, 0, 0, 0)
            $mainPanel.Controls.Add($checkbox)
            $allCheckboxes += $checkbox
        }
    }
    # Add all panels to the form
    $form.Controls.AddRange(@($mainPanel, $topPanel, $bottomPanel))

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedPackages = $allCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Text }
        return $selectedPackages
    }
    else {
        return @()
    }
}

function Install-WingetPackage {
    <#
    .SYNOPSIS
        Installs a package using the WinGet PowerShell module.
    #>
    param(
        [string]$PackageId
    )

    Write-Status "Installing '$PackageId'..." -Type Info
    try {
        $package = Find-WinGetPackage -Id $PackageId -ErrorAction Stop
        if ($package) {
            Install-WinGetPackage -Id $PackageId -AcceptPackageAgreements -AcceptSourceAgreements -ErrorAction Stop | Out-Null
            Write-Status "'$PackageId' installed successfully." -Type Success
        }
        else {
            Write-Status "Package '$PackageId' not found." -Type Warning
        }
    }
    catch {
        if ($_.Exception.Message -like "*0x80070057*") { # Package already installed
            Write-Status "'$PackageId' is already installed." -Type Info
        } else {
            Write-Status "Failed to install '$PackageId'." -Type Error
            Write-Status "  Reason: $($_.Exception.Message)" -Type Info
        }
    }
}

function Install-LocalPackage {
    <#
    .SYNOPSIS
        Installs a local package from an .exe or .msi file.
    #>
    param(
        [string]$AppName
    )

    Write-Status "Prompting for installer: $AppName" -Type Step
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Please select the installer for $AppName"
    $openFileDialog.Filter = "Installers (*.exe, *.msi)|*.exe;*.msi"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $installerPath = $openFileDialog.FileName
        Write-Status "Attempting to install '$AppName' from '$installerPath'..." -Type Info

        try {
            # Attempt silent installation
            $arguments = @("/i", "`"$installerPath`"", "/qn", "/norestart")
            Start-Process "msiexec.exe" -ArgumentList $arguments -Wait -ErrorAction Stop
            Write-Status "'$AppName' installed successfully." -Type Success
        }
        catch {
            try {
                # Fallback for .exe installers
                $arguments = @("`"$installerPath`"", "/S", "/v`"/qn`"")
                Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -ErrorAction Stop
                Write-Status "'$AppName' installed successfully." -Type Success
            }
            catch {
                Write-Status "Failed to install '$AppName'. Please try installing it manually." -Type Error
                Write-Status "  Reason: $($_.Exception.Message)" -Type Info
            }
        }
    }
    else {
        Write-Status "Skipping installation of '$AppName' as no installer was selected." -Type Info
    }
}

function Optimize-System {
    <#
    .SYNOPSIS
        Applies system optimizations for performance and privacy.
    #>
    Write-Status -Message "Applying system optimizations..." -Type Step

    # Set power plan
    $powerPlanGuid = $null
    $powerPlanName = $null

    # Well-known GUIDs for power plans
    $ultimateGuid = "e9a42b02-d5df-448d-aa00-03f14749eb61"
    $highPerfGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
    $balancedGuid = "381b4222-f694-41f0-9685-ff5bb260df2e"

    $powerPlans = powercfg -list | ForEach-Object {
        if ($_ -match "GUID: (.*) \((.*)\)") {
            [PSCustomObject]@{
                Guid = $matches[1].Trim()
                Name = $matches[2].Trim()
            }
        }
    }

    $ultimatePlan = $powerPlans | Where-Object { $_.Guid -eq $ultimateGuid }
    $highPerfPlan = $powerPlans | Where-Object { $_.Guid -eq $highPerfGuid }
    $balancedPlan = $powerPlans | Where-Object { $_.Guid -eq $balancedGuid }

    if ($ultimatePlan) {
        $powerPlanGuid = $ultimatePlan.Guid
        $powerPlanName = $ultimatePlan.Name
    }
    elseif ($highPerfPlan) {
        $powerPlanGuid = $highPerfPlan.Guid
        $powerPlanName = $highPerfPlan.Name
    }
    elseif ($balancedPlan) {
        $powerPlanGuid = $balancedPlan.Guid
        $powerPlanName = $balancedPlan.Name
    }

    if ($powerPlanGuid) {
        try {
            powercfg -setactive $powerPlanGuid
            Write-Status "Power plan set to '$powerPlanName'." -Type Success
        }
        catch {
            Write-Status "Failed to set the power plan to '$powerPlanName'." -Type Warning
        }
    }
    else {
        Write-Status "Could not find a suitable power plan to apply." -Type Warning
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

    Write-Host "  → Creating installer cache cleanup task..."
    try {
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -Command `"Remove-Item -Path \"$env:TEMP\\*\" -Recurse -Force -ErrorAction SilentlyContinue`""
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At "3am"
        Register-ScheduledTask -TaskName "Installer-Cache-Cleanup" -Action $action -Trigger $trigger -Force -ErrorAction Stop | Out-Null
        Write-Status "Installer cache cleanup task created successfully." -Type Success
    } catch {
        Write-Status "Failed to create installer cache cleanup task." -Type Warning
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
        $localInstallers = @("Google Chrome", "NVIDIA App")
        foreach ($package in $selectedPackages) {
            if ($localInstallers -contains $package) {
                Install-LocalPackage -AppName $package
            }
            else {
                Install-WingetPackage -PackageId $package
            }
        }
    }
    else {
        Write-Status -Message "No applications selected. Skipping installation." -Type Info
    }

    # Phase 3: System Optimization & Hardening
    Write-Status -Message "PHASE 3: System Optimization & Hardening" -Type Step
    Optimize-System
    Enable-SecurityHardening

    # Phase 4: Cleanup
    Write-Status -Message "PHASE 4: Cleanup" -Type Step
    Remove-Bloatware

    # Phase 5: Maintenance
    Write-Status -Message "PHASE 5: Maintenance" -Type Step
    Schedule-MaintenanceTasks

    # Phase 6: Completion Summary
    Write-Status -Message "PHASE 6: Completion Summary" -Type Step
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
