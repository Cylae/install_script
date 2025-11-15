#Requires -Version 5.1
#Requires -PSEdition Desktop
#Requires -RunAsAdministrator

# Optimized PowerShell Deployment Script
#
# This script performs the following actions:
# 0. Ensures the script is run with Administrator privileges.
# 1. Upgrades all existing Winget packages.
# 2. Installs a defined list of packages from Winget.
# 3. Downloads and installs Google Chrome from a stable URL.
# 4. Provides an optional prompt to install NVIDIA App from a local file.
# 5. Registers a scheduled task to auto-update Winget packages at logon.
#
# Optimized by centralizing repetitive logic.

# --- 1. Function Definitions ---

Function Ensure-ModuleIsInstalled {
    param(
        [string]$ModuleName = "Microsoft.WinGet.Client"
    )

    # Check if the module is already imported or available
    if (Get-Module -Name $ModuleName -ListAvailable) {
        Write-Host "Module '$ModuleName' is already installed." -ForegroundColor Green
        Import-Module -Name $ModuleName
        return
    }

    Write-Host "Module '$ModuleName' not found. Starting installation process..." -ForegroundColor Yellow

    # Ensure PowerShellGet is up-to-date
    try {
        Write-Host "Checking for the latest version of PowerShellGet..."
        Install-Module -Name PowerShellGet -Force -ErrorAction Stop
        Write-Host "PowerShellGet is up-to-date." -ForegroundColor Green
    } catch {
        Write-Host "Failed to update PowerShellGet. Error: $_" -ForegroundColor Red
        return
    }

    # Ensure PSGallery is a trusted repository
    if ((Get-PSRepository | Where-Object { $_.Name -eq 'PSGallery' }).InstallationPolicy -ne 'Trusted') {
        Write-Host "Setting PSGallery as a trusted repository."
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
    }

    # Install the module
    try {
        Write-Host "Installing module '$ModuleName'..." -ForegroundColor Cyan
        Install-Module -Name $ModuleName -Force -ErrorAction Stop
        Write-Host "Module '$ModuleName' installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to install module '$ModuleName'. Error: $_" -ForegroundColor Red
        return
    }
}

Function Select-Applications {
    # Ensure WinGet client module is installed
    Ensure-ModuleIsInstalled

    # Import the module
    try {
        Write-Host "Importing module '$ModuleName'..."
        Import-Module -Name $ModuleName -ErrorAction Stop
        Write-Host "Module '$ModuleName' imported successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to import module '$ModuleName' after installation. Error: $_" -ForegroundColor Red
    }
}

Function Select-Applications {
    # Ensure WinGet client module is installed
    Ensure-ModuleIsInstalled

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Application Installer"
    $form.Size = New-Object System.Drawing.Size(300, 400)
    $form.StartPosition = "CenterScreen"

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75, 330)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(150, 330)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    $checkboxPanel = New-Object System.Windows.Forms.Panel
    $checkboxPanel.Location = New-Object System.Drawing.Point(10, 40)
    $checkboxPanel.Size = New-Object System.Drawing.Size(260, 280)
    $checkboxPanel.AutoScroll = $true
    $form.Controls.Add($checkboxPanel)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(10, 10)
    $searchBox.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($searchBox)

    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Location = New-Object System.Drawing.Point(165, 8)
    $searchButton.Size = New-Object System.Drawing.Size(75, 23)
    $searchButton.Text = "Search"
    $form.Controls.Add($searchButton)

    $statusBar = New-Object System.Windows.Forms.StatusBar
    $form.Controls.Add($statusBar)

    $populateList = {
        param ($packages, $preSelected)
        $checkboxPanel.Controls.Clear()
        $y = 0
        foreach ($package in $packages) {
            $checkbox = New-Object System.Windows.Forms.CheckBox
            $checkbox.Text = $package.Id
            $checkbox.Location = New-Object System.Drawing.Point(0, $y)
            $checkbox.Size = New-Object System.Drawing.Size(240, 20)
            if ($preSelected -contains $package.Id) {
                $checkbox.Checked = $true
            }
            $checkboxPanel.Controls.Add($checkbox)
            $y += 20
        }
    }

    $searchButton.Add_Click({
        $statusBar.Text = "Searching for applications..."
        $foundPackages = Find-WinGetPackage -Query $searchBox.Text | Where-Object { $_.Id -ne "Google.Chrome" }
        & $populateList $foundPackages $initialPackages.Id
        $statusBar.Text = "Search complete."
    })

    # Pre-populate with original packages
    $initialPackageIds = @(
        "Microsoft.AppInstaller",
        "Microsoft.WindowsTerminal",
        "Notepad++.Notepad++",
        "Discord.Discord",
        "Microsoft.DirectX",
        "Microsoft.VCRedist.2015+.x64",
        "VideoLAN.VLC",
        "Microsoft.VCRedist.2015+.x86",
        "Microsoft.DotNet.Runtime.9"
    )
    $initialPackages = @()
    foreach ($packageId in $initialPackageIds) {
        $package = Find-WinGetPackage -Id $packageId
        if ($null -ne $package) {
            $initialPackages += $package
        }
    }
    & $populateList $initialPackages $initialPackageIds

    $form.TopMost = $true
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedPackages = @()
        foreach ($checkbox in $checkboxPanel.Controls) {
            if ($checkbox.Checked) {
                $selectedPackages += $checkbox.Text
            }
        }
        return $selectedPackages
    } else {
        return @()
    }
}

Function Start-Log {
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $logFile = "$env:USERPROFILE\Desktop\log-$timestamp.log"
    Start-Transcript -Path $logFile
}

Start-Log

Function Invoke-WingetInstall {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$PackageId
    )

    Write-Host "--- Using Winget to install: $PackageId ---" -ForegroundColor Cyan

    # Execute the install command
    winget install --id $PackageId --accept-source-agreements --accept-package-agreements

    # Check the exit code for status
    if ($LASTEXITCODE -eq 0) {
        Write-Host "$PackageId installed successfully." -ForegroundColor Green
    } elseif ($LASTEXITCODE -eq -1978335189) {
        # Code -1978335189 detected (e.g., already installed or no applicable update), suppressing output.
        Write-Host "$PackageId is already installed or no update found." -ForegroundColor Yellow
    } else {
        Write-Host "Error or failure while installing $PackageId (Exit Code: $LASTEXITCODE)." -ForegroundColor Red
    }
    Write-Host "" # Adds a blank line for readability
}

Function Install-ChromeFromUrl {
    [CmdletBinding()]
    param (
        # Stable URL for Chrome Enterprise (64-bit MSI)
        [string]$Url = "https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi"
    )

    $AppName = "Google Chrome"
    $InstallerType = "msi"

    Write-Host "--- Installing $AppName from URL ---" -ForegroundColor Cyan

    # Define a temporary file path
    $tempFile = Join-Path $env:TEMP "Chrome-Installer.$InstallerType"

    # --- Download ---
    try {
        Write-Host "Downloading $AppName from $Url..."
        Invoke-WebRequest -Uri $Url -OutFile $tempFile -UseBasicParsing -ErrorAction Stop
        Write-Host "Download complete: $tempFile" -ForegroundColor Green
    } catch {
        Write-Host "Failed to download $AppName. Error: $_" -ForegroundColor Red
        return
    }

    # --- Install ---
    try {
        Write-Host "Starting silent installation for $AppName..."

        # MSI uses msiexec.exe for silent install
        $executable = "msiexec.exe"
        # /i for install, /qn for no UI
        $processArgs = "/i `"$tempFile`" /qn"

        $process = Start-Process -FilePath $executable -ArgumentList $processArgs -Wait -PassThru -ErrorAction Stop

        if ($process.ExitCode -eq 0) {
            Write-Host "$AppName installed successfully." -ForegroundColor Green
        } elseif ($process.ExitCode -eq 3010) {
            # 3010 is a common MSI exit code for "Success, reboot required"
            Write-Host "$AppName installed successfully. A reboot is required." -ForegroundColor Yellow
        } else {
            Write-Host "Error or failure during $AppName installation (Exit Code: $($process.ExitCode))." -ForegroundColor Red
        }
    } catch {
        Write-Host "Failed to start the installer process for $AppName. Error: $_" -ForegroundColor Red
    } finally {
        # --- Clean up ---
        if (Test-Path $tempFile) {
            Write-Host "Cleaning up installer: $tempFile"
            Remove-Item $tempFile -ErrorAction SilentlyContinue
        }
    }
    Write-Host ""
}

Function Apply-SystemOptimizations {
    [CmdletBinding()]
    param()

    Write-Host "--- Applying System Optimizations ---" -ForegroundColor Cyan

    # Helper function to safely set registry properties
    function Set-RegistryValue {
        param (
            [string]$Path,
            [string]$Name,
            $Value,
            [string]$Type = "DWord"
        )
        if (-not (Test-Path -Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
    }

    # Enable Ultimate Performance Plan
    powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61

    # Debloat
    Get-AppxPackage -AllUsers *Microsoft.XboxApp* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.Xbox.TCUI* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.XboxGameOverlay* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.XboxGamingOverlay* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.XboxSpeechToTextOverlay* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.ZuneVideo* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.ZuneMusic* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.WindowsMaps* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.People* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.YourPhone* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.MixedReality.Portal* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.MicrosoftSolitaireCollection* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.MinecraftUWP* | Remove-AppxPackage
    Get-AppxPackage -AllUsers *Microsoft.Microsoft3DViewer* | Remove-AppxPackage

    # Privacy
    Set-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "HideFileExt" -Value 0
    Set-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 1
    Set-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowSyncProviderNotifications" -Value 0
    Set-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SilentInstalledAppsEnabled" -Value 0
    Set-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-310093Enabled" -Value 0
    Set-RegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SubscribedContent-338389Enabled" -Value 0
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -Value 1
    Set-RegistryValue -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0

    # UI/UX
    Set-RegistryValue -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0
    Set-RegistryValue -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0

    # Restart Explorer to apply theme changes
    Stop-Process -Name explorer -Force
}

Function Install-LocalPackage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$AppName,

        [Parameter(Mandatory=$true)]
        [string]$DialogTitle,

        [Parameter(Mandatory=$true)]
        [string]$SilentArguments,

        [Parameter(Mandatory=$false)]
        [string]$Filter = "Installers (*.exe, *.msi)|*.exe;*.msi"
    )

    # Load required assembly only when this function is called
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    } catch {
        Write-Host "Failed to load System.Windows.Forms. Skipping optional local installs." -ForegroundColor Red
        return
    }

    Write-Host "--- Optional $AppName Installation ---" -ForegroundColor Cyan
    Write-Host "A file dialog will open. Please select the $AppName installer file."
    Write-Host "If you don't want to install it, click 'Cancel'."

    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = $DialogTitle
    $fileDialog.Filter = $Filter

    $result = $fileDialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $installerPath = $fileDialog.FileName
        Write-Host "--- Attempting to install $AppName from local file: $installerPath ---"

        try {
            $process = Start-Process -FilePath $installerPath -ArgumentList $SilentArguments -Wait -PassThru -ErrorAction Stop

            if ($process.ExitCode -eq 0) {
                Write-Host "$AppName installed successfully from file." -ForegroundColor Green
            } else {
                Write-Host "Error or failure while installing $AppName from file (Exit Code: $($process.ExitCode))." -ForegroundColor Red
            }
        } catch {
            Write-Host "Failed to start the installer process for $AppName. Error: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "$AppName installation skipped (user clicked Cancel)." -ForegroundColor Yellow
    }
    Write-Host "" # Adds a blank line for readability
}

Function Register-WingetAutoUpgradeTask {
    [CmdletBinding()]
    param()

    $TaskName = "Winget-Auto-Upgrade"
    Write-Host "--- Registering Automatic Winget Upgrade Task ---" -ForegroundColor Cyan

    try {
        # 1. Define the Action
        # We must use the full command syntax Winget requires for silent operation
        $action = New-ScheduledTaskAction -Execute "winget" -Argument "upgrade --all --accept-source-agreements --accept-package-agreements"

        # 2. Define the Trigger (Runs when any user logs on)
        $trigger = New-ScheduledTaskTrigger -AtLogOn

        # 3. Define the Principal (Runs as the user, but elevated)
        # -RunLevel Highest ensures it has admin rights to install software
        # -UserId "INTERACTIVE" applies this to any user who logs in.
        $principal = New-ScheduledTaskPrincipal -RunLevel Highest -UserId "NT AUTHORITY\INTERACTIVE"

        # 4. Define the Settings (Using universally compatible parameters)
        # We are omitting battery settings for maximum compatibility.
        # The task will use the defaults (don't start on battery, stop if switching to battery).
        $settings = New-ScheduledTaskSettingsSet `
            -RunOnlyIfNetworkAvailable `         # Waits for network connection
            -MultipleInstances IgnoreNew `       # Replaced -NewInstance for older systems
            -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
            -StartWhenAvailable                 # Run if the logon was missed (e.g. machine was off)

        # 5. Register the Task
        # -Force will overwrite the task if it already exists
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force -ErrorAction Stop

        Write-Host "Scheduled task '$TaskName' registered successfully." -ForegroundColor Green
        Write-Host "It will run 'winget upgrade --all' at every user logon."

    } catch {
        Write-Host "Failed to register scheduled task '$TaskName'. Error: $_" -ForegroundColor Red
    }
    Write-Host ""
}


# --- 2. Upgrading all existing packages ---
Write-Host "--- 1. Upgrading all existing packages (using Winget) ---" -ForegroundColor Cyan
Write-Host "This will update all packages currently managed by Winget."

# Execute the upgrade command
winget upgrade --all --accept-source-agreements --accept-package-agreements

if ($LASTEXITCODE -eq 0) {
    Write-Host "All packages upgraded successfully." -ForegroundColor Green
} elseif ($LASTEXITCODE -eq -1978335189) {
    # Code -1978335189 detected (e.g., no packages to upgrade), suppressing output.
    Write-Host "No packages needed upgrading." -ForegroundColor Yellow
} else {
    Write-Host "Error or failure during winget upgrade (Exit Code: $LASTEXITCODE)." -ForegroundColor Red
}
Write-Host ""

# --- 3. Automated Download Installation ---
Write-Host "--- 2. Installing Google Chrome ---" -ForegroundColor Cyan
Install-ChromeFromUrl

# --- 4. Winget Package Installation ---
Write-Host "--- 3. Installing Winget Packages ---" -ForegroundColor Cyan
Write-Host "Installing all required software from the list..."

$selectedPackages = Select-Applications

if ($selectedPackages) {
    # Loop to install each package using the reusable function
    foreach ($package in $selectedPackages) {
        Invoke-WingetInstall -PackageId $package
    }
} else {
    Write-Host "No applications selected. Skipping installation." -ForegroundColor Yellow
}

# --- 5. System Optimizations ---
Write-Host "--- 4. Applying System Optimizations ---" -ForegroundColor Cyan
Apply-SystemOptimizations

# --- 6. Optional Local Installations ---
Write-Host "--- 5. Optional Local Package Installations ---" -ForegroundColor Cyan
Install-LocalPackage -AppName "NVIDIA App" -DialogTitle "Select NVIDIA App Installer (Optional - Click Cancel to skip)" -SilentArguments "-s"

# --- 7. Register Auto-Upgrade Task ---
Write-Host "--- 6. Registering Scheduled Task ---" -ForegroundColor Cyan
Register-WingetAutoUpgradeTask

Write-Host "--- All package installations are complete. ---" -ForegroundColor Cyan

Stop-Transcript