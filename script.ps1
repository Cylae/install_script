# Optimized PowerShell Deployment Script
#
# This script performs the following actions:
# 1. Upgrades all existing Winget packages.
# 2. Installs a defined list of packages from Winget.
# 3. Downloads and installs Google Chrome from a stable URL.
# 4. Provides an optional prompt to install NVIDIA App from a local file.
# 5. Registers a scheduled task to auto-update Winget packages at logon.
#
# Optimized by centralizing repetitive logic.

# --- 1. Function Definitions ---

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

# --- 3. Winget Package Installation ---
Write-Host "--- 2. Installing Winget Packages ---" -ForegroundColor Cyan
Write-Host "Installing all required software from the list..."

# Combined list of all packages to install via Winget
$wingetPackages = @(
    # Core
    "Microsoft.AppInstaller",
    "Microsoft.WindowsTerminal",
    # Remaining
    "Notepad++.Notepad++",
    "Discord.Discord",
    "Microsoft.DirectX",
    "Microsoft.VCRedist.2015+.x64",
    "VideoLAN.VLC",
    "Microsoft.VCRedist.2015+.x86",
    "Microsoft.DotNet.Runtime.9"
    # NVIDIA.NVIDIAApp removed as requested
)

# Loop to install each package using the reusable function
foreach ($package in $wingetPackages) {
    Invoke-WingetInstall -PackageId $package
}

# --- 4. Automated Download Installation ---
Write-Host "--- 3. Installing packages from URL ---" -ForegroundColor Cyan
Install-ChromeFromUrl

# --- 5. Optional Local Installations ---
Write-Host "--- 4. Optional Local Package Installations ---" -ForegroundColor Cyan
Install-LocalPackage -AppName "NVIDIA App" -DialogTitle "Select NVIDIA App Installer (Optional - Click Cancel to skip)" -SilentArguments "-s"

# --- 6. Register Auto-Upgrade Task ---
Write-Host "--- 5. Registering Scheduled Task ---" -ForegroundColor Cyan
Register-WingetAutoUpgradeTask

Write-Host "--- All package installations are complete. ---" -ForegroundColor Cyan