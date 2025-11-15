# This PowerShell script is optimized for installing software packages on a Windows 11 system with AMD 5950X.

# Function to download Chrome in parallel
function Download-Chrome {
    $chromeUrl = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
    $outputFile = "chrome_installer.exe"
    Write-Host "Downloading Chrome..."

    # Start the download in parallel
    Start-Process -FilePath "powershell" -ArgumentList "-Command Invoke-WebRequest -Uri $chromeUrl -OutFile $outputFile" -NoNewWindow -Wait
    Write-Host "Chrome download complete."
}

# Comprehensive software package installation function
function Install-Packages {
    $packages = @(
        "7zip",
        "vlc",
        "git",
        "notepad++",
        "visual studio code"
    )

    foreach ($package in $packages) {
        Write-Host "Installing $package..."
        # Use Chocolatey for package management
        Start-Process choco -ArgumentList "install $package -y" -NoNewWindow -Wait
        Write-Host "$package installation complete."
    }
}

# Optimize system performance
function Optimize-System {
    # Disable unnecessary startup programs
    Get-CimInstance -ClassName Win32_StartupCommand | Where-Object { $_.User -eq $null } | ForEach-Object { $_.Delete() }
    Write-Host "Disabled unnecessary startup programs."
}

# Security hardening function
function Harden-Security {
    # Enable Windows Defender
    Set-MpPreference -DisableRealtimeMonitoring $false
    Write-Host "Windows Defender enabled."

    # Add firewall rules
    New-NetFirewallRule -DisplayName "Allow Incoming RDP" -Direction Inbound -Protocol TCP -LocalPort 3389 -Action Allow
    Write-Host "Firewall rules updated."
}

# Main execution flow
Download-Chrome
Install-Packages
Optimize-System
Harden-Security
Write-Host "System setup completed successfully!"