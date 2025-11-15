# Manual Testing and Validation Guide

This document provides a step-by-step guide for manually testing and validating the PowerShell deployment script.

## Prerequisites

* A clean installation of Windows 11 24H2.
* The `script.ps1` file.

## Testing Steps

1. **Run the script as an administrator.**
    * Right-click on the `script.ps1` file and select "Run with PowerShell".
    * The script should open a PowerShell window and start executing.

2. **Verify the application selection GUI.**
    * The script should open a GUI window with a list of available applications.
    * The list of applications should be populated from `winget search`.
    * You should be able to select and deselect applications using the checkboxes.
    * You should be able to search for applications using the search box.
    * You should be able to refresh the list of applications using the refresh button.
    * Click the "OK" button to start the installation.

3. **Verify the application installation.**
    * The script should install the selected applications.
    * The script should log the installation progress to the console and to a log file on the desktop.
    * The script should not install any applications that were not selected.

4. **Verify the system optimizations.**
    * The script should apply the system optimizations.
    * The script should log the optimization progress to the console and to a log file on the desktop.
    * You should be able to verify that the optimizations have been applied by checking the system settings.

5. **Verify the optional local installation.**
    * The script should prompt you to select an installer for the NVIDIA App.
    * If you select an installer, the script should install the NVIDIA App.
    * If you click "Cancel", the script should skip the installation of the NVIDIA App.

6. **Verify the scheduled task.**
    * The script should create a scheduled task to automatically upgrade Winget packages at logon.
    * You should be able to verify that the scheduled task has been created by opening the Task Scheduler.

7. **Verify the log file.**
    * The script should create a log file on the desktop.
    * The log file should contain a record of all the actions performed by the script.
    * The log file should not contain any errors.

## Expected Results

* The script should run without any errors.
* The selected applications should be installed.
* The system optimizations should be applied.
* The optional local installation should work as expected.
* The scheduled task should be created.
* The log file should be created and should not contain any errors.
