## Overview

This script was inspired by Microsoft’s CollectUpdate.vbs, adapting its functionality for Office 2016 patch collection. This VBScript extracts and renames Microsoft Office 2016 patch files (.msp) from C:\Windows\Installer, sorting them by modification date and naming them as num-OriginalFileName.msp (e.g., 0-IEAWSDC-X-NONE.MSP). It creates a log file only when errors occur, ensuring minimal output unless issues arise.

## Usage

Run the script with administrative privileges (right-click, "Run as administrator") to access system files and avoid permission issues.
Patches are copied to %SystemDrive%\Office2016Updates\ (or %TEMP%\Office2016Updates\ if the first fails), opening the folder upon completion.
Check error_log.txt in the target folder for any errors, like copy failures.

## Next Steps

Verify system updates using wmic qfe list or Windows Update history if you expect missing patches. Ensure admin rights for future runs.
