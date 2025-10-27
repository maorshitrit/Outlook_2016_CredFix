# Outlook Credential Reset Script (Outlook 2016)

## Overview

This script is designed to solve a common issue in **Outlook 2016**, where after a user changes their password, Outlook continues to use the old cached credentials, causing login errors. 

The script performs the following:

1. Retrieves the current password hash of the user from **Active Directory**.  
2. Compares it to a previously stored hash.  
3. If a password change is detected:
    - Clears relevant Outlook/Office credentials from **Windows Credential Manager**.  
    - Logs all actions for auditing.  
    - Updates the stored hash file with the new password hash.

---

## Files

- **Password Hash File**:  
  Stores the last known password hash.  
  **Location:** `$env:APPDATA\OutlookPasswordHash.txt` (or `$env:APPDATA\OutlookPasswordHash_TEST.txt` in test mode)

- **Log File**:  
  Records all actions performed by the script for monitoring and debugging.  
  **Location:** `$env:TEMP\OutlookCredReset.log` (or `$env:TEMP\OutlookCredReset_TEST.log` in test mode)

---

## Usage

1. Run the script in **PowerShell**.  
2. The script automatically detects if the password has changed.  
3. If a change is detected, the script will:
    - Delete Outlook-related credentials from the Windows Credential Manager.
    - Write all actions to the log file.
    - Update the password hash file.

---

## Notes

- Designed for **Outlook 2016** in Windows environments with Active Directory.  
- Helps prevent repeated login errors due to cached credentials after a password change.  
- Can be run in **test mode** to verify functionality without deleting actual credentials.

---

## Example Locations in Test Mode

```text
Password Hash File: C:\Users\<username>\AppData\Roaming\OutlookPasswordHash_TEST.txt
Log File: C:\Users\<username>\AppData\Local\Temp\OutlookCredReset_TEST.log
