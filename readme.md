# Microsoft Excel Multiple File Opening Fix

 If you experience delays or slow performance when opening Excel files while another Excel file is already open, this script can help fix that problem.

## Instructions

### Download and Run the Script

To download and execute the script with one command, follow these steps:

1. **Open PowerShell as Administrator:**
   - Right-click on the Start menu and select “Windows PowerShell (Admin)” or “Windows Terminal (Admin)”.

2. **Execute the Following Command:**
   ```powershell
   Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ignatiosdev/Excel-Multiple-File-Opening-Fix/main/script.ps1" -OutFile "script.ps1"; .\script.ps1

### **How It Works**

This PowerShell script fixes the delay experienced when opening multiple Excel files from Windows File Explorer while another Excel instance is already running.

#### **How It Works:**
1. **Registry Paths**: The script updates specific registry keys related to Excel's Dynamic Data Exchange (DDE) settings:
   - `HKEY_CLASSES_ROOT\Excel.Sheet.12\shell\Open\ddeexec` (for Excel 2013 and newer)
   - `HKEY_CLASSES_ROOT\Excel.Sheet.8\shell\Open\ddeexec` (for Excel 97-2003)
2. **Registry Updates**: It sets the `Default` value of these keys to `[open("%1" /ou "%u")]`, optimizing the way Excel handles file openings via DDE, leading to faster performance when opening multiple files.
3. **Path Creation**: If the specified registry paths do not exist, the script creates them to ensure that the necessary registry structure is in place.

#### **Why the Bug Happens:**
The issue arises due to outdated or incorrect DDE settings in the Windows Registry. DDE is a protocol used for inter-process communication and can sometimes become misconfigured, causing delays in opening files when multiple Excel instances are involved.

#### **Acknowledgment:**
A special shoutout to [adamAmiga0](https://answers.microsoft.com/en-us/msoffice/forum/all/excel-opens-2nd-subsequent-file-from-explorer/907165a8-44a4-4212-a871-6525b7553eaa) for discovering and sharing this solution. Their insights on Microsoft Answers were instrumental in identifying the root cause and providing a fix for this common issue.

