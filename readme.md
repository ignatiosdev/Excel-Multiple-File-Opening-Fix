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
