# Excel Multiple File Opening Fix

This repository provides a PowerShell script to resolve performance issues with opening subsequent Excel files from File Explorer. If you experience delays or slow performance when opening Excel files while another Excel file is already open, this script can help fix that problem.

## Purpose

The script modifies specific registry settings to ensure faster opening of Excel files. It addresses an issue where Excel files opened from File Explorer take a long time to open if another Excel file is already running.

## Instructions

### Download and Run the Script

To download and execute the script with one command, follow these steps:

1. **Open PowerShell as Administrator:**
   - Right-click on the Start menu and select “Windows PowerShell (Admin)” or “Windows Terminal (Admin)”.

2. **Execute the Following Command:**
   ```powershell
   Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ignatiosdev/Excel-Multiple-File-Opening-Fix/main/Fix-Excel-Opening-Issue.ps1" -OutFile "Fix-Excel-Opening-Issue.ps1"; .\Fix-Excel-Opening-Issue.ps1
