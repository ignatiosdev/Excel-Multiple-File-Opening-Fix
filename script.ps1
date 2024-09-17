# PowerShell Script to Fix Excel File Opening Issue

# Define registry paths and values
$registryPaths = @(
    "HKLM:\Software\Classes\Excel.Sheet.12\shell\Open\ddeexec",  # For Excel 2013 and newer
    "HKLM:\Software\Classes\Excel.Sheet.8\shell\Open\ddeexec"    # For Excel 97-2003
)

$desiredValue = '[open("%1" /ou "%u")]'

# Function to set registry value
function Set-RegistryValue {
    param (
        [string]$path,
        [string]$value
    )

    # Check if the path exists
    if (Test-Path $path) {
        # Set the registry value
        Set-ItemProperty -Path $path -Name "(Default)" -Value $value
        Write-Output "Successfully set registry key: $path"
    } else {
        Write-Output "Registry path not found: $path. Attempting to create path."
        try {
            # Create necessary registry paths if they do not exist
            $parentPath = [System.IO.Path]::GetDirectoryName($path)
            if (-not (Test-Path $parentPath)) {
                New-Item -Path $parentPath -Force | Out-Null
            }
            # Create the registry path
            New-Item -Path $path -Force | Out-Null
            # Set the registry value
            Set-ItemProperty -Path $path -Name "(Default)" -Value $value
            Write-Output "Successfully created and set registry key: $path"
        } catch {
            Write-Output "Failed to create or set registry key: $path. Error: $_"
        }
    }
}

# Apply the fix
foreach ($path in $registryPaths) {
    Set-RegistryValue -path $path -value $desiredValue
}

Write-Output "Registry updates complete. Please restart Excel for changes to take effect."
