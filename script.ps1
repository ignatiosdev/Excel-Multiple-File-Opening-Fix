# PowerShell Script to Fix Excel File Opening Issue

# Define registry paths and values
$registryPaths = @(
    "HKCR:\Excel.Sheet.8\shell\Open\ddeexec",
    "HKCR:\Excel.Sheet.12\shell\Open\ddeexec"
)

$desiredValue = '[open("%1" /ou "%u")]'

# Function to set registry value
function Set-RegistryValue {
    param (
        [string]$path,
        [string]$value
    )

    if (Test-Path $path) {
        Set-ItemProperty -Path $path -Name "(Default)" -Value $value
        Write-Output "Successfully set registry key: $path"
    } else {
        Write-Output "Registry path not found: $path"
    }
}

# Apply the fix
foreach ($path in $registryPaths) {
    Set-RegistryValue -path $path -value $desiredValue
}

Write-Output "Registry updates complete. Please restart Excel for changes to take effect."
