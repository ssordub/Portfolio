@echo off
Rem Make powershell read this file, skip a number of lines, and execute it.
Rem This works around .ps1 bad file association as non executables.
PowerShell -Command "Get-Content '%~dpnx0' | Select-Object -Skip 5 | Out-String | Invoke-Expression"
goto :eof

do {
    # List all environment variables
    Get-ChildItem Env: | Format-Table -AutoSize

    # Prompt the user for an environment variable name
    $envVarName = Read-Host "Enter the name of the environment variable to modify (or type 'exit' to quit)"

    if ($envVarName -eq 'exit') {
        Write-Host "Exiting the script."
        break
    }

    # Validate the environment variable name
    if (-not (Test-Path "Env:\$envVarName")) {
        Write-Host "Error: Environment variable '$envVarName' does not exist."
    } else {
        # Prompt the user for a new value
        $newValue = Read-Host "Enter the new value for '$envVarName'"

        # Confirm the action
        $confirmation = Read-Host "Are you sure you want to update the value of '$envVarName' to '$newValue'? (Y/N)"
        if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
            # Set the new value for the environment variable
            [Environment]::SetEnvironmentVariable($envVarName, $newValue, "Machine")
            Write-Host "Environment variable '$envVarName' has been updated with the new value."
        } else {
            Write-Host "Update canceled."
        }
    }
} while ($true)