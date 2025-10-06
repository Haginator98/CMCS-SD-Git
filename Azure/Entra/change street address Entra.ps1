# Install-Module Microsoft.Graph -Scope CurrentUser

# Log in to Microsoft Graph
#Connect-MgGraph -Scopes "User.ReadWrite.All"

# Check if Microsoft.Graph module is installed
$Module = Get-Module -Name Microsoft.Graph -ListAvailable
if ($Module.Count -eq 0) {
    Write-Host "Microsoft.Graph module is not available." -ForegroundColor Yellow
    $Confirm = Read-Host "Are you sure you want to install the module? [Y] Yes [N] No"
    if ($Confirm -match "[yY]") {
        Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Cyan
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
        if ($?) {
            Write-Host "Microsoft.Graph module installed successfully." -ForegroundColor Green
        } else {
            Write-Host "Failed to install Microsoft.Graph module." -ForegroundColor Red
            Exit
        }
    } else {
        Write-Host "Microsoft.Graph module is required. Please install it using Install-Module Microsoft.Graph cmdlet."
        Exit
    }
}
# Do NOT import the module explicitly
Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome

Write-Host "This script will change the Street Address attribute for users in Entra ID (Azure AD)." -ForegroundColor Cyan
Start-Sleep -Seconds 1

# Prompt for old and new address
$oldAddress = Read-Host "Enter the old street address to search for"
$newAddress = Read-Host "Enter the new street address to set"

# Get all users (may take time if there are many users)
$allUsers = Get-MgUser -All -Property "displayName,streetAddress,userPrincipalName,id"
Write-Host "Fetching all users..." -ForegroundColor Cyan

# Filter locally on streetAddress
$usersToUpdate = $allUsers | Where-Object { $_.StreetAddress -eq $oldAddress }

if ($usersToUpdate.Count -eq 0) {
    Write-Host "No users found with street address '$oldAddress'" -ForegroundColor Yellow
} else {
    Write-Host "Found $($usersToUpdate.Count) users with street address '$oldAddress'" -ForegroundColor Cyan

    $ready = $false
    while (-not $ready) {
        $confirm = Read-Host "Are you sure you want to set street address '$newAddress' for $($usersToUpdate.Count) users? (y = yes, n = no, l = list users)"
        switch ($confirm.ToLower()) {
            'y' {
                $ready = $true
            }
            'n' {
                Write-Host "Operation cancelled. Exiting script" -ForegroundColor Red
                Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
                Disconnect-MgGraph | Out-Null
                Write-Host "Disconnected. Script finished." -ForegroundColor Green
                Start-Sleep -Seconds 2
                return
            }
            'l' {
                Write-Host "Users to be updated:" -ForegroundColor Cyan
                $usersToUpdate | Select-Object UserPrincipalName, DisplayName, streetAddress | Format-Table
            }
            default {
                Write-Host "Please enter y (yes), n (no), or l (list users)." -ForegroundColor Yellow
            }
        }
    }

    foreach ($user in $usersToUpdate) {
        Write-Host "Updating $($user.UserPrincipalName)..."
        Update-MgUser -UserId $user.Id -StreetAddress $newAddress
    }

    Write-Host "Update completed." -ForegroundColor Green
}
# Disconnect from Microsoft Graph
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected. Script finished." -ForegroundColor Green
Start-Sleep -Seconds 3
Clear-Host