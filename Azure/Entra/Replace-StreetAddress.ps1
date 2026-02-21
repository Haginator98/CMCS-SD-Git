# Import required modules (installed via Tools.ps1)
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome

Write-Host "This script will change the Street Address attribute for users in Entra ID (Azure AD)." -ForegroundColor Cyan
Start-Sleep -Seconds 1
# Get all users (may take time if there are many users)
$allUsers = Get-MgUser -All -Property "displayName,streetAddress,userPrincipalName,id"
Write-Host "Fetching all users..." -ForegroundColor Cyan

# Prompt for old address
$oldAddress = Read-Host "Enter the old street address to search for"

# Filter users by department
$users = $allUsers | Where-Object { $_.StreetAddress -eq $oldAddress }

if ($users.Count -eq 0) {
    Write-Host "No users found with department '$oldAddress'" -ForegroundColor Yellow
    Write-Host "Exiting script." -ForegroundColor Red
    return
} else {
    Write-Host "Found $($users.Count) users with department '$oldAddress'" -ForegroundColor Cyan
}

# Prompt for new address
$newAddress = Read-Host "Enter the new street address to set"

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