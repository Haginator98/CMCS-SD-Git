# Import required modules (installed via Tools.ps1)
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Prompt user to sign in to Microsoft Graph
Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome
Write-Host "This script will change the Department attribute for users in Entra ID (Azure AD)." -ForegroundColor Cyan
Start-Sleep -Seconds 1

$continueEditing = 'y'

while ($continueEditing -eq 'y') {
    # Retrieve all users (may take time if you have many users)
    Write-Host "Fetching all users..." -ForegroundColor Cyan
    $allUsers = Get-MgUser -All -Property "displayName,department,userPrincipalName,id"

    # Old Department name
    $oldDepartment = Read-Host "Enter the old department name to search for"

    # Filter users by department
    $users = $allUsers | Where-Object { $_.Department -eq $oldDepartment }

    if ($users.Count -eq 0) {
        Write-Host "No users found with department '$oldDepartment'" -ForegroundColor Yellow
    } else {
        Write-Host "Found $($users.Count) users with department '$oldDepartment'" -ForegroundColor Cyan

        # New Department name
        $newDepartment = Read-Host "Enter the new Department name to be set"

        # Confirm before making changes
        $ready = $false
        while (-not $ready) {
            $confirm = Read-Host "Are you sure you want to set '$newDepartment' for $($users.Count) users? (y = yes, n = no, l = list users)"
            switch ($confirm.ToLower()) {
                'y' {
                    $ready = $true
                }
                'n' {
                    Write-Host "Operation cancelled. Exiting script" -ForegroundColor Red
                    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
                    Disconnect-MgGraph | Out-Null
                    Start-Sleep -Seconds 2
                    Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green
                    return
                }
                'l' {
                    Write-Host "Users to be updated:" -ForegroundColor Cyan
                    $users | Select-Object UserPrincipalName, DisplayName,
                        @{Name = 'oldDepartment'; Expression = { $_.Department }},
                        @{Name = 'newDepartment'; Expression = { $newDepartment }} | Format-Table
                }
                default {
                    Write-Host "Please enter y (yes), n (no), or l (list users)." -ForegroundColor Yellow
                }
            }
        }

        $updatedCount = 0

        foreach ($user in $users) {
            Write-Host "Updating $($user.UserPrincipalName) with new department '$newDepartment'" -ForegroundColor Green
            Update-MgUser -UserId $user.Id -Department $newDepartment
            $updatedCount++
        }

        Write-Host "$updatedCount / $($users.Count) users updated." -ForegroundColor Green
    }

    do {
        $continueEditing = (Read-Host "Do you want to make more changes? (y/n)").ToLower()
        if ($continueEditing -notin @('y', 'n')) {
            Write-Host "Please enter y (yes) or n (no)." -ForegroundColor Yellow
        }
    } while ($continueEditing -notin @('y', 'n'))
}

# Disconnect from Microsoft Graph
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Start-Sleep -Seconds 2
Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green

