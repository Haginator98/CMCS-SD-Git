# Import required modules (installed via Tools.ps1)
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Prompt user to sign in to Microsoft Graph
Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome
Write-Host "This script will change the Office Location attribute for users in Entra ID (Azure AD)." -ForegroundColor Cyan
Start-Sleep -Seconds 1

# Retrieve all users (may take time if you have many users)
Write-Host "Fetching all users..." -ForegroundColor Cyan
$allUsers = Get-MgUser -All -Property "displayName,officeLocation,userPrincipalName,id"

# Old Office Location name
$oldOfficeLocation = Read-Host "Enter the old office location to search for"

# Filter users by office location
$users = $allUsers | Where-Object { $_.OfficeLocation -eq $oldOfficeLocation }

if ($users.Count -eq 0) {
    Write-Host "No users found with office location '$oldOfficeLocation'" -ForegroundColor Yellow
    Write-Host "Exiting script." -ForegroundColor Red
    return
} else {
    Write-Host "Found $($users.Count) users with office location '$oldOfficeLocation'" -ForegroundColor Cyan
}

# New Office Location name
$newOfficeLocation = Read-Host "Enter the new Office Location to be set"

# Confirm before making changes
$ready = $false
while (-not $ready) {
    $confirm = Read-Host "Are you sure you want to set '$newOfficeLocation' for $($users.Count) users? (y = yes, n = no, l = list users)"
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
            $users | Select-Object UserPrincipalName, DisplayName, OfficeLocation | Format-Table
        }
        default {
            Write-Host "Please enter y (yes), n (no), or l (list users)." -ForegroundColor Yellow
        }
    }
}

$counter = 0
foreach ($user in $users) {
    $counter++
    $percentComplete = [math]::Round(($counter / $users.Count) * 100)
    Write-Progress -Activity "Updating Office Location" -Status "Processing $counter of $($users.Count): $($user.UserPrincipalName)" -PercentComplete $percentComplete
    Update-MgUser -UserId $user.Id -OfficeLocation $newOfficeLocation
}

Write-Progress -Activity "Updating Office Location" -Completed
Write-Host "Update complete. $($users.Count) users updated." -ForegroundColor Green

# Disconnect from Microsoft Graph
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected. Script finished." -ForegroundColor Green
Start-Sleep -Seconds 4
