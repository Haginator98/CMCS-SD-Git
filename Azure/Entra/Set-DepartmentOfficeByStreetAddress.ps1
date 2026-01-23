# Import required modules (installed via Tools.ps1)
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Connect to Microsoft Graph
Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome

Write-Host "This script will change Department and Office location for users based on their Street Address..." -ForegroundColor Cyan

#Fetch all users with required properties
Write-Host "Fetching all users" -ForegroundColor Cyan
$allUsers = Get-MgUser -All -Property "displayName,streetAddress,department,officeLocation,userPrincipalName,id"

# Target street address value for filtering
do {
    $targetStreet = Read-Host "Enter the street address to target (e.g., '8410')"
    $usersToUpdate = $allUsers | Where-Object { $_.StreetAddress -eq $targetStreet }
    if ($usersToUpdate.Count -eq 0) {
        Write-Host "No users found with street address '$targetStreet'. Please try again." -ForegroundColor Yellow
    }
} while ($usersToUpdate.Count -eq 0)

Write-Host "Found $($usersToUpdate.Count) users with street address '$targetStreet'" -ForegroundColor Green

Write-Host "You will be prompted to enter new values for Department and Office Location." -ForegroundColor Cyan
# New values to update
$newDepartment = Read-Host "Enter the new department name to be set" 
$newOffice = Read-Host "Enter the new office location to be set"

# Fetch all users with required properties
#$allUsers = Get-MgUser -All -Property "displayName,streetAddress,department,officeLocation,userPrincipalName,id"

# Filter users where streetAddress is exactly "8410"
$usersToUpdate = $allUsers | Where-Object { $_.StreetAddress -eq $targetStreet }

if ($usersToUpdate.Count -eq 0) {
    Write-Host "No users found with street address '$targetStreet'" -ForegroundColor Yellow
} else {
    Write-Host "Found $($usersToUpdate.Count) users with street address '$targetStreet'" -ForegroundColor Cyan

    Write-Host "Are you sure you want to set Department '$newDepartment' and OfficeLocation '$newOffice' for $($usersToUpdate.Count) users? (y = yes, n = no, l = list users)" -ForegroundColor Yellow
    $ready = $false
    while (-not $ready) {
        $confirm = Read-Host "(y = yes, n = no, l = list users)"
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
                $usersToUpdate | Select-Object UserPrincipalName, DisplayName, StreetAddress, Department, OfficeLocation | Format-Table
            }
            default {
                Write-Host "Please enter y (yes), n (no), or l (list users)." -ForegroundColor Yellow
            }
        }
    }

    foreach ($user in $usersToUpdate) {
        Write-Host "Updating $($user.UserPrincipalName)..."

        Update-MgUser -UserId $user.Id `
            -Department $newDepartment `
            -OfficeLocation $newOffice
    }

    Write-Host "All matching users have been updated." -ForegroundColor Green
    Start-Sleep -Seconds 1
}

Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected. Script finished." -ForegroundColor Green
Start-Sleep -Seconds 2
Clear-Host