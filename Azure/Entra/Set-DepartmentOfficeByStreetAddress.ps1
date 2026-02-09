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

# Filter users where streetAddress matches target
$usersToUpdate = $allUsers | Where-Object { $_.StreetAddress -eq $targetStreet }

if ($usersToUpdate.Count -eq 0) {
    Write-Host "No users found with street address '$targetStreet'" -ForegroundColor Yellow
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    return
} else {
    Write-Host "Found $($usersToUpdate.Count) users with street address '$targetStreet'" -ForegroundColor Cyan
    Write-Host "`nUsers to be updated:" -ForegroundColor Yellow
    $usersToUpdate | Select-Object UserPrincipalName, DisplayName, StreetAddress, @{Name='Department';Expression={if($_.Department){$_.Department}else{"no department"}}}, @{Name='Office';Expression={if($_.OfficeLocation){$_.OfficeLocation}else{"no office"}}} | Format-Table

    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "CONFIRMATION" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Department will be set to: $newDepartment" -ForegroundColor Cyan
    Write-Host "Office will be set to: $newOffice" -ForegroundColor Cyan
    Write-Host "Number of users to update: $($usersToUpdate.Count)" -ForegroundColor Cyan
    $confirm = Read-Host "`nDo you want to proceed with these changes? [Y] Yes [N] No"
    
    if ($confirm -notmatch "[yY]") {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        return
    }

    $counter = 0
    foreach ($user in $usersToUpdate) {
        $counter++
        Write-Progress -Activity "Updating users" -Status "Processing $counter of $($usersToUpdate.Count)" -PercentComplete (($counter / $usersToUpdate.Count) * 100)
        Write-Host "Updating $($user.UserPrincipalName)..." -ForegroundColor Cyan

        Update-MgUser -UserId $user.Id `
            -Department $newDepartment `
            -OfficeLocation $newOffice
    }
    Write-Progress -Activity "Updating users" -Completed

    Write-Host "`n[SUCCESS] All $($usersToUpdate.Count) users have been updated." -ForegroundColor Green
}

Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected. Script finished." -ForegroundColor Green
Start-Sleep -Seconds 2
Clear-Host