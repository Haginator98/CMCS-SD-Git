# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All"

# Target street address value for filtering
$targetStreet = "8410"

# New values to update
$newDepartment = "Politikk og samfunnskontakt "
$newOffice = "Forskning, forebygging og kreftbehandling"

# Fetch all users with required properties
$allUsers = Get-MgUser -All -Property "displayName,streetAddress,department,officeLocation,userPrincipalName,id"

# Filter users where streetAddress is exactly "8410"
$usersToUpdate = $allUsers | Where-Object { $_.StreetAddress -eq $targetStreet }

if ($usersToUpdate.Count -eq 0) {
    Write-Host "No users found with street address '$targetStreet'" -ForegroundColor Yellow
} else {
    Write-Host "Found $($usersToUpdate.Count) users with street address '$targetStreet'" -ForegroundColor Cyan

    foreach ($user in $usersToUpdate) {
        Write-Host "Updating $($user.UserPrincipalName)..."

        Update-MgUser -UserId $user.Id `
            -Department $newDepartment `
            -OfficeLocation $newOffice
    }

    Write-Host "All matching users have been updated." -ForegroundColor Green
}