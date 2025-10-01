# Install-Module Microsoft.Graph -Scope CurrentUser

# Sign in to Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All"

# Specify old and new department names
$oldDepartment = "Administrasjon og giveroppføling"
$newDepartment = "Administrasjon og giveroppfølging"

# Retrieve all users (may take time if you have many users)
$allUsers = Get-MgUser -All -Property "displayName,department,userPrincipalName,id"

# Filter users by department
$users = $allUsers | Where-Object { $_.Department -eq $oldDepartment }

if ($users.Count -eq 0) {
    Write-Host "No users found with department '$oldDepartment'" -ForegroundColor Yellow
} else {
    Write-Host "Found $($users.Count) users with department '$oldDepartment'" -ForegroundColor Cyan

    foreach ($user in $users) {
        Write-Host "Updating $($user.UserPrincipalName) with new department '$newDepartment'" -ForegroundColor Green
        Update-MgUser -UserId $user.Id -Department $newDepartment
    }

    Write-Host "Update complete." -ForegroundColor Green
}