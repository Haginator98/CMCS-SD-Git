# Import required modules (installed via Tools.ps1)
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Connect to Microsoft Graph
Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Set Manager Based on Street Address" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "This script will update the Manager attribute for users based on their Street Address.`n" -ForegroundColor Cyan

# Fetch all users with required properties
Write-Host "Fetching all users..." -ForegroundColor Yellow
$allUsers = Get-MgUser -All -Property "displayName,streetAddress,userPrincipalName,id,mail"
Write-Host "Retrieved $($allUsers.Count) users from Entra ID.`n" -ForegroundColor Green

# Get target street address
do {
    $targetStreet = Read-Host "Enter the street address to target (e.g., '8410')"
    $usersToUpdate = $allUsers | Where-Object { $_.StreetAddress -eq $targetStreet }
    
    if ($usersToUpdate.Count -eq 0) {
        Write-Host "No users found with street address '$targetStreet'. Please try again." -ForegroundColor Yellow
    } else {
        Write-Host "Found $($usersToUpdate.Count) users with street address '$targetStreet'`n" -ForegroundColor Green
    }
} while ($usersToUpdate.Count -eq 0)

# Get the new manager
$validManager = $false
do {
    $managerEmail = Read-Host "Enter the email address of the new manager"
    
    try {
        $manager = Get-MgUser -UserId $managerEmail -ErrorAction Stop
        Write-Host "Manager found: $($manager.DisplayName) ($($manager.UserPrincipalName))" -ForegroundColor Green
        $validManager = $true
    } catch {
        Write-Host "Could not find manager with email: $managerEmail" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        $retry = Read-Host "`nTry again? (y/n)"
        if ($retry -notmatch "[yY]") {
            Write-Host "Operation cancelled. Exiting script." -ForegroundColor Red
            Disconnect-MgGraph | Out-Null
            return
        }
    }
} while (-not $validManager)

# Confirm operation
Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "CONFIRMATION REQUIRED" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Street Address: $targetStreet" -ForegroundColor Cyan
Write-Host "New Manager: $($manager.DisplayName) ($($manager.UserPrincipalName))" -ForegroundColor Cyan
Write-Host "Users to update: $($usersToUpdate.Count)" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Yellow

$ready = $false
while (-not $ready) {
    $confirm = Read-Host "Proceed with update? (y = yes, n = no, l = list users)"
    switch ($confirm.ToLower()) {
        'y' {
            $ready = $true
        }
        'n' {
            Write-Host "Operation cancelled. Exiting script." -ForegroundColor Red
            Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
            Disconnect-MgGraph | Out-Null
            Write-Host "Disconnected. Script finished." -ForegroundColor Green
            Start-Sleep -Seconds 2
            return
        }
        'l' {
            Write-Host "`nFetching current manager information..." -ForegroundColor Yellow
            $userList = @()
            foreach ($u in $usersToUpdate) {
                try {
                    $currentManager = Get-MgUserManager -UserId $u.Id -ErrorAction SilentlyContinue
                    $managerName = if ($currentManager) { 
                        $currentManager.AdditionalProperties.displayName 
                    } else { 
                        "(No manager set)" 
                    }
                } catch {
                    $managerName = "(No manager set)"
                }
                
                $userList += [PSCustomObject]@{
                    UserPrincipalName = $u.UserPrincipalName
                    DisplayName = $u.DisplayName
                    StreetAddress = $u.StreetAddress
                    CurrentManager = $managerName
                }
            }
            Write-Host "`nUsers to be updated:" -ForegroundColor Cyan
            $userList | Format-Table -AutoSize
            Write-Host ""
        }
        default {
            Write-Host "Please enter y (yes), n (no), or l (list users)." -ForegroundColor Yellow
        }
    }
}

# Update users
Write-Host "`nStarting update process..." -ForegroundColor Cyan
$successCount = 0
$errorCount = 0
$errors = @()

foreach ($user in $usersToUpdate) {
    try {
        Write-Host "Updating $($user.UserPrincipalName)..." -ForegroundColor Yellow
        
        # Set the manager reference
        $managerReference = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($manager.Id)"
        }
        
        Set-MgUserManagerByRef -UserId $user.Id -BodyParameter $managerReference
        
        Write-Host "  ✓ Successfully updated $($user.DisplayName)" -ForegroundColor Green
        $successCount++
    } catch {
        Write-Host "  ✗ Failed to update $($user.DisplayName): $($_.Exception.Message)" -ForegroundColor Red
        $errorCount++
        $errors += [PSCustomObject]@{
            User = $user.UserPrincipalName
            Error = $_.Exception.Message
        }
    }
}

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "UPDATE SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Successfully updated: $successCount users" -ForegroundColor Green
if ($errorCount -gt 0) {
    Write-Host "Failed updates: $errorCount users" -ForegroundColor Red
    Write-Host "`nErrors:" -ForegroundColor Red
    $errors | Format-Table -AutoSize
}
Write-Host "========================================`n" -ForegroundColor Cyan

# Disconnect
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected. Script finished." -ForegroundColor Green
Start-Sleep -Seconds 2
