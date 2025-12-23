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
        Write-Host "Microsoft.Graph module is required. Please install it using Install-Module Microsoft.Graph cmdlet." -ForegroundColor Red
        Exit
    }
}

Write-Host "Importing Microsoft.Graph module..." -ForegroundColor Yellow
Import-Module Microsoft.Graph

# Start timer
$scriptStartTime = Get-Date

Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All", "UserAuthenticationMethod.ReadWrite.All" -NoWelcome

Write-Host "This script will reset MFA (Multi-Factor Authentication) methods for a user in Entra ID." -ForegroundColor Cyan
Write-Host "WARNING: This will remove all MFA methods and sign the user out of all sessions." -ForegroundColor Yellow
Start-Sleep -Seconds 2

# Prompt for user
$userUPN = Read-Host "Enter the user's email address (UPN)"

# Verify user exists
try {
    $user = Get-MgUser -Filter "userPrincipalName eq '$userUPN'" -ErrorAction Stop
    
    if (-not $user) {
        Write-Host "Error: User '$userUPN' not found in Entra ID." -ForegroundColor Red
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Exit
    }
    
    Write-Host "Found user: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Green
} catch {
    Write-Host "Error: Failed to find user. $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# Get current MFA methods
Write-Host "`nRetrieving current MFA methods..." -ForegroundColor Cyan
try {
    $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction Stop
    
    if ($authMethods.Count -gt 0) {
        Write-Host "Current MFA methods registered:" -ForegroundColor Yellow
        foreach ($method in $authMethods) {
            $methodType = $method.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.', ''
            Write-Host "  - $methodType (ID: $($method.Id))" -ForegroundColor White
        }
    } else {
        Write-Host "No MFA methods currently registered." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Warning: Could not retrieve MFA methods. $($_.Exception.Message)" -ForegroundColor Yellow
}

# Reset options
Write-Host "`nReset options:" -ForegroundColor Yellow
Write-Host "[1] Revoke MFA sessions (force re-registration on next sign-in)"
Write-Host "[2] Delete all MFA methods (removes phone, authenticator app, etc.)"
Write-Host "[3] Both (recommended - complete reset)"
$resetOption = Read-Host "Choose option (1-3)"

$revokeSession = $false
$deleteMethods = $false

switch ($resetOption) {
    "1" { $revokeSession = $true }
    "2" { $deleteMethods = $true }
    "3" { 
        $revokeSession = $true
        $deleteMethods = $true
    }
    default {
        Write-Host "Invalid option. Exiting..." -ForegroundColor Red
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Exit
    }
}

# Confirmation
Write-Host "`n=== CONFIRMATION ===" -ForegroundColor Yellow
Write-Host "User: $($user.DisplayName) ($($user.UserPrincipalName))"
Write-Host "Actions to perform:"
if ($revokeSession) { Write-Host "  - Revoke MFA sessions (sign out all devices)" -ForegroundColor Cyan }
if ($deleteMethods) { Write-Host "  - Delete all MFA methods" -ForegroundColor Cyan }
Write-Host "`nWARNING: User will need to set up MFA again on next sign-in!" -ForegroundColor Red

$confirm = Read-Host "`nDo you want to reset MFA for this user? [Y] Yes [N] No"
if ($confirm -notmatch "[yY]") {
    Write-Host "MFA reset cancelled by user." -ForegroundColor Yellow
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# Perform reset
$successCount = 0
$failCount = 0
$failedActions = @()

Write-Host "`nResetting MFA..." -ForegroundColor Cyan

# Revoke MFA sessions
if ($revokeSession) {
    try {
        Revoke-MgUserSignInSession -UserId $user.Id -ErrorAction Stop
        Write-Host "✓ MFA sessions revoked successfully (user signed out)" -ForegroundColor Green
        $successCount++
    } catch {
        Write-Host "✗ Failed to revoke MFA sessions" -ForegroundColor Red
        $failCount++
        $failedActions += [PSCustomObject]@{
            Action = "Revoke MFA Sessions"
            Error = $_.Exception.Message
        }
    }
}

# Delete MFA methods
if ($deleteMethods) {
    try {
        $methodsToDelete = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction Stop
        
        if ($methodsToDelete.Count -eq 0) {
            Write-Host "ℹ No MFA methods to delete" -ForegroundColor Yellow
        } else {
            $deletedCount = 0
            $deleteFailCount = 0
            
            foreach ($method in $methodsToDelete) {
                try {
                    # Skip password authentication method (cannot be deleted)
                    $methodType = $method.AdditionalProperties.'@odata.type'
                    if ($methodType -eq '#microsoft.graph.passwordAuthenticationMethod') {
                        continue
                    }
                    
                    Remove-MgUserAuthenticationMethod -UserId $user.Id -AuthenticationMethodId $method.Id -ErrorAction Stop
                    $deletedCount++
                } catch {
                    $deleteFailCount++
                }
            }
            
            if ($deletedCount -gt 0) {
                Write-Host "✓ Deleted $deletedCount MFA method(s)" -ForegroundColor Green
                $successCount++
            }
            
            if ($deleteFailCount -gt 0) {
                Write-Host "⚠ Failed to delete $deleteFailCount method(s)" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "✗ Failed to delete MFA methods" -ForegroundColor Red
        $failCount++
        $failedActions += [PSCustomObject]@{
            Action = "Delete MFA Methods"
            Error = $_.Exception.Message
        }
    }
}

# Summary
Write-Host "`n=== MFA RESET SUMMARY ===" -ForegroundColor Cyan
Write-Host "User: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor White

if ($revokeSession) {
    Write-Host "MFA sessions: Revoked (user signed out from all devices)" -ForegroundColor Green
}

if ($deleteMethods) {
    Write-Host "MFA methods: Deleted" -ForegroundColor Green
}

if ($failCount -eq 0) {
    Write-Host "`n✓ MFA reset completed successfully!" -ForegroundColor Green
    Write-Host "NEXT STEPS:" -ForegroundColor Yellow
    Write-Host "  1. User will be prompted to set up MFA on next sign-in" -ForegroundColor White
    Write-Host "  2. User may need to verify identity via admin approval" -ForegroundColor White
    Write-Host "  3. Inform user to have their phone/authenticator app ready" -ForegroundColor White
} else {
    Write-Host "`n⚠ MFA reset completed with errors" -ForegroundColor Yellow
}

# List failed actions
if ($failCount -gt 0) {
    Write-Host "`n=== FAILED ACTIONS LIST ===" -ForegroundColor Yellow
    foreach ($failed in $failedActions) {
        Write-Host "  - $($failed.Action)" -ForegroundColor Red
        Write-Host "    Reason: $($failed.Error)" -ForegroundColor DarkGray
    }
}

# Calculate total execution time
$scriptEndTime = Get-Date
$executionTime = $scriptEndTime - $scriptStartTime
$minutes = [math]::Floor($executionTime.TotalMinutes)
$seconds = $executionTime.Seconds

# Disconnect from Microsoft Graph
Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected successfully." -ForegroundColor Green

# Display execution time
Write-Host "`n=== SCRIPT COMPLETED ===" -ForegroundColor Cyan
if ($minutes -gt 0) {
    Write-Host "Total execution time: $minutes minutes and $seconds seconds" -ForegroundColor White
} else {
    Write-Host "Total execution time: $seconds seconds" -ForegroundColor White
}
Start-Sleep -Seconds 3
