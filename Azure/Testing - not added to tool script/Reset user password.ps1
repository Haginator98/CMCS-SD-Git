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

# Write-Host "Importing Microsoft.Graph module..." -ForegroundColor Yellow
# Import-Module Microsoft.Graph

# Start timer
$scriptStartTime = Get-Date

Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome

Write-Host "This script will reset password for a user in Entra ID." -ForegroundColor Cyan
Start-Sleep -Seconds 1

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

# Password options
Write-Host "`nPassword options:" -ForegroundColor Yellow
Write-Host "[1] Generate random password"
Write-Host "[2] Enter custom password"
$passwordOption = Read-Host "Choose option (1 or 2)"

if ($passwordOption -eq "1") {
    # Generate random password
    Add-Type -AssemblyName 'System.Web'
    $newPassword = [System.Web.Security.Membership]::GeneratePassword(12, 2)
    Write-Host "Generated password: $newPassword" -ForegroundColor Cyan
} elseif ($passwordOption -eq "2") {
    $newPassword = Read-Host "Enter new password" -AsSecureString
    $newPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($newPassword))
} else {
    Write-Host "Invalid option. Exiting..." -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# Force change password on next sign-in
$forceChange = Read-Host "Force user to change password at next sign-in? [Y] Yes [N] No"
$forceChangePassword = $forceChange -match "[yY]"

# Confirmation
Write-Host "`n=== CONFIRMATION ===" -ForegroundColor Yellow
Write-Host "User: $($user.DisplayName) ($($user.UserPrincipalName))"
Write-Host "New password: $newPassword"
Write-Host "Force change at next sign-in: $forceChangePassword"

$confirm = Read-Host "`nDo you want to reset the password? [Y] Yes [N] No"
if ($confirm -notmatch "[yY]") {
    Write-Host "Password reset cancelled by user." -ForegroundColor Yellow
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# Reset password
try {
    $passwordProfile = @{
        Password = $newPassword
        ForceChangePasswordNextSignIn = $forceChangePassword
    }
    
    Update-MgUser -UserId $user.Id -PasswordProfile $passwordProfile -ErrorAction Stop
    
    Write-Host "`n=== PASSWORD RESET SUCCESSFUL ===" -ForegroundColor Green
    Write-Host "User: $($user.DisplayName)"
    Write-Host "New password: $newPassword"
    Write-Host "Force change at next sign-in: $forceChangePassword"
    
    if ($passwordOption -eq "1") {
        Write-Host "`nIMPORTANT: Save this password and send it to the user securely!" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "`n=== PASSWORD RESET FAILED ===" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
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
