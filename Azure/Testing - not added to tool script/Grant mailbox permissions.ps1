# Check if ExchangeOnlineManagement module is installed
$Module = Get-Module -Name ExchangeOnlineManagement -ListAvailable
if ($Module.Count -eq 0) {
    Write-Host "ExchangeOnlineManagement module is not available." -ForegroundColor Yellow
    $Confirm = Read-Host "Are you sure you want to install the module? [Y] Yes [N] No"
    if ($Confirm -match "[yY]") {
        Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Cyan
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
        if ($?) {
            Write-Host "ExchangeOnlineManagement module installed successfully." -ForegroundColor Green
        } else {
            Write-Host "Failed to install ExchangeOnlineManagement module." -ForegroundColor Red
            Exit
        }
    } else {
        Write-Host "ExchangeOnlineManagement module is required. Please install it using Install-Module ExchangeOnlineManagement cmdlet." -ForegroundColor Red
        Exit
    }
}

Write-Host "Importing ExchangeOnlineManagement module..." -ForegroundColor Yellow
Import-Module ExchangeOnlineManagement

# Start timer
$scriptStartTime = Get-Date

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline | Out-Null

Write-Host "This script will grant mailbox permissions to a user." -ForegroundColor Cyan
Start-Sleep -Seconds 1

# Prompt for mailbox
$mailbox = Read-Host "Enter the mailbox email address (target mailbox)"

# Verify mailbox exists
try {
    $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
    Write-Host "Found mailbox: $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress))" -ForegroundColor Green
} catch {
    Write-Host "Error: Mailbox '$mailbox' not found." -ForegroundColor Red
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Prompt for user to grant permissions to
$delegateUser = Read-Host "Enter the email address of the user to grant permissions to"

# Verify delegate user exists
try {
    $delegate = Get-Mailbox -Identity $delegateUser -ErrorAction Stop
    Write-Host "Found user: $($delegate.DisplayName) ($($delegate.PrimarySmtpAddress))" -ForegroundColor Green
} catch {
    Write-Host "Error: User '$delegateUser' not found." -ForegroundColor Red
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Select permission type
Write-Host "`nSelect permission type(s) to grant:" -ForegroundColor Yellow
Write-Host "[1] Full Access (read, delete, manage mailbox)"
Write-Host "[2] Send As (send email as the mailbox owner)"
Write-Host "[3] Send on Behalf (send email on behalf of the mailbox owner)"
Write-Host "[4] All of the above"
$permissionType = Read-Host "Choose option (1-4)"

$grantFullAccess = $false
$grantSendAs = $false
$grantSendOnBehalf = $false

switch ($permissionType) {
    "1" { $grantFullAccess = $true }
    "2" { $grantSendAs = $true }
    "3" { $grantSendOnBehalf = $true }
    "4" { 
        $grantFullAccess = $true
        $grantSendAs = $true
        $grantSendOnBehalf = $true
    }
    default {
        Write-Host "Invalid option. Exiting..." -ForegroundColor Red
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Exit
    }
}

# Confirmation
Write-Host "`n=== CONFIRMATION ===" -ForegroundColor Yellow
Write-Host "Mailbox: $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress))"
Write-Host "User to grant permissions to: $($delegate.DisplayName) ($($delegate.PrimarySmtpAddress))"
Write-Host "Permissions to grant:"
if ($grantFullAccess) { Write-Host "  - Full Access" -ForegroundColor Cyan }
if ($grantSendAs) { Write-Host "  - Send As" -ForegroundColor Cyan }
if ($grantSendOnBehalf) { Write-Host "  - Send on Behalf" -ForegroundColor Cyan }

$confirm = Read-Host "`nDo you want to grant these permissions? [Y] Yes [N] No"
if ($confirm -notmatch "[yY]") {
    Write-Host "Permission grant cancelled by user." -ForegroundColor Yellow
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Grant permissions
$successCount = 0
$failCount = 0
$failedPermissions = @()

Write-Host "`nGranting permissions..." -ForegroundColor Cyan

# Grant Full Access
if ($grantFullAccess) {
    try {
        Add-MailboxPermission -Identity $mailbox -User $delegateUser -AccessRights FullAccess -InheritanceType All -ErrorAction Stop | Out-Null
        Write-Host "✓ Full Access granted successfully" -ForegroundColor Green
        $successCount++
    } catch {
        Write-Host "✗ Failed to grant Full Access" -ForegroundColor Red
        $failCount++
        $failedPermissions += [PSCustomObject]@{
            Permission = "Full Access"
            Error = $_.Exception.Message
        }
    }
}

# Grant Send As
if ($grantSendAs) {
    try {
        Add-RecipientPermission -Identity $mailbox -Trustee $delegateUser -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
        Write-Host "✓ Send As granted successfully" -ForegroundColor Green
        $successCount++
    } catch {
        Write-Host "✗ Failed to grant Send As" -ForegroundColor Red
        $failCount++
        $failedPermissions += [PSCustomObject]@{
            Permission = "Send As"
            Error = $_.Exception.Message
        }
    }
}

# Grant Send on Behalf
if ($grantSendOnBehalf) {
    try {
        Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo @{Add=$delegateUser} -ErrorAction Stop
        Write-Host "✓ Send on Behalf granted successfully" -ForegroundColor Green
        $successCount++
    } catch {
        Write-Host "✗ Failed to grant Send on Behalf" -ForegroundColor Red
        $failCount++
        $failedPermissions += [PSCustomObject]@{
            Permission = "Send on Behalf"
            Error = $_.Exception.Message
        }
    }
}

# Summary
Write-Host "`n=== PERMISSION GRANT SUMMARY ===" -ForegroundColor Cyan
Write-Host "Total permissions processed: $(($grantFullAccess -as [int]) + ($grantSendAs -as [int]) + ($grantSendOnBehalf -as [int]))" -ForegroundColor White
Write-Host "Successfully granted: $successCount" -ForegroundColor Green
Write-Host "Failed: $failCount" -ForegroundColor Red

# List failed permissions
if ($failCount -gt 0) {
    Write-Host "`n=== FAILED PERMISSIONS LIST ===" -ForegroundColor Yellow
    foreach ($failed in $failedPermissions) {
        Write-Host "  - $($failed.Permission)" -ForegroundColor Red
        Write-Host "    Reason: $($failed.Error)" -ForegroundColor DarkGray
    }
}

# Calculate total execution time
$scriptEndTime = Get-Date
$executionTime = $scriptEndTime - $scriptStartTime
$minutes = [math]::Floor($executionTime.TotalMinutes)
$seconds = $executionTime.Seconds

# Disconnect from Exchange Online
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected successfully." -ForegroundColor Green

# Display execution time
Write-Host "`n=== SCRIPT COMPLETED ===" -ForegroundColor Cyan
if ($minutes -gt 0) {
    Write-Host "Total execution time: $minutes minutes and $seconds seconds" -ForegroundColor White
} else {
    Write-Host "Total execution time: $seconds seconds" -ForegroundColor White
}
Start-Sleep -Seconds 3
