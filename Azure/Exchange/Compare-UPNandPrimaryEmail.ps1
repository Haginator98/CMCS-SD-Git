# Compare-UPNandPrimaryEmail.ps1
# Compares UPN with Primary Email address and displays aliases for mismatched users

[CmdletBinding()]
param ()

# Check if connected to Exchange Online
try {
    $null = Get-OrganizationConfig -ErrorAction Stop
    Write-Host "Connected to Exchange Online" -ForegroundColor Green
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
}
catch {
    Write-Host "Not connected to Exchange Online. Connecting..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "Successfully connected to Exchange Online" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] Failed to connect to Exchange Online: $_" -ForegroundColor Red
        exit 1
    }
}

Write-Host "`nRetrieving all mailboxes..." -ForegroundColor Cyan
Write-Progress -Activity "Fetching mailboxes" -Status "Please wait..."

# Get all mailboxes with relevant properties
$mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress, Alias, EmailAddresses

Write-Progress -Activity "Fetching mailboxes" -Completed
Write-Host "Found $($mailboxes.Count) mailboxes. Analyzing..." -ForegroundColor Cyan

# Filter mailboxes where UPN doesn't match Primary SMTP
$mismatchedUsers = @()
$counter = 0

foreach ($mailbox in $mailboxes) {
    $counter++
    Write-Progress -Activity "Analyzing mailboxes" -Status "Processing $counter of $($mailboxes.Count)" -PercentComplete (($counter / $mailboxes.Count) * 100)
    
    # Compare UPN with Primary SMTP (case-insensitive)
    if ($mailbox.UserPrincipalName -ne $mailbox.PrimarySmtpAddress) {
        
        # Extract all email aliases (SMTP addresses)
        $smtpAliases = $mailbox.EmailAddresses | Where-Object { $_ -like "smtp:*" } | ForEach-Object { $_ -replace "smtp:", "" }
        $allAliases = if ($smtpAliases) { $smtpAliases -join "; " } else { "no aliases" }
        
        $mismatchedUsers += [PSCustomObject]@{
            DisplayName       = $mailbox.DisplayName
            UserPrincipalName = $mailbox.UserPrincipalName
            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
            Alias             = if ($mailbox.Alias) { $mailbox.Alias } else { "no alias" }
            AllEmailAliases   = $allAliases
        }
    }
}

Write-Progress -Activity "Analyzing mailboxes" -Completed

# Display results
Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "Users with UPN != Primary Email" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Found $($mismatchedUsers.Count) users with mismatched UPN and Primary Email`n" -ForegroundColor Cyan

if ($mismatchedUsers.Count -gt 0) {
    $mismatchedUsers | Format-Table -AutoSize -Property DisplayName, UserPrincipalName, PrimarySmtpAddress, Alias
    
    Write-Host "`nDetailed view with all aliases:" -ForegroundColor Yellow
    foreach ($user in $mismatchedUsers) {
        Write-Host "`n-----------------------------------" -ForegroundColor Gray
        Write-Host "Display Name:       $($user.DisplayName)" -ForegroundColor White
        Write-Host "UPN:                $($user.UserPrincipalName)" -ForegroundColor Cyan
        Write-Host "Primary Email:      $($user.PrimarySmtpAddress)" -ForegroundColor Green
        Write-Host "Alias:              $($user.Alias)" -ForegroundColor Magenta
        Write-Host "All Email Aliases:  $($user.AllEmailAliases)" -ForegroundColor DarkGray
    }
    
    # Ask for confirmation before exporting
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "EXPORT OPTION" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    $exportConfirmation = Read-Host "Do you want to export these $($mismatchedUsers.Count) users to CSV? [Y] Yes [N] No"
    
    if ($exportConfirmation -match "[yY]") {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $exportPath = Join-Path -Path $desktopPath -ChildPath "UPN-Email-Mismatch_$timestamp.csv"
        
        try {
            $mismatchedUsers | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
            Write-Host "[SUCCESS] Results exported to: $exportPath" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERROR] Failed to export to CSV: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Export cancelled by user." -ForegroundColor Yellow
    }
}
else {
    Write-Host "All users have matching UPN and Primary Email addresses." -ForegroundColor Green
}

# Disconnect from Exchange Online
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

Write-Host "`n========================================" -ForegroundColor Yellow
