# Requires Exchange Online PowerShell V2 module
# Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline
Write-Host "This script will help you find mailboxes based on an alias." -ForegroundColor Cyan
# Prompt for the alias to search
$Alias = Read-Host "Enter the alias address to search (without domain, e.g. 'jdoe')"

# Search for the mailbox with the alias
$Mailboxes = Get-Recipient -ResultSize Unlimited | Where-Object {
    $_.EmailAddresses -match "SMTP:$Alias@"
}

if ($Mailboxes) {
    Write-Host "Mailboxes matching alias '$Alias':"
    $Mailboxes | Select-Object Name,RecipientType,EmailAddresses | Format-Table -AutoSize
} else {
    Write-Host "No mailbox found with alias '$Alias'"
}

# Disconnect from Exchange Online
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected. Script finished." -ForegroundColor Green
Start-Sleep -Seconds 2
Clear-Host