# Requires Exchange Online PowerShell V2 module
# Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser

# Check if ExchangeOnlineManagement module is installed
$Module = Get-Module -Name ExchangeOnlineManagement -ListAvailable
if ($Module.Count -eq 0) {
    Write-Host ExchangeOnlineManagement module is not available -ForegroundColor yellow
    $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No"
    if ($Confirm -match "[yY]") {
        Install-Module ExchangeOnlineManagement
    } else {
        Write-Host ExchangeOnlineManagement module is required. Please install module using Install-Module ExchangeOnlineManagement cmdlet.
        Exit
    }
}
Write-Host Importing ExchangeOnlineManagement module... -ForegroundColor Yellow
Import-Module ExchangeOnlineManagement

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
    $null = Read-Host "Press Enter to confirm you have read the information"
} else {
    Write-Host "No mailbox found with alias '$Alias'"
    Start-Sleep -Seconds 2
}

# Disconnect from Exchange Online
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected. Script finished." -ForegroundColor Green
Start-Sleep -Seconds 3
Clear-Host