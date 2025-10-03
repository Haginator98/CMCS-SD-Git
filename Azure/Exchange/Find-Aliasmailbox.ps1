# Requires Exchange Online PowerShell V2 module
# Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser

# Prompt for email and password (for interactive login)
$UserCredential = Get-Credential -Message "Enter your Exchange admin credentials"

# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName $UserCredential.UserName -ShowProgress $true

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
Disconnect-ExchangeOnline -Confirm:$false
