# Import-Module ExchangeOnlineManagement
# Connect-ExchangeOnline

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


Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline | Out-Null
Write-Host "This script will help you find mailboxes based on an alias." -ForegroundColor Cyan

# Define the user to check permissions for
$user = Read-Host "Enter the user (alias, UPN, or email) to get shared mailbox access"

# Get only shared mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox
$total = $mailboxes.Count
$counter = 0

# Arrays to store results
$fullAccess = @()
$sendAs = @()

foreach ($mb in $mailboxes) {
    # Show progress bar
    Write-Progress -Activity "Checking shared mailbox permissions" `
                   -Status "Mailbox: $($mb.Identity)" `
                   -PercentComplete (($counter / $total) * 100)

    # --- FullAccess ---
    $fa = Get-MailboxPermission -Identity $mb.Identity -ErrorAction SilentlyContinue |
        Where-Object { $_.User.ToString() -eq $user -and $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited } |
        Select-Object @{Name="Mailbox";Expression={$_.Identity}},
                      @{Name="User";Expression={$_.User}},
                      @{Name="AccessType";Expression={"FullAccess"}}

    if ($fa) { $fullAccess += $fa }

    # --- SendAs ---
    $sa = Get-RecipientPermission -Identity $mb.Identity -ErrorAction SilentlyContinue |
        Where-Object { $_.Trustee -eq $user -and -not $_.IsInherited } |
        Select-Object @{Name="Mailbox";Expression={$_.Identity}},
                      @{Name="User";Expression={$_.Trustee}},
                      @{Name="AccessType";Expression={"SendAs"}}

    if ($sa) { $sendAs += $sa }
}
Write-Progress -Activity "Checking shared mailbox permissions" -Completed

# Merge results
$results = $fullAccess + $sendAs

# Output
if ($results) {
    $results | Format-Table -AutoSize

    # Use cross-platform Desktop path
    $desktop = [Environment]::GetFolderPath('Desktop')
    $csvPath = "$desktop/SMB_$($user)_$(Get-Date -Format 'yyyyMMdd').csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Start-Sleep -Seconds 2  # Wait to ensure export is complete
    Write-Host "Results exported to $csvPath" -ForegroundColor Green
    Start-Sleep -Seconds 4
    Write-Host "Script will now disconnect from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected. Script finished." -ForegroundColor Green
    Start-Sleep -Seconds 1
    Clear-Host
} else {
    Write-Host "No shared mailboxes found for $user" -ForegroundColor Yellow
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected. Script finished." -ForegroundColor Green
    Start-Sleep -Seconds 2
    Clear-Host
}
