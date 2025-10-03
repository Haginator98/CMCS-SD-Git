# Import-Module ExchangeOnlineManagement
# Connect-ExchangeOnline

# Define the user you want to check
$user = Read-Host "Enter the user (UPN or alias)"

# Get only shared mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox
$total = $mailboxes.Count
$counter = 0

# Arrays to store results
$fullAccess = @()
$sendAs = @()

foreach ($mb in $mailboxes) {
    $counter++

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

# Merge results
$results = $fullAccess + $sendAs

# Output
if ($results) {
    $results | Format-Table -AutoSize
} else {
    Write-Host "No shared mailboxes found for $user" -ForegroundColor Yellow
}