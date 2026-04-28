# Grant a user access to another user's calendar in Exchange Online
# Import required modules (installed via Tools.ps1)
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false
Write-Host "Connected!" -ForegroundColor Green

try {
    # --- Mailbox owner(s) (whose calendar we are sharing) ---
    $ownerInput = Read-Host "`nEnter UPN(s) of mailbox OWNER(s) (whose calendar will be shared, separated by ',')"
    $ownerList = $ownerInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

    if ($ownerList.Count -eq 0) {
        Write-Host "No valid UPNs entered. Aborting." -ForegroundColor Red
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Exit
    }

    Write-Host "`nValidating $($ownerList.Count) owner mailbox(es)..." -ForegroundColor Cyan
    $owners = @()
    $invalidOwners = @()

    foreach ($upn in $ownerList) {
        try {
            $mbx = Get-Mailbox -Identity $upn -ErrorAction Stop
            $owners += $mbx
            Write-Host "  [OK] $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress))" -ForegroundColor Green
        }
        catch {
            Write-Host "  [FAIL] '$upn' - mailbox not found" -ForegroundColor Red
            $invalidOwners += $upn
        }
    }

    if ($owners.Count -eq 0) {
        Write-Host "`nNo valid owner mailboxes found. Aborting." -ForegroundColor Red
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Exit
    }

    if ($invalidOwners.Count -gt 0) {
        Write-Host "`n$($invalidOwners.Count) owner UPN(s) could not be resolved and will be skipped." -ForegroundColor Yellow
    }

    # --- Delegate user(s) (who will get access) ---
    $delegateInput = Read-Host "`nEnter UPN(s) of user(s) to GRANT access (separated by ',')"
    $delegateList = $delegateInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

    if ($delegateList.Count -eq 0) {
        Write-Host "No valid UPNs entered. Aborting." -ForegroundColor Red
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Exit
    }

    Write-Host "`nValidating $($delegateList.Count) delegate user(s)..." -ForegroundColor Cyan
    $delegates = @()
    $invalidDelegates = @()

    foreach ($upn in $delegateList) {
        try {
            $user = Get-Mailbox -Identity $upn -ErrorAction Stop
            $delegates += $user
            Write-Host "  [OK] $($user.DisplayName) ($($user.PrimarySmtpAddress))" -ForegroundColor Green
        }
        catch {
            Write-Host "  [FAIL] '$upn' - mailbox not found" -ForegroundColor Red
            $invalidDelegates += $upn
        }
    }

    if ($delegates.Count -eq 0) {
        Write-Host "`nNo valid delegate mailboxes found. Aborting." -ForegroundColor Red
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Exit
    }

    if ($invalidDelegates.Count -gt 0) {
        Write-Host "`n$($invalidDelegates.Count) delegate UPN(s) could not be resolved and will be skipped." -ForegroundColor Yellow
    }

    # --- Permission level selection ---
    Write-Host "`nSelect the access level the delegate(s) will get on '$($ownerMailbox.DisplayName)' calendar:" -ForegroundColor Yellow
    Write-Host "  [1]  Owner             - Full control (read, create, edit, delete all + folder rights)"
    Write-Host "  [2]  PublishingEditor  - Read, create, edit, delete all items + create subfolders"
    Write-Host "  [3]  Editor            - Read, create, edit, delete all items"
    Write-Host "  [4]  PublishingAuthor  - Read, create items + edit/delete own + create subfolders"
    Write-Host "  [5]  Author            - Read, create items + edit/delete own"
    Write-Host "  [6]  NonEditingAuthor  - Read, create items + delete own"
    Write-Host "  [7]  Reviewer          - Read only (full details)"
    Write-Host "  [8]  Contributor       - Create items only (cannot read)"
    Write-Host "  [9]  LimitedDetails    - Free/busy + subject + location"
    Write-Host "  [10] AvailabilityOnly  - Free/busy only"
    Write-Host "  [11] None              - Remove all access"

    $permissionMap = @{
        "1"  = "Owner"
        "2"  = "PublishingEditor"
        "3"  = "Editor"
        "4"  = "PublishingAuthor"
        "5"  = "Author"
        "6"  = "NonEditingAuthor"
        "7"  = "Reviewer"
        "8"  = "Contributor"
        "9"  = "LimitedDetails"
        "10" = "AvailabilityOnly"
        "11" = "None"
    }

    do {
        $permChoice = Read-Host "`nEnter your choice (1-11)"
        if (-not $permissionMap.ContainsKey($permChoice)) {
            Write-Host "Invalid choice. Please enter a number between 1 and 11." -ForegroundColor Red
        }
    } while (-not $permissionMap.ContainsKey($permChoice))

    $accessRight = $permissionMap[$permChoice]

    # --- Show what will be changed ---
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "The following calendar permission will be applied:" -ForegroundColor Yellow
    Write-Host "  Calendar owner(s) : $($owners.Count) mailbox(es)" -ForegroundColor White
    Write-Host "  Access level      : $accessRight" -ForegroundColor White
    Write-Host "  Delegate(s)       : $($delegates.Count) user(s)" -ForegroundColor White
    Write-Host "  Total operations  : $($owners.Count * $delegates.Count)" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan

    Write-Host "`nCalendar owner(s) being shared:" -ForegroundColor Yellow
    $owners | ForEach-Object { Write-Host "  - $($_.DisplayName) ($($_.PrimarySmtpAddress))" -ForegroundColor White }

    Write-Host "`nDelegates that will receive '$accessRight' access:" -ForegroundColor Yellow
    $delegates | ForEach-Object { Write-Host "  - $($_.DisplayName) ($($_.PrimarySmtpAddress))" -ForegroundColor White }

    Write-Host "`nWould you like to proceed?" -ForegroundColor Yellow
    $confirm = Read-Host "[Y] Yes [N] No"

    if ($confirm -notmatch "^[yY]$") {
        Write-Host "`nOperation cancelled by user." -ForegroundColor Red
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Exit
    }

    # --- Apply changes ---
    Write-Host "`nApplying changes..." -ForegroundColor Cyan

    $total = $owners.Count * $delegates.Count
    $current = 0
    $successCount = 0
    $failedOps = @()

    foreach ($ownerMailbox in $owners) {
        $calPath = "$($ownerMailbox.PrimarySmtpAddress):\Calendar"

        foreach ($delegate in $delegates) {
            $current++
            Write-Progress -Activity "Updating calendar permissions" `
                -Status "Processing $current of $total: $($delegate.DisplayName) on $($ownerMailbox.DisplayName)" `
                -PercentComplete (($current / $total) * 100)

            try {
                $existing = Get-MailboxFolderPermission -Identity $calPath -User $delegate.PrimarySmtpAddress -ErrorAction SilentlyContinue

                if ($accessRight -eq "None") {
                    if ($existing) {
                        Remove-MailboxFolderPermission -Identity $calPath -User $delegate.PrimarySmtpAddress -Confirm:$false -ErrorAction Stop
                        Write-Host "  [OK] Removed $($delegate.DisplayName) from $($ownerMailbox.DisplayName)'s calendar" -ForegroundColor Green
                    }
                    else {
                        Write-Host "  [SKIP] $($delegate.DisplayName) had no existing permission on $($ownerMailbox.DisplayName)" -ForegroundColor DarkGray
                    }
                }
                else {
                    if ($existing) {
                        Set-MailboxFolderPermission -Identity $calPath -User $delegate.PrimarySmtpAddress -AccessRights $accessRight -ErrorAction Stop
                    }
                    else {
                        Add-MailboxFolderPermission -Identity $calPath -User $delegate.PrimarySmtpAddress -AccessRights $accessRight -ErrorAction Stop
                    }
                    Write-Host "  [OK] $($delegate.DisplayName) -> $accessRight on $($ownerMailbox.DisplayName)'s calendar" -ForegroundColor Green
                }

                $successCount++
            }
            catch {
                $failedOps += [PSCustomObject]@{
                    Owner       = $ownerMailbox.DisplayName
                    OwnerEmail  = $ownerMailbox.PrimarySmtpAddress
                    Delegate    = $delegate.DisplayName
                    Email       = $delegate.PrimarySmtpAddress
                    Error       = $_.Exception.Message
                }
                Write-Host "  [FAIL] $($delegate.DisplayName) on $($ownerMailbox.DisplayName)'s calendar" -ForegroundColor Red
                Write-Host "    Reason: $($_.Exception.Message)" -ForegroundColor DarkGray
            }
        }
    }

    Write-Progress -Activity "Updating calendar permissions" -Completed

    # --- Summary ---
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Calendar owners       : $($owners.Count)" -ForegroundColor White
    Write-Host "Delegates             : $($delegates.Count)" -ForegroundColor White
    Write-Host "Access level applied  : $accessRight" -ForegroundColor White
    Write-Host "Total operations      : $total" -ForegroundColor White
    Write-Host "Successfully updated  : $successCount" -ForegroundColor Green
    Write-Host "Failed                : $($failedOps.Count)" -ForegroundColor $(if ($failedOps.Count -gt 0) { "Red" } else { "Green" })

    if ($failedOps.Count -gt 0) {
        Write-Host "`nFailed operations:" -ForegroundColor Yellow
        foreach ($failed in $failedOps) {
            Write-Host "  - $($failed.Delegate) ($($failed.Email)) on $($failed.Owner)'s calendar" -ForegroundColor Red
            Write-Host "    Reason: $($failed.Error)" -ForegroundColor DarkGray
        }
    }
}
finally {
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected." -ForegroundColor Green
}
