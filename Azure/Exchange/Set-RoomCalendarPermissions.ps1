# Set calendar permissions on room mailboxes to show organizer and subject (non-anonymous)
# Import required modules (installed via Tools.ps1)
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false
Write-Host "Connected!" -ForegroundColor Green

try {
    # --- Select target method ---
    Write-Host "`nHow would you like to select room mailboxes?" -ForegroundColor Yellow
    Write-Host "[1] By domain (e.g. @kibo.dk)"
    Write-Host "[2] Manual input (UPN separated by ',')"
    $selectionMode = Read-Host "`nEnter your choice (1 or 2)"

    $roomMailboxes = @()

    # --- Domain selection ---
    if ($selectionMode -eq "1") {
        $allRooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited -ErrorAction Stop

        do {
            $domainInput = Read-Host "`nEnter domain to filter on (e.g. contoso.com)"
            $domainInput = $domainInput.TrimStart("@").Trim()

            Write-Host "`nSearching for room mailboxes in domain '@$domainInput'..." -ForegroundColor Cyan
            $roomMailboxes = $allRooms | Where-Object { $_.PrimarySmtpAddress -like "*@$domainInput" }

            if ($roomMailboxes.Count -eq 0) {
                Write-Host "No room mailboxes found for domain '@$domainInput'. Please try again." -ForegroundColor Red
            }
        } while ($roomMailboxes.Count -eq 0)

        Write-Host "Found $($roomMailboxes.Count) room mailbox(es) in '@$domainInput':" -ForegroundColor Green
        $roomMailboxes | ForEach-Object { Write-Host "  - $($_.DisplayName) ($($_.PrimarySmtpAddress))" -ForegroundColor White }
    }
    # --- Manual UPN input ---
    elseif ($selectionMode -eq "2") {
        $manualInput = Read-Host "`nEnter room mailbox UPNs (separated by ',')"
        $upnList = $manualInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

        if ($upnList.Count -eq 0) {
            Write-Host "No valid UPNs entered." -ForegroundColor Red
            Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
            Disconnect-ExchangeOnline -Confirm:$false
            Exit
        }

        Write-Host "`nValidating $($upnList.Count) UPN(s)..." -ForegroundColor Cyan
        $invalidRooms = @()

        foreach ($upn in $upnList) {
            try {
                $room = Get-Mailbox -Identity $upn -RecipientTypeDetails RoomMailbox -ErrorAction Stop
                $roomMailboxes += $room
                Write-Host "  [OK] $($room.DisplayName) ($($room.PrimarySmtpAddress))" -ForegroundColor Green
            }
            catch {
                Write-Host "  [FAIL] '$upn' - not found or not a room mailbox" -ForegroundColor Red
                $invalidRooms += $upn
            }
        }

        if ($roomMailboxes.Count -eq 0) {
            Write-Host "`nNo valid room mailboxes found. Aborting." -ForegroundColor Red
            Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
            Disconnect-ExchangeOnline -Confirm:$false
            Exit
        }

        if ($invalidRooms.Count -gt 0) {
            Write-Host "`n$($invalidRooms.Count) UPN(s) could not be resolved and will be skipped." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "Invalid choice. Aborting." -ForegroundColor Red
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Exit
    }

    # --- Show what will be changed ---
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "The following changes will be applied to $($roomMailboxes.Count) room mailbox(es):" -ForegroundColor Yellow
    Write-Host "  - Calendar folder permission (Default user): AvailabilityOnly -> LimitedDetails" -ForegroundColor White
    Write-Host "  - CalendarProcessing: DeleteSubject = false" -ForegroundColor White
    Write-Host "  - CalendarProcessing: DeleteComments = false" -ForegroundColor White
    Write-Host "  - CalendarProcessing: AddOrganizerToSubject = true" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "`nRooms that will be updated:" -ForegroundColor Yellow
    $roomMailboxes | ForEach-Object { Write-Host "  - $($_.DisplayName) ($($_.PrimarySmtpAddress))" -ForegroundColor White }

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

    $total = $roomMailboxes.Count
    $current = 0
    $successCount = 0
    $failedRooms = @()

    foreach ($room in $roomMailboxes) {
        $current++
        Write-Progress -Activity "Updating room calendar permissions" `
            -Status "Processing $current of $total: $($room.DisplayName)" `
            -PercentComplete (($current / $total) * 100)

        $roomErrors = @()

        # Set calendar folder permission for Default user
        try {
            $calPath = "$($room.PrimarySmtpAddress):\Calendar"
            $existing = Get-MailboxFolderPermission -Identity $calPath -User Default -ErrorAction SilentlyContinue

            if ($existing) {
                Set-MailboxFolderPermission -Identity $calPath -User Default -AccessRights LimitedDetails -ErrorAction Stop
            }
            else {
                Add-MailboxFolderPermission -Identity $calPath -User Default -AccessRights LimitedDetails -ErrorAction Stop
            }
        }
        catch {
            $roomErrors += "FolderPermission: $($_.Exception.Message)"
        }

        # Set CalendarProcessing
        try {
            Set-CalendarProcessing -Identity $room.Identity `
                -DeleteSubject $false `
                -DeleteComments $false `
                -AddOrganizerToSubject $true `
                -ErrorAction Stop
        }
        catch {
            $roomErrors += "CalendarProcessing: $($_.Exception.Message)"
        }

        if ($roomErrors.Count -eq 0) {
            $successCount++
            Write-Host "  [OK] $($room.DisplayName) ($($room.PrimarySmtpAddress))" -ForegroundColor Green
        }
        else {
            $failedRooms += [PSCustomObject]@{
                DisplayName = $room.DisplayName
                Email       = $room.PrimarySmtpAddress
                Errors      = $roomErrors -join " | "
            }
            Write-Host "  [FAIL] $($room.DisplayName) ($($room.PrimarySmtpAddress))" -ForegroundColor Red
            foreach ($err in $roomErrors) {
                Write-Host "    Reason: $err" -ForegroundColor DarkGray
            }
        }
    }

    Write-Progress -Activity "Updating room calendar permissions" -Completed

    # --- Summary ---
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Total rooms processed : $total" -ForegroundColor White
    Write-Host "Successfully updated  : $successCount" -ForegroundColor Green
    Write-Host "Failed                : $($failedRooms.Count)" -ForegroundColor $(if ($failedRooms.Count -gt 0) { "Red" } else { "Green" })

    if ($failedRooms.Count -gt 0) {
        Write-Host "`nFailed rooms:" -ForegroundColor Yellow
        foreach ($failed in $failedRooms) {
            Write-Host "  - $($failed.DisplayName) ($($failed.Email))" -ForegroundColor Red
            Write-Host "    Reason: $($failed.Errors)" -ForegroundColor DarkGray
        }
    }
}
finally {
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected." -ForegroundColor Green
}
