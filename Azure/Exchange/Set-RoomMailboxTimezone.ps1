# Get and change the timezone (and regional settings) for a Room Mailbox
# Uses Microsoft Graph only (no ExchangeOnlineManagement)
# Required modules installed via Tools.ps1
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All", "MailboxSettings.ReadWrite" -NoWelcome
Write-Host "Connected!" -ForegroundColor Green

try {
    # Get room mailbox email from user
    $RoomEmail = Read-Host "`nEnter room mailbox email address"

    if ([string]::IsNullOrWhiteSpace($RoomEmail)) {
        Write-Host "No email address provided." -ForegroundColor Red
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        return
    }

    # Verify the user/mailbox exists
    try {
        Write-Host "`nLooking up mailbox: $RoomEmail..." -ForegroundColor Cyan
        $user = Get-MgUser -UserId $RoomEmail -Property "Id,DisplayName,UserPrincipalName,Mail" -ErrorAction Stop
        Write-Host "Found mailbox: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Green
    }
    catch {
        Write-Host "Error: Could not find a user/mailbox with address '$RoomEmail'." -ForegroundColor Red
        Write-Host "Verify that the address is correct." -ForegroundColor Yellow
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        return
    }

    # Get current mailbox settings (regional configuration)
    try {
        Write-Host "Getting current mailbox/regional settings..." -ForegroundColor Cyan
        $settings = Get-MgUserMailboxSetting -UserId $user.Id -ErrorAction Stop
    }
    catch {
        Write-Host "Error retrieving mailbox settings: $_" -ForegroundColor Red
        Write-Host "The account may not have a mailbox, or you may lack 'MailboxSettings.ReadWrite' permission." -ForegroundColor Yellow
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        return
    }

    $currentTZ       = $settings.TimeZone
    $currentLocale   = if ($settings.Language) { $settings.Language.Locale }      else { $null }
    $currentLangName = if ($settings.Language) { $settings.Language.DisplayName } else { $null }
    $currentDateFmt  = $settings.DateFormat
    $currentTimeFmt  = $settings.TimeFormat

    # Display current settings
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "CURRENT REGIONAL SETTINGS" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Mailbox    : $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor White
    Write-Host "TimeZone   : $(if ($currentTZ)      { $currentTZ }                          else { 'not set' })" -ForegroundColor Yellow
    Write-Host "Language   : $(if ($currentLocale)  { "$currentLangName ($currentLocale)" } else { 'not set' })" -ForegroundColor Yellow
    Write-Host "DateFormat : $(if ($currentDateFmt) { $currentDateFmt }                     else { 'not set' })" -ForegroundColor Yellow
    Write-Host "TimeFormat : $(if ($currentTimeFmt) { $currentTimeFmt }                     else { 'not set' })" -ForegroundColor Yellow

    # Ask if the user wants to change the timezone
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "CHANGE TIMEZONE" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    $change = Read-Host "Do you want to change the timezone for this mailbox? [Y] Yes [N] No"
    if ($change -notmatch "[yY]") {
        Write-Host "No changes made. Operation cancelled by user." -ForegroundColor Yellow
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        return
    }

    # Preset list of common timezones (Windows IDs - default format used by Graph)
    $presets = @(
        "W. Europe Standard Time"      # Norway, Sweden, Denmark, Germany, Netherlands
        "GMT Standard Time"            # United Kingdom, Ireland
        "Central Europe Standard Time" # Poland, Czech Republic, Hungary
        "FLE Standard Time"            # Finland, Baltic states
        "Romance Standard Time"        # France, Spain
        "UTC"
        "Eastern Standard Time"        # US East
        "Pacific Standard Time"        # US West
    )

    Write-Host "`nSelect a timezone:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $presets.Count; $i++) {
        Write-Host "[$($i + 1)] $($presets[$i])" -ForegroundColor White
    }
    Write-Host "[C] Custom (enter timezone ID manually)" -ForegroundColor White
    Write-Host "[L] List all available timezones" -ForegroundColor White
    Write-Host "[0] Cancel" -ForegroundColor Yellow

    $choice = Read-Host "`nEnter choice"

    $newTZ = $null
    switch -Regex ($choice) {
        '^0$' {
            Write-Host "No changes made. Operation cancelled by user." -ForegroundColor Yellow
            Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
            return
        }
        '^[lL]$' {
            Write-Host "`nAvailable Windows timezone IDs:" -ForegroundColor Cyan
            [System.TimeZoneInfo]::GetSystemTimeZones() |
                Select-Object Id, DisplayName |
                Sort-Object Id |
                Format-Table -AutoSize | Out-Host
            $newTZ = Read-Host "Enter the timezone ID (e.g. 'W. Europe Standard Time')"
        }
        '^[cC]$' {
            $newTZ = Read-Host "Enter the timezone ID (e.g. 'W. Europe Standard Time')"
        }
        '^\d+$' {
            $idx = [int]$choice - 1
            if ($idx -ge 0 -and $idx -lt $presets.Count) {
                $newTZ = $presets[$idx]
            } else {
                Write-Host "Invalid choice." -ForegroundColor Red
                Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
                return
            }
        }
        default {
            Write-Host "Invalid choice." -ForegroundColor Red
            Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
            return
        }
    }

    if ([string]::IsNullOrWhiteSpace($newTZ)) {
        Write-Host "No timezone provided. No changes made." -ForegroundColor Yellow
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        return
    }

    # Validate timezone ID against the local system list (best-effort)
    $validTZ = [System.TimeZoneInfo]::GetSystemTimeZones() | Where-Object { $_.Id -eq $newTZ }
    if (-not $validTZ) {
        Write-Host "`n[WARNING] '$newTZ' was not found in the local system timezone list." -ForegroundColor Yellow
        Write-Host "It may still be valid in Microsoft Graph, but please double-check the spelling." -ForegroundColor Yellow
        $continueAnyway = Read-Host "Continue anyway? [Y] Yes [N] No"
        if ($continueAnyway -notmatch "[yY]") {
            Write-Host "No changes made. Operation cancelled by user." -ForegroundColor Yellow
            Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
            return
        }
    }

    # Final confirmation
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "CONFIRM CHANGE" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Mailbox    : $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor White
    Write-Host "Current TZ : $(if ($currentTZ) { $currentTZ } else { 'not set' })" -ForegroundColor White
    Write-Host "New TZ     : $newTZ" -ForegroundColor Green

    $confirm = Read-Host "`nApply this change? [Y] Yes [N] No"
    if ($confirm -notmatch "[yY]") {
        Write-Host "No changes made. Operation cancelled by user." -ForegroundColor Yellow
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        return
    }

    try {
        Write-Host "`nUpdating timezone..." -ForegroundColor Cyan
        Update-MgUserMailboxSetting -UserId $user.Id -BodyParameter @{ timeZone = $newTZ } -ErrorAction Stop
        Write-Host "Timezone updated to '$newTZ' for $($user.UserPrincipalName)" -ForegroundColor Green

        # Show updated values
        $updated = Get-MgUserMailboxSetting -UserId $user.Id
        Write-Host "`nUpdated regional settings:" -ForegroundColor Cyan
        Write-Host "TimeZone   : $($updated.TimeZone)" -ForegroundColor Green
        Write-Host "Language   : $(if ($updated.Language)   { "$($updated.Language.DisplayName) ($($updated.Language.Locale))" } else { 'not set' })" -ForegroundColor Green
        Write-Host "DateFormat : $(if ($updated.DateFormat) { $updated.DateFormat } else { 'not set' })" -ForegroundColor Green
        Write-Host "TimeFormat : $(if ($updated.TimeFormat) { $updated.TimeFormat } else { 'not set' })" -ForegroundColor Green
    }
    catch {
        Write-Host "Error updating timezone: $_" -ForegroundColor Red
        Write-Host "The timezone ID may be invalid or you may lack the 'MailboxSettings.ReadWrite' permission." -ForegroundColor Yellow
    }
}
finally {
    Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Write-Host "Disconnected. Script finished." -ForegroundColor Green
    Start-Sleep -Seconds 2
}
