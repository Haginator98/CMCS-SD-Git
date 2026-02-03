# Get and manage settings for Room Mailboxes
# Import required modules (installed via Tools.ps1)
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false
Write-Host "Connected!" -ForegroundColor Green

# Main script logic wrapped in try-finally to ensure disconnect
try {
    # Get room mailbox email from user
    $RoomEmail = Read-Host "`nEnter room mailbox email address"
    
    # Get room mailbox
    try {
        Write-Host "`nGetting room: $RoomEmail..." -ForegroundColor Cyan
        $room = Get-Mailbox -Identity $RoomEmail -RecipientTypeDetails RoomMailbox -ErrorAction Stop
        Write-Host "Found room mailbox: $($room.DisplayName)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error: Could not find room mailbox with email $RoomEmail" -ForegroundColor Red
        Write-Host "Please verify the email address is correct and is a room mailbox." -ForegroundColor Yellow
        return
    }

    # Get settings for the room
    try {
        Write-Host "Getting settings..." -ForegroundColor Cyan
        
        # Get calendar processing settings
        $calSettings = Get-CalendarProcessing -Identity $room.Identity
        
        # Get mailbox settings
        $mailbox = Get-Mailbox -Identity $room.Identity
        
        # Get calendar folder permissions
        $calendarPermissions = Get-MailboxFolderPermission -Identity "$($room.PrimarySmtpAddress):\Calendar" -ErrorAction SilentlyContinue
        
        # Build result object
        $roomInfo = [PSCustomObject]@{
            DisplayName                    = $room.DisplayName
            Email                          = $room.PrimarySmtpAddress
            Alias                          = $room.Alias
            AutomateProcessing            = $calSettings.AutomateProcessing
            AllowConflicts                = $calSettings.AllowConflicts
            ConflictPercentageAllowed     = $calSettings.ConflictPercentageAllowed
            MaximumConflictInstances      = $calSettings.MaximumConflictInstances
            BookingWindowInDays           = $calSettings.BookingWindowInDays
            MaximumDurationInMinutes      = $calSettings.MaximumDurationInMinutes
            TentativePendingApproval      = $calSettings.TentativePendingApproval
            AllowRecurringMeetings        = $calSettings.AllowRecurringMeetings
            ScheduleOnlyDuringWorkHours   = $calSettings.ScheduleOnlyDuringWorkHours
            EnforceSchedulingHorizon      = $calSettings.EnforceSchedulingHorizon
            BookInPolicy                  = if ($calSettings.BookInPolicy) { ($calSettings.BookInPolicy -join "; ") } else { "no policy" }
            RequestInPolicy               = if ($calSettings.RequestInPolicy) { ($calSettings.RequestInPolicy -join "; ") } else { "no policy" }
            RequestOutOfPolicy            = if ($calSettings.RequestOutOfPolicy) { ($calSettings.RequestOutOfPolicy -join "; ") } else { "no policy" }
            AllBookInPolicy               = $calSettings.AllBookInPolicy
            AllRequestInPolicy            = $calSettings.AllRequestInPolicy
            AllRequestOutOfPolicy         = $calSettings.AllRequestOutOfPolicy
            ResourceDelegates             = if ($calSettings.ResourceDelegates) { ($calSettings.ResourceDelegates -join "; ") } else { "no delegates" }
            ForwardRequestsToDelegates    = $calSettings.ForwardRequestsToDelegates
            DeleteAttachments             = $calSettings.DeleteAttachments
            DeleteComments                = $calSettings.DeleteComments
            DeleteSubject                 = $calSettings.DeleteSubject
            AddOrganizerToSubject         = $calSettings.AddOrganizerToSubject
            RemovePrivateProperty         = $calSettings.RemovePrivateProperty
            ProcessExternalMeetingMessages = $calSettings.ProcessExternalMeetingMessages
            AddAdditionalResponse         = $calSettings.AddAdditionalResponse
            AdditionalResponse            = if ($calSettings.AdditionalResponse) { $calSettings.AdditionalResponse } else { "no additional response" }
            City                          = if ($mailbox.City) { $mailbox.City } else { "no city" }
            CountryOrRegion               = if ($mailbox.CountryOrRegion) { $mailbox.CountryOrRegion } else { "no country" }
            Office                        = if ($mailbox.Office) { $mailbox.Office } else { "no office" }
            Phone                         = if ($mailbox.Phone) { $mailbox.Phone } else { "no phone" }
            ResourceCapacity              = if ($mailbox.ResourceCapacity) { $mailbox.ResourceCapacity } else { "not set" }
            CalendarPermissions           = if ($calendarPermissions) { 
                ($calendarPermissions | ForEach-Object { "$($_.User): $($_.AccessRights)" }) -join "; " 
            } else { "default permissions" }
            WhenCreated                   = $mailbox.WhenCreated
            WhenChanged                   = $mailbox.WhenChanged
        }
    }
    catch {
        Write-Host "Error getting settings: $_" -ForegroundColor Red
        return
    }
    
    # Display settings
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "ROOM MAILBOX SETTINGS" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "`n--- $($roomInfo.DisplayName) ---" -ForegroundColor Yellow
    Write-Host "Email: $($roomInfo.Email)" -ForegroundColor White
    
    # Highlight critical settings for conflict handling
    Write-Host "`n** CONFLICT HANDLING **" -ForegroundColor Magenta
    Write-Host "Automate Processing: $($roomInfo.AutomateProcessing)" -ForegroundColor $(if ($roomInfo.AutomateProcessing -eq "AutoAccept") { "Green" } else { "Yellow" })
    Write-Host "Allow Conflicts: $($roomInfo.AllowConflicts)" -ForegroundColor $(if ($roomInfo.AllowConflicts -eq $false) { "Green" } else { "Red" })
    Write-Host "Conflict Percentage Allowed: $($roomInfo.ConflictPercentageAllowed)%" -ForegroundColor White
    Write-Host "Maximum Conflict Instances: $($roomInfo.MaximumConflictInstances)" -ForegroundColor White
    Write-Host "Tentative Pending Approval: $($roomInfo.TentativePendingApproval)" -ForegroundColor $(if ($roomInfo.TentativePendingApproval -eq $true) { "Red" } else { "Green" })
    
    Write-Host "`n** BOOKING SETTINGS **" -ForegroundColor Cyan
    Write-Host "Booking Window (days): $($roomInfo.BookingWindowInDays)" -ForegroundColor White
    Write-Host "Max Duration (min): $($roomInfo.MaximumDurationInMinutes)" -ForegroundColor White
    Write-Host "Allow Recurring Meetings: $($roomInfo.AllowRecurringMeetings)" -ForegroundColor White
    Write-Host "Schedule Only During Work Hours: $($roomInfo.ScheduleOnlyDuringWorkHours)" -ForegroundColor White
    
    Write-Host "`n** PRIVACY SETTINGS **" -ForegroundColor Cyan
    Write-Host "Delete Attachments: $($roomInfo.DeleteAttachments)" -ForegroundColor White
    Write-Host "Delete Comments: $($roomInfo.DeleteComments)" -ForegroundColor White
    Write-Host "Delete Subject: $($roomInfo.DeleteSubject)" -ForegroundColor White
    Write-Host "Process External Meetings: $($roomInfo.ProcessExternalMeetingMessages)" -ForegroundColor White
    
    Write-Host "`n** ACCESS & PERMISSIONS **" -ForegroundColor Cyan
    Write-Host "Capacity: $($roomInfo.ResourceCapacity)" -ForegroundColor White
    
    if ($roomInfo.ResourceDelegates -ne "no delegates") {
        Write-Host "Delegates: $($roomInfo.ResourceDelegates)" -ForegroundColor White
    } else {
        Write-Host "Delegates: no delegates" -ForegroundColor White
    }
    
    if ($roomInfo.BookInPolicy -ne "no policy" -and -not $roomInfo.AllBookInPolicy) {
        Write-Host "Book In Policy: $($roomInfo.BookInPolicy)" -ForegroundColor White
    }
    elseif ($roomInfo.AllBookInPolicy) {
        Write-Host "Book In Policy: Everyone can book directly" -ForegroundColor Green
    } else {
        Write-Host "Book In Policy: no policy (requires approval)" -ForegroundColor Yellow
    }
    
    Write-Host "Calendar Folder Permissions: $($roomInfo.CalendarPermissions)" -ForegroundColor White
    
    Write-Host "`n** OTHER INFO **" -ForegroundColor Cyan
    Write-Host "Location: $($roomInfo.Office), $($roomInfo.City)" -ForegroundColor White
    
    # Add diagnostic warning if configuration might cause double booking
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "DIAGNOSTIC NOTES" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    
    if ($roomInfo.AutomateProcessing -ne "AutoAccept") {
        Write-Host "[WARNING] AutomateProcessing is NOT set to AutoAccept!" -ForegroundColor Red
        Write-Host "  This means the room won't automatically reject conflicting bookings." -ForegroundColor Red
        Write-Host "  Meetings will appear as 'Tentative' until manually processed." -ForegroundColor Red
    }
    
    if ($roomInfo.AllowConflicts -eq $true) {
        Write-Host "[WARNING] AllowConflicts is set to True!" -ForegroundColor Red
        Write-Host "  Multiple overlapping meetings are allowed." -ForegroundColor Red
    }
    
    if ($roomInfo.TentativePendingApproval -eq $true) {
        Write-Host "[INFO] TentativePendingApproval is True" -ForegroundColor Yellow
        Write-Host "  Meetings will show as 'Tentative' until approved by delegate." -ForegroundColor Yellow
    }
    
    if ($roomInfo.ConflictPercentageAllowed -gt 0) {
        Write-Host "[INFO] ConflictPercentageAllowed: $($roomInfo.ConflictPercentageAllowed)%" -ForegroundColor Yellow
        Write-Host "  Some percentage of conflict is allowed." -ForegroundColor Yellow
    }
    
    if ($roomInfo.CalendarPermissions -like "*Editor*" -or $roomInfo.CalendarPermissions -like "*Owner*") {
        Write-Host "[INFO] Users with Editor/Owner permissions can bypass conflict checks" -ForegroundColor Yellow
    }
    
    if ($roomInfo.AutomateProcessing -eq "AutoAccept" -and $roomInfo.AllowConflicts -eq $false -and $roomInfo.TentativePendingApproval -eq $false) {
        Write-Host "[OK] Configuration looks correct for preventing double bookings" -ForegroundColor Green
    }
    
    # Ask if user wants to modify settings
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "MODIFY SETTINGS" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    
    $modifyChoice = Read-Host "Do you want to modify settings for this room? [Y] Yes [N] No"
    if ($modifyChoice -match "[yY]") {
        
        Write-Host "`nModifying settings for: $($roomInfo.DisplayName)" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "[1] AutomateProcessing (Current: $($roomInfo.AutomateProcessing))" -ForegroundColor White
        Write-Host "[2] BookingWindowInDays (Current: $($roomInfo.BookingWindowInDays))" -ForegroundColor White
        Write-Host "[3] MaximumDurationInMinutes (Current: $($roomInfo.MaximumDurationInMinutes))" -ForegroundColor White
        Write-Host "[4] AllowConflicts (Current: $($roomInfo.AllowConflicts))" -ForegroundColor White
        Write-Host "[5] TentativePendingApproval (Current: $($roomInfo.TentativePendingApproval))" -ForegroundColor White
        Write-Host "[6] DeleteAttachments (Current: $($roomInfo.DeleteAttachments))" -ForegroundColor White
        Write-Host "[7] DeleteComments (Current: $($roomInfo.DeleteComments))" -ForegroundColor White
        Write-Host "[8] DeleteSubject (Current: $($roomInfo.DeleteSubject))" -ForegroundColor White
        Write-Host "[9] ProcessExternalMeetingMessages (Current: $($roomInfo.ProcessExternalMeetingMessages))" -ForegroundColor White
        Write-Host "[0] Cancel" -ForegroundColor Yellow
        
        $settingChoice = Read-Host "`nSelect setting to modify (0-9)"
        
        switch ($settingChoice) {
            "1" {
                Write-Host "`nAutomate Processing options:" -ForegroundColor Cyan
                Write-Host "[1] AutoAccept - Automatically accept/decline meeting requests" -ForegroundColor White
                Write-Host "[2] AutoUpdate - Update calendar but don't respond" -ForegroundColor White
                Write-Host "[3] None - Manual processing required" -ForegroundColor White
                $choice = Read-Host "Select option (1-3)"
                $value = switch ($choice) {
                    "1" { "AutoAccept" }
                    "2" { "AutoUpdate" }
                    "3" { "None" }
                }
                if ($value) {
                    Write-Host "`nSetting AutomateProcessing to $value..." -ForegroundColor Cyan
                    Set-CalendarProcessing -Identity $roomInfo.Email -AutomateProcessing $value
                    Write-Host "Successfully updated!" -ForegroundColor Green
                }
            }
            "2" {
                $value = Read-Host "Enter new BookingWindowInDays (0-1080, current: $($roomInfo.BookingWindowInDays))"
                Write-Host "`nSetting BookingWindowInDays to $value..." -ForegroundColor Cyan
                Set-CalendarProcessing -Identity $roomInfo.Email -BookingWindowInDays $value
                Write-Host "Successfully updated!" -ForegroundColor Green
            }
            "3" {
                $value = Read-Host "Enter new MaximumDurationInMinutes (current: $($roomInfo.MaximumDurationInMinutes))"
                Write-Host "`nSetting MaximumDurationInMinutes to $value..." -ForegroundColor Cyan
                Set-CalendarProcessing -Identity $roomInfo.Email -MaximumDurationInMinutes $value
                Write-Host "Successfully updated!" -ForegroundColor Green
            }
            "4" {
                $value = Read-Host "Allow conflicts? [Y] Yes [N] No (current: $($roomInfo.AllowConflicts))"
                $boolValue = $value -match "[yY]"
                Write-Host "`nSetting AllowConflicts to $boolValue..." -ForegroundColor Cyan
                Set-CalendarProcessing -Identity $roomInfo.Email -AllowConflicts:$boolValue
                Write-Host "Successfully updated!" -ForegroundColor Green
            }
            "5" {
                $value = Read-Host "Set TentativePendingApproval (meetings wait for approval)? [Y] Yes [N] No (current: $($roomInfo.TentativePendingApproval))"
                $boolValue = $value -match "[yY]"
                Write-Host "`nSetting TentativePendingApproval to $boolValue..." -ForegroundColor Cyan
                Set-CalendarProcessing -Identity $roomInfo.Email -TentativePendingApproval:$boolValue
                Write-Host "Successfully updated!" -ForegroundColor Green
                if ($boolValue -eq $false) {
                    Write-Host "Note: Meetings will now be automatically accepted/rejected without waiting for approval." -ForegroundColor Green
                }
            }
            "6" {
                $value = Read-Host "Delete attachments? [Y] Yes [N] No (current: $($roomInfo.DeleteAttachments))"
                $boolValue = $value -match "[yY]"
                Write-Host "`nSetting DeleteAttachments to $boolValue..." -ForegroundColor Cyan
                Set-CalendarProcessing -Identity $roomInfo.Email -DeleteAttachments:$boolValue
                Write-Host "Successfully updated!" -ForegroundColor Green
            }
            "7" {
                $value = Read-Host "Delete comments? [Y] Yes [N] No (current: $($roomInfo.DeleteComments))"
                $boolValue = $value -match "[yY]"
                Write-Host "`nSetting DeleteComments to $boolValue..." -ForegroundColor Cyan
                Set-CalendarProcessing -Identity $roomInfo.Email -DeleteComments:$boolValue
                Write-Host "Successfully updated!" -ForegroundColor Green
            }
            "8" {
                $value = Read-Host "Delete subject? [Y] Yes [N] No (current: $($roomInfo.DeleteSubject))"
                $boolValue = $value -match "[yY]"
                Write-Host "`nSetting DeleteSubject to $boolValue..." -ForegroundColor Cyan
                Set-CalendarProcessing -Identity $roomInfo.Email -DeleteSubject:$boolValue
                Write-Host "Successfully updated!" -ForegroundColor Green
            }
            "9" {
                $value = Read-Host "Process external meeting messages? [Y] Yes [N] No (current: $($roomInfo.ProcessExternalMeetingMessages))"
                $boolValue = $value -match "[yY]"
                Write-Host "`nSetting ProcessExternalMeetingMessages to $boolValue..." -ForegroundColor Cyan
                Set-CalendarProcessing -Identity $roomInfo.Email -ProcessExternalMeetingMessages:$boolValue
                Write-Host "Successfully updated!" -ForegroundColor Green
            }
            "0" {
                Write-Host "Modification cancelled." -ForegroundColor Yellow
            }
            default {
                Write-Host "Invalid selection." -ForegroundColor Red
            }
        }
    }
}
finally {
    # Always disconnect from Exchange Online
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Done!" -ForegroundColor Green
}
