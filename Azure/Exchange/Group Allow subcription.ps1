# Script to allow or block group subscription in Exchange Online
# Checks if ExchangeOnlineManagement module is installed, if not prompts to install it
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
Start-Sleep -Seconds 2

# Get group name from user
$GroupName = Read-Host "Enter the group name (Display Name or Alias)"

Write-Host "Checking default settings for '$GroupName'..." -ForegroundColor Cyan
Get-UnifiedGroup -Identity $GroupName | fl DisplayName, SubscriptionEnabled, AlwaysSubscribeMembersToCalendarEvents, AutoSubscribeNewMembers

# Enable or disable subscription if group exists
if ($?) {
    $choice = Read-Host "Do you want to Enable or Disable subscription for this group? [E]nable / [D]isable / [N]o change"
    switch ($choice.ToLower()) {
        'e' {
            Set-UnifiedGroup -Identity $GroupName -SubscriptionEnabled $true -AlwaysSubscribeMembersToCalendarEvents $true -AutoSubscribeNewMembers $true
            Start-Sleep -Seconds 2
            Write-Host "Subscription ENABLED for group, showing new values '$GroupName'." -ForegroundColor Green
            Get-UnifiedGroup -Identity $GroupName | fl DisplayName, SubscriptionEnabled, AlwaysSubscribeMembersToCalendarEvents, AutoSubscribeNewMembers
        }
        'd' {
            Set-UnifiedGroup -Identity $GroupName -SubscriptionEnabled $false -AlwaysSubscribeMembersToCalendarEvents $false -AutoSubscribeNewMembers $false
            Start-Sleep -Seconds 2
            Write-Host "Subscription DISABLED for group, showing new values '$GroupName'." -ForegroundColor Green
            Get-UnifiedGroup -Identity $GroupName | fl DisplayName, SubscriptionEnabled, AlwaysSubscribeMembersToCalendarEvents, AutoSubscribeNewMembers
        }
        default {
            Write-Host "No changes made to group '$GroupName'." -ForegroundColor Yellow
        }
    }
    Write-Host "Exiting script." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected. Script finished." -ForegroundColor Green
    Start-Sleep -Seconds 3
    Clear-Host
} else {
    Write-Host "Group '$GroupName' not found. Please check the name and try again." -ForegroundColor Red
    Start-Sleep -Seconds 2
    & $PSCommandPath
    return
}


#Get-UnifiedGroup | Select DisplayName, Identity

#Set-UnifiedGroup -Identity Read-Host "Enter your group name here" -AlwaysSubscribeMembersToCalendarEvents $true -AutoSubscribeNewMembers $true