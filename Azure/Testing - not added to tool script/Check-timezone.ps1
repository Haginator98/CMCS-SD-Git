#---------------------------------------------------------------------------#
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

Connect-ExchangeOnline 
#---------------------------------------------------------------------------#
Write-Host "This script will check the timezone for the room mailbox" -ForegroundColor Cyan
$room = Read-Host "Enter the room mailbox email address"

# Hent tidssoneinformasjon
$timezoneInfo = Get-MailboxRegionalConfiguration -Identity $room | Select-Object TimeZone, Language, DateFormat, TimeFormat

# Skriv ut resultatet
if ($timezoneInfo) {
    Write-Host "Timezone: $($timezoneInfo.TimeZone)" -ForegroundColor Yellow
    Write-Host "Language: $($timezoneInfo.Language)" -ForegroundColor Yellow
    Write-Host "Date Format: $($timezoneInfo.DateFormat)" -ForegroundColor Yellow
    Write-Host "Time Format: $($timezoneInfo.TimeFormat)" -ForegroundColor Yellow
} else {
    Write-Host "No timezone information found for $room" -ForegroundColor Red
    Write-Host "Exiting script." -ForegroundColor Red
    Start-Sleep -Seconds 2
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

Write-Host "Timezone information retrieved. Do you want to change it?" -ForegroundColor Green
$change = Read-Host "Change timezone? [Y] Yes [N] No"
if ($change -match "[yY]") {
    $newTZ = Read-Host "Enter the new timezone (e.g., 'W. Europe Standard Time')"
    Set-MailboxRegionalConfiguration -Identity $room -TimeZone $newTZ
    Write-Host "Timezone updated to $newTZ for mailbox $room" -ForegroundColor Green
    Start-Sleep -Seconds 2
} else {
    Write-Host "No changes made to the timezone." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
}

Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected. Script finished." -ForegroundColor Green
Start-Sleep -Seconds 2
