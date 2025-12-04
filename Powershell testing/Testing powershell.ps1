#Powershell: 

Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName epostadresse@dittdomene.no

Get-Mailbox -RecipientTypeDetails RoomMailbox | 
Where-Object {$_.Building -eq "Dronning Maudsgt 10"} | 
Select Name, Office, Building, Floor

Get-Mailbox -RecipientTypeDetails RoomMailbox | 
Where-Object {$_.Building -eq "Dronning Maudsgt 10"} | 
Select Name, Office, Building, Floor

Get-Mailbox -RecipientTypeDetails RoomMailbox | 
Select-Object Name, DisplayName, ResourceCapacity, Office, City, Building, Floor






# Sammenlign to møterom i Exchange Online
# Du legger inn UPN-ene manuelt

$room1 = Read-Host "Skriv inn UPN for rom 1"
$room2 = Read-Host "Skriv inn UPN for rom 2"

$r1 = Get-Mailbox -Identity $room1
$r2 = Get-Mailbox -Identity $room2

$fields = 'DisplayName','PrimarySmtpAddress','Office','Building','Floor','ResourceCapacity','City','Department','CountryOrRegion'

Write-Host "`n🔍 Sammenligner $($r1.DisplayName) og $($r2.DisplayName)...`n" -ForegroundColor Cyan

foreach ($f in $fields) {
    if ($r1.$f -ne $r2.$f) {
        Write-Host ("{0,-20}: {1}  ≠  {2}" -f $f, $r1.$f, $r2.$f) -ForegroundColor Yellow
    }
}

Write-Host "`n✅ Ferdig!" -ForegroundColor Green