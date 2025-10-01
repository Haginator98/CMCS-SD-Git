# Install-Module Microsoft.Graph -Scope CurrentUser

# Logg inn i Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All"

# Angi gamle og nye adresser
$gammelAdresse = "6610"
$nyAdresse = "8410"

# Hent alle brukere (kan ta tid hvis mange brukere)
$alleBrukere = Get-MgUser -All -Property "displayName,streetAddress,userPrincipalName,id"

# Filtrer lokalt på streetAddress
$brukere = $alleBrukere | Where-Object { $_.StreetAddress -eq $gammelAdresse }

if ($brukere.Count -eq 0) {
    Write-Host "Ingen brukere med '$gammelAdresse'" -ForegroundColor Yellow
} else {
    Write-Host "Fant $($brukere.Count) brukere med '$gammelAdresse'" -ForegroundColor Cyan

    foreach ($bruker in $brukere) {
        Write-Host "Oppdaterer $($bruker.UserPrincipalName)..."
        Update-MgUser -UserId $bruker.Id -StreetAddress $nyAdresse
    }

    Write-Host "Ferdig oppdatert." -ForegroundColor Green
}