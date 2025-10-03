# Install-Module Microsoft.Graph -Scope CurrentUser

# Logg inn i Microsoft Graph
#Connect-MgGraph -Scopes "User.ReadWrite.All"

# Angi gamle og nye adresser
$gammelAdresse = Read-Host "Skriv inn gammel adresse"
$nyAdresse = Read-Host "Skriv inn ny adresse"

# Hent alle brukere (kan ta tid hvis mange brukere)
$alleBrukere = Get-MgUser -All -Property "displayName,streetAddress,userPrincipalName,id"
    Write-Host "Henter alle brukere..." -ForegroundColor Cyan

# Filtrer lokalt på streetAddress
$brukere = $alleBrukere | Where-Object { $_.StreetAddress -eq $gammelAdresse }

if ($brukere.Count -eq 0) {
    Write-Host "Ingen brukere med '$gammelAdresse'" -ForegroundColor Yellow
} else {
    Write-Host "Fant $($brukere.Count) brukere med '$gammelAdresse'" -ForegroundColor Cyan

    $ready = $false
    while (-not $ready) {
        $confirm = Read-Host "Er du sikker på at du vil angi '$nyAdresse' som Street Address for $($brukere.Count) brukere? (y = yes, n = no, l = list users)"
        switch ($confirm.ToLower()) {
            'y' {
                $ready = $true
            }
            'n' {
                Write-Host "Operation cancelled." -ForegroundColor Red
                return
            }
            'l' {
                Write-Host "Brukere som vil bli oppdatert:" -ForegroundColor Cyan
                $brukere | Select-Object UserPrincipalName, DisplayName, streetAddress | Format-Table
            }
            default {
                Write-Host "y (yes), n (no), or l (list users)." -ForegroundColor Yellow
            }
        }
    }

    foreach ($bruker in $brukere) {
        Write-Host "Oppdaterer $($bruker.UserPrincipalName)..."
        Update-MgUser -UserId $bruker.Id -StreetAddress $nyAdresse
    }

    Write-Host "Ferdig oppdatert." -ForegroundColor Green
}