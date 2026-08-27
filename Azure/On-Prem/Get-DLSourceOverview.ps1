# Import required modules (installed via Tools.ps1)
Import-Module ExchangeOnlineManagement -ErrorAction Stop

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline | Out-Null

Write-Host "`nThis script counts all distribution lists and splits them by source (on-prem synced vs cloud-only)." -ForegroundColor Cyan
Write-Host "Fetching distribution lists from Exchange Online..." -ForegroundColor Cyan

try {
    # IsDirSynced = $true means the object is synced from on-prem AD (AD Connect / Entra Connect)
    $allDLs = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop |
        Select-Object DisplayName, PrimarySmtpAddress, GroupType, RecipientTypeDetails, IsDirSynced, WhenCreated
} catch {
    Write-Host "Error: Could not retrieve distribution lists." -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

if (-not $allDLs -or $allDLs.Count -eq 0) {
    Write-Host "No distribution lists found in the tenant." -ForegroundColor Yellow
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

$onPremDLs = $allDLs | Where-Object { $_.IsDirSynced -eq $true }
$cloudDLs  = $allDLs | Where-Object { $_.IsDirSynced -ne $true }

$total    = $allDLs.Count
$onPrem   = @($onPremDLs).Count
$cloud    = @($cloudDLs).Count

if ($total -gt 0) {
    $onPremPct = [math]::Round(($onPrem / $total) * 100, 1)
    $cloudPct  = [math]::Round(($cloud  / $total) * 100, 1)
} else {
    $onPremPct = 0
    $cloudPct  = 0
}

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "DISTRIBUTION LIST OVERVIEW" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host ("Total distribution lists : {0}" -f $total)   -ForegroundColor Green
Write-Host ("On-prem synced (AD)      : {0} ({1}%)" -f $onPrem, $onPremPct) -ForegroundColor Cyan
Write-Host ("Cloud-only (Entra/EXO)   : {0} ({1}%)" -f $cloud,  $cloudPct)  -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Yellow

# Ask before exporting details to CSV
$exportChoice = Read-Host "Do you want to export the full list to CSV on your desktop? [Y] Yes [N] No"
if ($exportChoice -match "[yY]") {
    $Desktop   = [Environment]::GetFolderPath("Desktop")
    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $ExportFile = Join-Path $Desktop "DistributionLists_Overview_$Timestamp.csv"

    $allDLs |
        Select-Object DisplayName,
                      PrimarySmtpAddress,
                      RecipientTypeDetails,
                      GroupType,
                      @{Name = "Source"; Expression = { if ($_.IsDirSynced) { "On-Prem" } else { "Cloud" } }},
                      WhenCreated |
        Sort-Object Source, DisplayName |
        Export-Csv -Path $ExportFile -NoTypeInformation -Encoding UTF8

    Write-Host "Exported to: $ExportFile" -ForegroundColor Green
} else {
    Write-Host "Export skipped." -ForegroundColor Yellow
}

Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected successfully." -ForegroundColor Green
Start-Sleep -Seconds 2
