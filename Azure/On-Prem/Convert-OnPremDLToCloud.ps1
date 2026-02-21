Import-Module ExchangeOnlineManagement -ErrorAction Stop

Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline | Out-Null

Write-Host "`nThis script helps convert a synced/on-prem Distribution List to cloud-only." -ForegroundColor Cyan
Write-Host "Flow: Export details -> Manual on-prem delete/sync -> Create cloud DL -> Restore owners/members -> Add X500" -ForegroundColor Yellow

$sourceDL = $null
$oldDL = $null
while (-not $oldDL) {
    $sourceDL = Read-Host "`nEnter current synced/on-prem DL email address"

    try {
        $oldDL = Get-DistributionGroup -Identity $sourceDL -ErrorAction Stop
        Write-Host "Found Distribution List $($oldDL.DisplayName)" -ForegroundColor Green
    } catch {
        Write-Host "Could not find Distribution List '$sourceDL'." -ForegroundColor Red
        $retry = Read-Host "Try again? [Y] Yes [N] No"

        if ($retry -notmatch "[yY]") {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
            Disconnect-ExchangeOnline -Confirm:$false
            Exit
        }
    }
}

$members = Get-DistributionGroupMember -Identity $oldDL.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue |
    Select-Object @{Name='Email';Expression={
        if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress.ToString() } else { 'no email address' }
    }}

$owners = @()
if ($oldDL.ManagedBy -and $oldDL.ManagedBy.Count -gt 0) {
    $owners = $oldDL.ManagedBy | ForEach-Object {
        $ownerRecipient = Get-Recipient -Identity $_ -ErrorAction SilentlyContinue
        [PSCustomObject]@{
            Email = if ($ownerRecipient -and $ownerRecipient.PrimarySmtpAddress) { $ownerRecipient.PrimarySmtpAddress.ToString() } else { 'no email address' }
        }
    }
}

$emailAddresses = @($oldDL.EmailAddresses | ForEach-Object { $_.ToString() })
$legacyExchangeDN = if ([string]::IsNullOrWhiteSpace($oldDL.LegacyExchangeDN)) { 'no legacyExchangeDN' } else { $oldDL.LegacyExchangeDN }

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "EXPORT PREVIEW" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "DL: $($oldDL.DisplayName) <$($oldDL.PrimarySmtpAddress)>" -ForegroundColor White
Write-Host "Members found: $($members.Count)" -ForegroundColor White
Write-Host "Owners found: $($owners.Count)" -ForegroundColor White
Write-Host "Email addresses found: $($emailAddresses.Count)" -ForegroundColor White

$confirmExport = Read-Host "`nDo you want to export backup files to Desktop? [Y] Yes [N] No"
if ($confirmExport -notmatch "[yY]") {
    Write-Host "Export cancelled by user." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

$desktop = [Environment]::GetFolderPath("Desktop")
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$safeAlias = if ([string]::IsNullOrWhiteSpace($oldDL.Alias)) { "DL" } else { $oldDL.Alias }

$combinedFile = Join-Path $desktop "DL_CloudConversion_AllData_${safeAlias}_$timestamp.csv"

$metadata = [PSCustomObject]@{
    DisplayName = $oldDL.DisplayName
    Alias = if ([string]::IsNullOrWhiteSpace($oldDL.Alias)) { 'no alias' } else { $oldDL.Alias }
    PrimarySmtpAddress = if ($oldDL.PrimarySmtpAddress) { $oldDL.PrimarySmtpAddress.ToString() } else { 'no primary smtp' }
    LegacyExchangeDN = $legacyExchangeDN
    HiddenFromAddressListsEnabled = $oldDL.HiddenFromAddressListsEnabled
    RequireSenderAuthenticationEnabled = $oldDL.RequireSenderAuthenticationEnabled
    ModerationEnabled = $oldDL.ModerationEnabled
}

$conversionData = @()

$conversionData += [PSCustomObject]@{ RecordType='Metadata'; Key='DisplayName'; Value=$metadata.DisplayName }
$conversionData += [PSCustomObject]@{ RecordType='Metadata'; Key='Alias'; Value=$metadata.Alias }
$conversionData += [PSCustomObject]@{ RecordType='Metadata'; Key='PrimarySmtpAddress'; Value=$metadata.PrimarySmtpAddress }
$conversionData += [PSCustomObject]@{ RecordType='Metadata'; Key='LegacyExchangeDN'; Value=$metadata.LegacyExchangeDN }
$conversionData += [PSCustomObject]@{ RecordType='Metadata'; Key='HiddenFromAddressListsEnabled'; Value=$metadata.HiddenFromAddressListsEnabled }
$conversionData += [PSCustomObject]@{ RecordType='Metadata'; Key='RequireSenderAuthenticationEnabled'; Value=$metadata.RequireSenderAuthenticationEnabled }
$conversionData += [PSCustomObject]@{ RecordType='Metadata'; Key='ModerationEnabled'; Value=$metadata.ModerationEnabled }

foreach ($member in $members) {
    $conversionData += [PSCustomObject]@{ RecordType='Member'; Key='Email'; Value=$member.Email }
}

foreach ($owner in $owners) {
    $conversionData += [PSCustomObject]@{ RecordType='Owner'; Key='Email'; Value=$owner.Email }
}

foreach ($address in $emailAddresses) {
    $conversionData += [PSCustomObject]@{ RecordType='Address'; Key='ProxyAddress'; Value=$address }
}

$conversionData | Export-Csv -Path $combinedFile -NoTypeInformation -Encoding UTF8

Write-Host "`nExport complete:" -ForegroundColor Green
Write-Host "- $combinedFile" -ForegroundColor Green

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "MANUAL STEP REQUIRED" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "1) Delete or de-scope the on-prem DL from sync" -ForegroundColor Yellow
Write-Host "2) Run/trigger Entra Connect sync" -ForegroundColor Yellow
Write-Host "3) Wait until DL is removed from Exchange Online" -ForegroundColor Yellow
Write-Host "" 

$continueAfterManual = Read-Host "When done, do you want to continue to cloud creation phase? [Y] Yes [N] No"
if ($continueAfterManual -notmatch "[yY]") {
    Write-Host "Stopped by user after export." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

while ($true) {
    Write-Host "`nCheck status options: [C] Check if old DL still exists, [P] Proceed anyway, [N] No/Cancel" -ForegroundColor Cyan
    $statusChoice = Read-Host "Enter choice"

    if ($statusChoice -match "^[cC]$") {
        $stillExists = Get-DistributionGroup -Identity $sourceDL -ErrorAction SilentlyContinue
        if ($stillExists) {
            Write-Host "Old synced DL still exists in cloud. Wait for sync completion before creating new cloud DL." -ForegroundColor Red
        } else {
            Write-Host "Old DL is no longer found. Safe to continue." -ForegroundColor Green
            break
        }
    } elseif ($statusChoice -match "^[pP]$") {
        $forceProceed = Read-Host "Proceed even if old object may still exist? [Y] Yes [N] No"
        if ($forceProceed -match "[yY]") { break }
    } elseif ($statusChoice -match "^[nN]$") {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Exit
    } else {
        Write-Host "Invalid choice." -ForegroundColor Red
    }
}

$newDisplayName = Read-Host "`nEnter new cloud DL display name (Press Enter for '$($oldDL.DisplayName)')"
if ([string]::IsNullOrWhiteSpace($newDisplayName)) { $newDisplayName = $oldDL.DisplayName }

$newPrimarySmtp = Read-Host "Enter new cloud DL primary SMTP (Press Enter for '$($oldDL.PrimarySmtpAddress)')"
if ([string]::IsNullOrWhiteSpace($newPrimarySmtp)) { $newPrimarySmtp = $oldDL.PrimarySmtpAddress.ToString() }

$defaultAlias = if ([string]::IsNullOrWhiteSpace($oldDL.Alias)) { ($newPrimarySmtp -split '@')[0] } else { $oldDL.Alias }
$newAlias = Read-Host "Enter alias (Press Enter for '$defaultAlias')"
if ([string]::IsNullOrWhiteSpace($newAlias)) { $newAlias = $defaultAlias }

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "CREATE CLOUD DL PREVIEW" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Display name: $newDisplayName" -ForegroundColor White
Write-Host "Primary SMTP: $newPrimarySmtp" -ForegroundColor White
Write-Host "Alias: $newAlias" -ForegroundColor White
Write-Host "Members to import: $($members.Count)" -ForegroundColor White
Write-Host "Owners to import: $($owners.Count)" -ForegroundColor White

$confirmCreate = Read-Host "`nDo you want to create and configure cloud DL now? [Y] Yes [N] No"
if ($confirmCreate -notmatch "[yY]") {
    Write-Host "Cloud creation cancelled by user." -ForegroundColor Yellow
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

try {
    New-DistributionGroup -Name $newDisplayName -Alias $newAlias -PrimarySmtpAddress $newPrimarySmtp -Type Distribution -ErrorAction Stop | Out-Null
    Write-Host "Cloud DL created successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to create cloud DL: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

try {
    Set-DistributionGroup -Identity $newPrimarySmtp `
        -HiddenFromAddressListsEnabled $oldDL.HiddenFromAddressListsEnabled `
        -RequireSenderAuthenticationEnabled $oldDL.RequireSenderAuthenticationEnabled `
        -ModerationEnabled $oldDL.ModerationEnabled `
        -ErrorAction SilentlyContinue
} catch {
    Write-Host "Could not copy one or more DL settings: $($_.Exception.Message)" -ForegroundColor Yellow
}

$addressesToAdd = @()
foreach ($address in $emailAddresses) {
    if ($address -match "^SMTP:") { continue }
    if ($address -notin $addressesToAdd) { $addressesToAdd += $address }
}

if ($oldDL.LegacyExchangeDN) {
    $x500Address = "X500:$($oldDL.LegacyExchangeDN)"
    if ($x500Address -notin $addressesToAdd) {
        $addressesToAdd += $x500Address
    }
}

if ($addressesToAdd.Count -gt 0) {
    try {
        Set-DistributionGroup -Identity $newPrimarySmtp -EmailAddresses @{Add = $addressesToAdd} -ErrorAction Stop
        Write-Host "Added $($addressesToAdd.Count) proxy/X500 address(es)." -ForegroundColor Green
    } catch {
        Write-Host "Could not add all proxy/X500 addresses: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$ownerSuccess = 0
$ownerErrors = @()
if ($owners.Count -gt 0) {
    foreach ($owner in $owners) {
        if ([string]::IsNullOrWhiteSpace($owner.Email) -or $owner.Email -eq 'no email address') { continue }
        try {
            Set-DistributionGroup -Identity $newPrimarySmtp -ManagedBy @{Add = $owner.Email} -BypassSecurityGroupManagerCheck -ErrorAction Stop
            $ownerSuccess++
        } catch {
            $ownerErrors += "$($owner.Email) - $($_.Exception.Message)"
        }
    }
}

$memberSuccess = 0
$memberAlready = 0
$memberErrors = @()

if ($members.Count -gt 0) {
    for ($index = 0; $index -lt $members.Count; $index++) {
        $currentEmail = $members[$index].Email

        Write-Progress -Activity "Importing members to cloud DL" `
                       -Status "Processing $($index + 1) of $($members.Count): $currentEmail" `
                       -PercentComplete ((($index + 1) / $members.Count) * 100)

        if ([string]::IsNullOrWhiteSpace($currentEmail) -or $currentEmail -eq 'no email address') {
            continue
        }

        try {
            Add-DistributionGroupMember -Identity $newPrimarySmtp -Member $currentEmail -ErrorAction Stop
            $memberSuccess++
        } catch {
            if ($_.Exception.Message -like "*already a member*") {
                $memberAlready++
            } else {
                $memberErrors += "$currentEmail - $($_.Exception.Message)"
            }
        }
    }
}

Write-Progress -Activity "Importing members to cloud DL" -Completed

$resultRows = @(
    [PSCustomObject]@{ Item = 'SourceDL'; Value = $sourceDL },
    [PSCustomObject]@{ Item = 'NewCloudDL'; Value = $newPrimarySmtp },
    [PSCustomObject]@{ Item = 'OwnersAdded'; Value = $ownerSuccess },
    [PSCustomObject]@{ Item = 'OwnerErrors'; Value = $ownerErrors.Count },
    [PSCustomObject]@{ Item = 'MembersAdded'; Value = $memberSuccess },
    [PSCustomObject]@{ Item = 'MembersAlreadyPresent'; Value = $memberAlready },
    [PSCustomObject]@{ Item = 'MemberErrors'; Value = $memberErrors.Count },
    [PSCustomObject]@{ Item = 'ProxyAndX500Added'; Value = $addressesToAdd.Count }
)

foreach ($row in $resultRows) {
    $conversionData += [PSCustomObject]@{ RecordType='Result'; Key=$row.Item; Value=$row.Value }
}

foreach ($ownerError in $ownerErrors) {
    $conversionData += [PSCustomObject]@{ RecordType='OwnerError'; Key='Error'; Value=$ownerError }
}

foreach ($memberError in $memberErrors) {
    $conversionData += [PSCustomObject]@{ RecordType='MemberError'; Key='Error'; Value=$memberError }
}

$conversionData | Export-Csv -Path $combinedFile -NoTypeInformation -Encoding UTF8

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "CONVERSION SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "New cloud DL: $newPrimarySmtp" -ForegroundColor White
Write-Host "Owners added: $ownerSuccess" -ForegroundColor Green
Write-Host "Owner errors: $($ownerErrors.Count)" -ForegroundColor Yellow
Write-Host "Members added: $memberSuccess" -ForegroundColor Green
Write-Host "Members already present: $memberAlready" -ForegroundColor Yellow
Write-Host "Member errors: $($memberErrors.Count)" -ForegroundColor Yellow
Write-Host "Proxy/X500 addresses added: $($addressesToAdd.Count)" -ForegroundColor Green
Write-Host "Combined file: $combinedFile" -ForegroundColor Green

if ($ownerErrors.Count -gt 0) {
    Write-Host "`nOwner errors:" -ForegroundColor Red
    $ownerErrors | ForEach-Object { Write-Host "- $_" -ForegroundColor Red }
}

if ($memberErrors.Count -gt 0) {
    Write-Host "`nMember errors:" -ForegroundColor Red
    $memberErrors | ForEach-Object { Write-Host "- $_" -ForegroundColor Red }
}
Start-Sleep -Seconds 3
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Done." -ForegroundColor Green
