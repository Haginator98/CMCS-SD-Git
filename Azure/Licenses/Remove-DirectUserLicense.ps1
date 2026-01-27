# Import required modules (installed via Tools.ps1)
Write-Host "Importing required modules..." -ForegroundColor Cyan
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Connect to Microsoft Graph
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All", "Organization.Read.All" -NoWelcome

# Get all available licenses in tenant for reference
Write-Host "Retrieving available licenses..." -ForegroundColor Cyan
$subscribedSkus = Get-MgSubscribedSku
$skuHashTable = @{}
foreach ($sku in $subscribedSkus) {
    $skuHashTable[$sku.SkuId] = $sku.SkuPartNumber
}

do {
    Write-Host "`n================================================" -ForegroundColor Cyan
    Write-Host "           REMOVE USER LICENSES" -ForegroundColor Cyan
    Write-Host "================================================`n" -ForegroundColor Cyan
    
    # Ask for user
    $userInput = Read-Host "Enter user email or UPN"
    
    if ([string]::IsNullOrWhiteSpace($userInput)) {
        Write-Host "No user specified. Exiting..." -ForegroundColor Yellow
        break
    }
    
    # Get user
    try {
        $user = Get-MgUser -UserId $userInput -Property Id, DisplayName, UserPrincipalName, AssignedLicenses -ErrorAction Stop
        
        if ([string]::IsNullOrEmpty($user.Id)) {
            Write-Host "Error: User object retrieved but ID is empty" -ForegroundColor Red
            continue
        }
        
        Write-Host "`nUser found: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Green
    }
    catch {
        Write-Host "Error: Could not find user with email/UPN: $userInput" -ForegroundColor Red
        Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
        continue
    }
    
    # Check if user has licenses
    if ($user.AssignedLicenses.Count -eq 0) {
        Write-Host "`nUser has no licenses assigned." -ForegroundColor Yellow
        continue
    }
    
    # Display all licenses
    Write-Host "`nUser's licenses:" -ForegroundColor Cyan
    Write-Host "===================" -ForegroundColor Cyan
    
    $licenseList = @()
    $index = 1
    
    foreach ($license in $user.AssignedLicenses) {
        $skuName = $skuHashTable[$license.SkuId]
        if ([string]::IsNullOrEmpty($skuName)) {
            $skuName = $license.SkuId
        }
        
        $licenseList += [PSCustomObject]@{
            Index = $index
            SkuId = $license.SkuId
            LicenseName = $skuName
        }
        
        Write-Host "$index. $skuName" -ForegroundColor White
        $index++
    }
    
    # Ask which license to remove
    Write-Host "`nSelect license(s) to remove:" -ForegroundColor Yellow
    Write-Host "You can enter one number (e.g. 1), multiple numbers separated by comma (e.g. 1,3,4)" -ForegroundColor Gray
    Write-Host "or 'all' to remove all licenses." -ForegroundColor Gray
    $selection = Read-Host "Your choice"
    
    if ([string]::IsNullOrWhiteSpace($selection)) {
        Write-Host "No license selected. Skipping..." -ForegroundColor Yellow
        continue
    }
    
    # Parse selection
    $licensesToRemove = @()
    
    if ($selection.ToLower() -eq 'all') {
        $licensesToRemove = $licenseList
        Write-Host "`nYou have selected to remove ALL licenses." -ForegroundColor Yellow
    }
    else {
        $selectedIndexes = $selection -split ',' | ForEach-Object { $_.Trim() }
        
        foreach ($idx in $selectedIndexes) {
            if ($idx -match '^\d+$') {
                $idxNum = [int]$idx
                $selectedLicense = $licenseList | Where-Object { $_.Index -eq $idxNum }
                
                if ($selectedLicense) {
                    $licensesToRemove += $selectedLicense
                }
                else {
                    Write-Host "Invalid selection: $idx" -ForegroundColor Red
                }
            }
            else {
                Write-Host "Invalid format: $idx" -ForegroundColor Red
            }
        }
    }
    
    if ($licensesToRemove.Count -eq 0) {
        Write-Host "No valid licenses selected. Skipping..." -ForegroundColor Yellow
        continue
    }
    
    # Confirmation
    Write-Host "`nYou are about to remove the following license(s):" -ForegroundColor Yellow
    foreach ($lic in $licensesToRemove) {
        Write-Host "  - $($lic.LicenseName)" -ForegroundColor White
    }
    Write-Host "From user: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Yellow
    
    $confirm = Read-Host "`nAre you sure? (Y/N)"
    
    if ($confirm -notmatch '^[Yy]') {
        Write-Host "Cancelled by user." -ForegroundColor Yellow
        continue
    }
    
    # Remove licenses
    try {
        $skuIdsToRemove = $licensesToRemove | ForEach-Object { $_.SkuId }
        
        Set-MgUserLicense -UserId $user.Id -RemoveLicenses $skuIdsToRemove -AddLicenses @() -ErrorAction Stop
        
        Write-Host "`n✓ License(s) removed!" -ForegroundColor Green
        
        foreach ($lic in $licensesToRemove) {
            Write-Host "  ✓ $($lic.LicenseName)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "`n✗ Error removing license(s)!" -ForegroundColor Red
        Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
        
        if ($_.Exception.Message -match "depends on the service plan") {
            Write-Host "`nThis error means other licenses have dependencies on this license." -ForegroundColor Yellow
            Write-Host "You need to remove the dependent licenses first, or remove all licenses together." -ForegroundColor Yellow
            Write-Host "Try selecting 'all' to remove all licenses at once." -ForegroundColor Yellow
        }
    }
    
    # Ask if user wants to continue with more users
    Write-Host "`n"
    $continue = Read-Host "Remove licenses from another user? (Y/N)"
    
} while ($continue -match '^[Yy]')

Write-Host "`nExiting..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Done!" -ForegroundColor Green
