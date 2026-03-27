# Import required modules (installed via Tools.ps1)
Import-Module ExchangeOnlineManagement -ErrorAction Stop

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline | Out-Null
Write-Host "This script will create a new Dynamic Distribution List based on recipient filter criteria." -ForegroundColor Cyan

# Gather DL information
$dlName = Read-Host "`nEnter the Display Name for the Dynamic DL"
if ([string]::IsNullOrWhiteSpace($dlName)) {
    Write-Host "Display Name cannot be empty." -ForegroundColor Red
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

$dlAlias = Read-Host "Enter the Alias (mail nickname, no spaces)"
if ([string]::IsNullOrWhiteSpace($dlAlias)) {
    Write-Host "Alias cannot be empty." -ForegroundColor Red
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

$dlPrimarySMTP = Read-Host "Enter the Primary SMTP address (e.g., group@domain.com)"
if ([string]::IsNullOrWhiteSpace($dlPrimarySMTP)) {
    Write-Host "Primary SMTP address cannot be empty." -ForegroundColor Red
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Recipient filter selection
Write-Host "`nChoose a filter type for the Dynamic DL:" -ForegroundColor Yellow
Write-Host "[1] Department"
Write-Host "[2] Company"
Write-Host "[3] State or Province"
Write-Host "[4] Custom Attribute (1-15)"
Write-Host "[5] Office"
Write-Host "[6] Custom RecipientFilter (advanced)"

$filterChoice = Read-Host "Enter your choice (1-6)"

switch ($filterChoice) {
    "1" {
        $filterValue = Read-Host "Enter the Department name"
        $recipientFilter = "Department -eq '$filterValue'"
    }
    "2" {
        $filterValue = Read-Host "Enter the Company name"
        $recipientFilter = "Company -eq '$filterValue'"
    }
    "3" {
        $filterValue = Read-Host "Enter the State or Province"
        $recipientFilter = "StateOrProvince -eq '$filterValue'"
    }
    "4" {
        $attrNumber = Read-Host "Enter the Custom Attribute number (1-15)"
        if ($attrNumber -notmatch "^([1-9]|1[0-5])$") {
            Write-Host "Invalid Custom Attribute number. Must be 1-15." -ForegroundColor Red
            Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
            Disconnect-ExchangeOnline -Confirm:$false
            Exit
        }
        $filterValue = Read-Host "Enter the Custom Attribute value"
        $recipientFilter = "CustomAttribute$attrNumber -eq '$filterValue'"
    }
    "5" {
        $filterValue = Read-Host "Enter the Office name"
        $recipientFilter = "Office -eq '$filterValue'"
    }
    "6" {
        $recipientFilter = Read-Host "Enter the full RecipientFilter string (e.g., Department -eq 'Sales' -and Company -eq 'Contoso')"
    }
    default {
        Write-Host "Invalid choice." -ForegroundColor Red
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Exit
    }
}

if ([string]::IsNullOrWhiteSpace($recipientFilter)) {
    Write-Host "RecipientFilter cannot be empty." -ForegroundColor Red
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Preview matching recipients
Write-Host "`nChecking how many recipients match the filter..." -ForegroundColor Cyan
try {
    $matchingRecipients = Get-Recipient -RecipientPreviewFilter $recipientFilter -ResultSize Unlimited -ErrorAction Stop
    $matchCount = ($matchingRecipients | Measure-Object).Count
} catch {
    Write-Host "Error: Invalid filter or unable to query recipients." -ForegroundColor Red
    Write-Host "Filter used: $recipientFilter" -ForegroundColor DarkGray
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor DarkGray
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Show summary and ask for confirmation
Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "DYNAMIC DISTRIBUTION LIST SUMMARY" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "Display Name     : $dlName" -ForegroundColor White
Write-Host "Alias            : $dlAlias" -ForegroundColor White
Write-Host "Primary SMTP     : $dlPrimarySMTP" -ForegroundColor White
Write-Host "RecipientFilter  : $recipientFilter" -ForegroundColor White
Write-Host "Matching users   : $matchCount" -ForegroundColor White
Write-Host "========================================`n" -ForegroundColor Yellow

if ($matchCount -eq 0) {
    Write-Host "WARNING: No recipients currently match this filter. The DL will be empty." -ForegroundColor Yellow
}

$confirmation = Read-Host "Do you want to create this Dynamic Distribution List? [Y] Yes [N] No"
if ($confirmation -notmatch "[yY]") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Create the Dynamic Distribution List
Write-Host "`nCreating Dynamic Distribution List..." -ForegroundColor Cyan
try {
    New-DynamicDistributionGroup `
        -Name $dlName `
        -Alias $dlAlias `
        -PrimarySmtpAddress $dlPrimarySMTP `
        -RecipientFilter $recipientFilter `
        -ErrorAction Stop | Out-Null

    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "Dynamic Distribution List created successfully!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Display Name     : $dlName" -ForegroundColor White
    Write-Host "Primary SMTP     : $dlPrimarySMTP" -ForegroundColor White
    Write-Host "Matching users   : $matchCount" -ForegroundColor White
    Write-Host "========================================`n" -ForegroundColor Green
} catch {
    Write-Host "`nError: Failed to create Dynamic Distribution List." -ForegroundColor Red
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor DarkGray
    Write-Host "This may mean the alias or SMTP address is already in use, or you lack the required permissions." -ForegroundColor Yellow
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Disconnect from Exchange Online
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected successfully." -ForegroundColor Green
Start-Sleep -Seconds 2
