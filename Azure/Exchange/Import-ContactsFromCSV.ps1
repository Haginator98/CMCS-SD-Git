# Bulk Import External Contacts to Exchange Online
# This script imports external contacts from a CSV file with only email addresses
# The email address will be used as the display name

# Import required modules (installed via Tools.ps1)
Import-Module ExchangeOnlineManagement -ErrorAction Stop

Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline | Out-Null

Write-Host "`nThis script will bulk import external contacts from a CSV file." -ForegroundColor Cyan
Write-Host "Please use the CSV from the CSV Template folder as this has the correct format." -ForegroundColor Cyan

# Prompt for CSV file
$csvPath = Read-Host "`nEnter the full path to your CSV file (or drag and drop the file here)"
$csvPath = $csvPath.Trim('"').Trim("'")

if (-not (Test-Path $csvPath)) {
    Write-Host "Error: File not found at path: $csvPath" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Import CSV
try {
    $contacts = Import-Csv $csvPath
    Write-Host "`nFound $($contacts.Count) contact(s) in CSV file." -ForegroundColor Green
} catch {
    Write-Host "Error reading CSV file: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Ask if contacts should be hidden from address book
$hideFromAddressBook = Read-Host "`nDo you want to hide these contacts from the Global Address List? [Y] Yes [N] No (Recommended: Yes if only for DL use)"

# Track results
$created = @()
$skipped = @()
$errors = @()

Write-Host "`nStarting import process..." -ForegroundColor Cyan
# Start timer
$importStartTime = Get-Date
$currentCount = 0
foreach ($contact in $contacts) {
    $currentCount++
    $email = $contact.Email.Trim()
    
    if ([string]::IsNullOrWhiteSpace($email)) {
        Write-Host "[$currentCount/$($contacts.Count)] Skipping empty email address" -ForegroundColor Yellow
        continue
    }
    
    Write-Progress -Activity "Importing External Contacts" `
                   -Status "Processing $currentCount of $($contacts.Count): $email" `
                   -PercentComplete (($currentCount / $contacts.Count) * 100)
    
    # Use email address as name (remove domain for display)
    $displayName = $email.Split('@')[0]
    
    try {
        # Check if contact already exists
        $existingContact = Get-MailContact -Identity $email -ErrorAction SilentlyContinue
        
        if ($existingContact) {
            Write-Host "[$currentCount/$($contacts.Count)] Contact already exists: $email" -ForegroundColor Yellow
            $skipped += $email
        } else {
            # Create new mail contact
            New-MailContact -Name $email `
                           -DisplayName $displayName `
                           -ExternalEmailAddress $email `
                           -FirstName $displayName `
                           -LastName "" | Out-Null
            
            Write-Host "[$currentCount/$($contacts.Count)] Created contact: $email" -ForegroundColor Green
            $created += $email
            
            # Hide from address book if requested
            if ($hideFromAddressBook -match "[yY]") {
                Set-MailContact -Identity $email -HiddenFromAddressListsEnabled $true
            }
        }
    } catch {
        Write-Host "[$currentCount/$($contacts.Count)] Error creating contact $email : $($_.Exception.Message)" -ForegroundColor Red
        $errors += "$email - $($_.Exception.Message)"
    }
    
    Start-Sleep -Milliseconds 500  # Small delay to avoid throttling
}

Write-Progress -Activity "Importing External Contacts" -Completed

# Calculate elapsed time
$importEndTime = Get-Date
$elapsedTime = $importEndTime - $importStartTime
$timeFormatted = "{0:D2}:{1:D2}:{2:D2}" -f $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Import Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "`nTime elapsed: $timeFormatted (HH:MM:SS)" -ForegroundColor Cyan
Write-Host "Successfully created: $($created.Count)" -ForegroundColor Green
Write-Host "Already existed (skipped): $($skipped.Count)" -ForegroundColor Yellow
Write-Host "Errors: $($errors.Count)" -ForegroundColor Red

if ($errors.Count -gt 0) {
    Write-Host "`nErrors encountered:" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
}

# Export results
$Desktop = [Environment]::GetFolderPath("Desktop")
$Date = Get-Date -Format "yyyy-MM-dd_HHmmss"
$ResultsFile = Join-Path $Desktop "Contact_Import_Results_$Date.csv"

$results = @()
$created | ForEach-Object { $results += [PSCustomObject]@{Email=$_; Status="Created"} }
$skipped | ForEach-Object { $results += [PSCustomObject]@{Email=$_; Status="Already Exists"} }
$errors | ForEach-Object { $results += [PSCustomObject]@{Email=$_; Status="Error"} }

$results | Export-Csv -Path $ResultsFile -NoTypeInformation -Encoding UTF8
Write-Host "`nDetailed results exported to: $ResultsFile" -ForegroundColor Green

# Next steps
if ($created.Count -gt 0) {
    Write-Host "`nNext steps:" -ForegroundColor Cyan
    Write-Host "1. Verify contacts in Exchange Admin Center > Recipients > Contacts" -ForegroundColor White
}

Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "`nScript completed!" -ForegroundColor Green
