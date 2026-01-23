# Import Members to Distribution List from CSV
# This script adds both internal users and external contacts to a DL
# CSV format: Single column with header "Email"

# Import required modules (installed via Tools.ps1)
Import-Module ExchangeOnlineManagement -ErrorAction Stop

Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline | Out-Null

Write-Host "`nThis script will add members (internal users and external contacts) to a Distribution List from a CSV file." -ForegroundColor Cyan
Write-Host "CSV format: Single column with header 'Email'" -ForegroundColor Yellow

# Prompt for target DL
$success = $false
while (-not $success) {
    $targetDL = Read-Host "`nEnter the Distribution List email address"
    
    # Verify DL exists
    try {
        $dlInfo = Get-DistributionGroup -Identity $targetDL -ErrorAction Stop
        Write-Host "Found Distribution List: $($dlInfo.DisplayName)" -ForegroundColor Green
        $success = $true
    } catch {
        Write-Host "Error: Distribution List '$targetDL' not found." -ForegroundColor Red
        $retry = Read-Host "Would you like to try again? [Y] Yes [N] No"
        if ($retry -notmatch "[yY]") {
            Write-Host "Script terminated." -ForegroundColor Yellow
            Disconnect-ExchangeOnline -Confirm:$false
            Exit
        }
    }
}

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
    $members = Import-Csv $csvPath
    Write-Host "`nFound $($members.Count) email address(es) in CSV file." -ForegroundColor Green
} catch {
    Write-Host "Error reading CSV file: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Show current member count
$currentMembers = Get-DistributionGroupMember -Identity $targetDL -ResultSize Unlimited
Write-Host "Current members in DL: $($currentMembers.Count)" -ForegroundColor Cyan

# Confirm before proceeding
$confirm = Read-Host "`nDo you want to add $($members.Count) member(s) to '$($dlInfo.DisplayName)'? [Y] Yes [N] No"
if ($confirm -notmatch "[yY]") {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Track results
$added = @()
$alreadyMember = @()
$errors = @()

Write-Host "`nStarting import process..." -ForegroundColor Cyan
# Start timer
$importStartTime = Get-Date
$currentCount = 0
foreach ($member in $members) {
    $currentCount++
    $email = $member.Email.Trim()
    
    if ([string]::IsNullOrWhiteSpace($email)) {
        Write-Host "[$currentCount/$($members.Count)] Skipping empty email address" -ForegroundColor Yellow
        continue
    }
    
    Write-Progress -Activity "Adding members to Distribution List" `
                   -Status "Processing $currentCount of $($members.Count): $email" `
                   -PercentComplete (($currentCount / $members.Count) * 100)
    
    try {
        # Try to add member (works for both internal users and external contacts)
        Add-DistributionGroupMember -Identity $targetDL -Member $email -ErrorAction Stop
        Write-Host "[$currentCount/$($members.Count)] Added: $email" -ForegroundColor Green
        $added += $email
    } catch {
        if ($_.Exception.Message -like "*already a member*") {
            Write-Host "[$currentCount/$($members.Count)] Already a member: $email" -ForegroundColor Yellow
            $alreadyMember += $email
        } else {
            Write-Host "[$currentCount/$($members.Count)] Error adding $email : $($_.Exception.Message)" -ForegroundColor Red
            $errors += "$email - $($_.Exception.Message)"
        }
    }
    
    Start-Sleep -Milliseconds 300  # Small delay to avoid throttling
}

Write-Progress -Activity "Adding members to Distribution List" -Completed

# Get updated member count
$updatedMembers = Get-DistributionGroupMember -Identity $targetDL -ResultSize Unlimited

# Calculate elapsed time
$importEndTime = Get-Date
$elapsedTime = $importEndTime - $importStartTime
$timeFormatted = "{0:D2}:{1:D2}:{2:D2}" -f $elapsedTime.Hours, $elapsedTime.Minutes, $elapsedTime.Seconds

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Import Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Distribution List: $($dlInfo.DisplayName)" -ForegroundColor White
Write-Host "Members before: $($currentMembers.Count)" -ForegroundColor White
Write-Host "Members after: $($updatedMembers.Count)" -ForegroundColor White
Write-Host "`nTime elapsed: $timeFormatted (HH:MM:SS)" -ForegroundColor Cyan
Write-Host "Successfully added: $($added.Count)" -ForegroundColor Green
Write-Host "Already members (skipped): $($alreadyMember.Count)" -ForegroundColor Yellow
Write-Host "Errors: $($errors.Count)" -ForegroundColor Red

if ($errors.Count -gt 0) {
    Write-Host "`nErrors encountered:" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
}

# Export results
$Desktop = [Environment]::GetFolderPath("Desktop")
$Date = Get-Date -Format "yyyy-MM-dd_HHmmss"
$ResultsFile = Join-Path $Desktop "DL_Import_Results_$($dlInfo.Alias)_$Date.csv"

$results = @()
$added | ForEach-Object { $results += [PSCustomObject]@{Email=$_; Status="Added"} }
$alreadyMember | ForEach-Object { $results += [PSCustomObject]@{Email=$_; Status="Already Member"} }
$errors | ForEach-Object { $results += [PSCustomObject]@{Email=$_; Status="Error"} }

if ($results.Count -gt 0) {
    $results | Export-Csv -Path $ResultsFile -NoTypeInformation -Encoding UTF8
    Write-Host "`nDetailed results exported to: $ResultsFile" -ForegroundColor Green
}

# Show sample of current members
if ($updatedMembers.Count -le 10) {
    Write-Host "`nCurrent DL members:" -ForegroundColor Cyan
    $updatedMembers | Select-Object DisplayName, PrimarySmtpAddress, RecipientType | Format-Table -AutoSize
} else {
    Write-Host "`nShowing first 10 members (total: $($updatedMembers.Count)):" -ForegroundColor Cyan
    $updatedMembers | Select-Object DisplayName, PrimarySmtpAddress, RecipientType -First 10 | Format-Table -AutoSize
}
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "`nScript completed!" -ForegroundColor Green
