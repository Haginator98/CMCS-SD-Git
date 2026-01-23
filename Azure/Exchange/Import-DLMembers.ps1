# Import required modules (installed via Tools.ps1)
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Start timer
$scriptStartTime = Get-Date

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline | Out-Null
Write-Host "This script will import members and owners to a Distribution List from CSV files." -ForegroundColor Cyan

# Prompt for target DL
$targetDL = Read-Host "Enter the target Distribution List email address"

# Verify DL exists
try {
    $dlInfo = Get-DistributionGroup -Identity $targetDL -ErrorAction Stop
    Write-Host "Found Distribution List: $($dlInfo.DisplayName)" -ForegroundColor Green
} catch {
    Write-Host "Error: Distribution List '$targetDL' not found." -ForegroundColor Red
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
    Exit
}

# Ask which files to import
Write-Host "`nWhat would you like to import?" -ForegroundColor Yellow
Write-Host "[1] Members only"
Write-Host "[2] Owners only"
Write-Host "[3] Both members and owners"
$choice = Read-Host "Enter your choice (1-3)"

$Desktop = [Environment]::GetFolderPath("Desktop")

# Import Members
if ($choice -eq "1" -or $choice -eq "3") {
    $membersFile = Read-Host "Enter the full path to the Members CSV file (or press Enter to browse from Desktop)"
    if ([string]::IsNullOrWhiteSpace($membersFile)) {
        $membersFileName = Read-Host "Enter the filename on your Desktop (e.g., DLName_Members_2025-12-11.csv)"
        $membersFile = Join-Path $Desktop $membersFileName
    }
    
    if (Test-Path $membersFile) {
        Write-Host "Importing members from: $membersFile" -ForegroundColor Cyan
        $members = Import-Csv $membersFile
        $totalMembers = $members.Count
        $successCount = 0
        $failCount = 0
        $failedMembers = @()
        $currentCount = 0
        
        foreach ($member in $members) {
            $currentCount++
            Write-Progress -Activity "Importing members" -Status "Processing $currentCount of $totalMembers" -PercentComplete (($currentCount / $totalMembers) * 100)
            
            try {
                Add-DistributionGroupMember -Identity $targetDL -Member $member.PrimarySmtpAddress -ErrorAction Stop
                $successCount++
            } catch {
                $failCount++
                $failedMembers += [PSCustomObject]@{
                    Email = $member.PrimarySmtpAddress
                    Error = $_.Exception.Message
                }
            }
        }
        
        Write-Progress -Activity "Importing members" -Completed
        
        # Summary
        Write-Host "`n=== MEMBERS IMPORT SUMMARY ===" -ForegroundColor Cyan
        Write-Host "Total members processed: $totalMembers" -ForegroundColor White
        Write-Host "Successfully added: $successCount" -ForegroundColor Green
        Write-Host "Failed: $failCount" -ForegroundColor Red
        
        # List failed members
        if ($failCount -gt 0) {
            Write-Host "`n=== FAILED MEMBERS LIST ===" -ForegroundColor Yellow
            foreach ($failed in $failedMembers) {
                Write-Host "  - $($failed.Email)" -ForegroundColor Red
                Write-Host "    Reason: $($failed.Error)" -ForegroundColor DarkGray
            }
        }
    } else {
        Write-Host "Members file not found: $membersFile" -ForegroundColor Red
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Exit
    }
}

# Import Owners
if ($choice -eq "2" -or $choice -eq "3") {
    $ownersFile = Read-Host "Enter the full path to the Owners CSV file (or press Enter to browse from Desktop)"
    if ([string]::IsNullOrWhiteSpace($ownersFile)) {
        $ownersFileName = Read-Host "Enter the filename on your Desktop (e.g., DLName_Owners_2025-12-11.csv)"
        $ownersFile = Join-Path $Desktop $ownersFileName
    }
    
    if (Test-Path $ownersFile) {
        Write-Host "`nImporting owners from: $ownersFile" -ForegroundColor Cyan
        $owners = Import-Csv $ownersFile
        $ownerEmails = $owners.PrimarySmtpAddress
        
        try {
            Set-DistributionGroup -Identity $targetDL -ManagedBy $ownerEmails -ErrorAction Stop
            Write-Host "`n=== OWNERS IMPORT SUMMARY ===" -ForegroundColor Cyan
            Write-Host "Owners updated successfully. Total owners: $($ownerEmails.Count)" -ForegroundColor Green
        } catch {
            Write-Host "`n=== OWNERS IMPORT FAILED ===" -ForegroundColor Red
            Write-Host "Failed to update owners." -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor DarkGray
            Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
            Disconnect-ExchangeOnline -Confirm:$false
            Exit
        }
    } else {
        Write-Host "Owners file not found: $ownersFile" -ForegroundColor Red
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false
        Exit
    }
}

# Calculate total execution time
$scriptEndTime = Get-Date
$executionTime = $scriptEndTime - $scriptStartTime
$minutes = [math]::Floor($executionTime.TotalMinutes)
$seconds = $executionTime.Seconds

# Disconnect from Exchange Online
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected successfully." -ForegroundColor Green

# Display execution time
Write-Host "`n=== SCRIPT COMPLETED ===" -ForegroundColor Cyan
if ($minutes -gt 0) {
    Write-Host "Total execution time: $minutes minutes and $seconds seconds" -ForegroundColor White
} else {
    Write-Host "Total execution time: $seconds seconds" -ForegroundColor White
}
Start-Sleep -Seconds 4
