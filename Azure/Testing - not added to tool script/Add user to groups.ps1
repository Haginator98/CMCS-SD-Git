# Check if Microsoft.Graph module is installed
$Module = Get-Module -Name Microsoft.Graph -ListAvailable
if ($Module.Count -eq 0) {
    Write-Host "Microsoft.Graph module is not available." -ForegroundColor Yellow
    $Confirm = Read-Host "Are you sure you want to install the module? [Y] Yes [N] No"
    if ($Confirm -match "[yY]") {
        Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Cyan
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
        if ($?) {
            Write-Host "Microsoft.Graph module installed successfully." -ForegroundColor Green
        } else {
            Write-Host "Failed to install Microsoft.Graph module." -ForegroundColor Red
            Exit
        }
    } else {
        Write-Host "Microsoft.Graph module is required. Please install it using Install-Module Microsoft.Graph cmdlet." -ForegroundColor Red
        Exit
    }
}

Write-Host "Importing Microsoft.Graph module..." -ForegroundColor Yellow
Import-Module Microsoft.Graph

# Start timer
$scriptStartTime = Get-Date

Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All", "GroupMember.ReadWrite.All" -NoWelcome

Write-Host "This script will add a user to one or more groups in Entra ID." -ForegroundColor Cyan
Start-Sleep -Seconds 1

# Prompt for user
$userUPN = Read-Host "Enter the user's email address (UPN)"

# Verify user exists
try {
    $user = Get-MgUser -Filter "userPrincipalName eq '$userUPN'" -ErrorAction Stop
    
    if (-not $user) {
        Write-Host "Error: User '$userUPN' not found in Entra ID." -ForegroundColor Red
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Exit
    }
    
    Write-Host "Found user: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Green
} catch {
    Write-Host "Error: Failed to find user. $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# Select input method
Write-Host "`nHow do you want to specify groups?" -ForegroundColor Yellow
Write-Host "[1] Enter group names/emails manually (one or more)"
Write-Host "[2] Import from CSV file"
$inputMethod = Read-Host "Choose option (1 or 2)"

$groupsToAdd = @()

if ($inputMethod -eq "1") {
    # Manual entry
    Write-Host "`nEnter group names or email addresses (one per line)." -ForegroundColor Cyan
    Write-Host "Press Enter on empty line when done." -ForegroundColor Cyan
    
    while ($true) {
        $groupInput = Read-Host "Group name/email"
        if ([string]::IsNullOrWhiteSpace($groupInput)) { break }
        $groupsToAdd += $groupInput
    }
    
    if ($groupsToAdd.Count -eq 0) {
        Write-Host "No groups specified. Exiting..." -ForegroundColor Yellow
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Exit
    }
    
} elseif ($inputMethod -eq "2") {
    # CSV import
    $Desktop = [Environment]::GetFolderPath("Desktop")
    $csvFile = Read-Host "Enter the full path to the CSV file (or press Enter to browse from Desktop)"
    if ([string]::IsNullOrWhiteSpace($csvFile)) {
        $csvFileName = Read-Host "Enter the filename on your Desktop (e.g., groups.csv)"
        $csvFile = Join-Path $Desktop $csvFileName
    }
    
    if (-not (Test-Path $csvFile)) {
        Write-Host "Error: CSV file not found at: $csvFile" -ForegroundColor Red
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Exit
    }
    
    try {
        $csvData = Import-Csv $csvFile -ErrorAction Stop
        Write-Host "Expected CSV column: GroupName or GroupEmail" -ForegroundColor Yellow
        
        foreach ($row in $csvData) {
            if ($row.GroupName) {
                $groupsToAdd += $row.GroupName
            } elseif ($row.GroupEmail) {
                $groupsToAdd += $row.GroupEmail
            }
        }
        
        if ($groupsToAdd.Count -eq 0) {
            Write-Host "No groups found in CSV. Make sure you have a 'GroupName' or 'GroupEmail' column." -ForegroundColor Red
            Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
            Disconnect-MgGraph | Out-Null
            Exit
        }
        
        Write-Host "Found $($groupsToAdd.Count) groups in CSV file." -ForegroundColor Green
    } catch {
        Write-Host "Error: Failed to read CSV file. $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Exit
    }
} else {
    Write-Host "Invalid option. Exiting..." -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# Confirmation
Write-Host "`n=== CONFIRMATION ===" -ForegroundColor Yellow
Write-Host "User: $($user.DisplayName) ($($user.UserPrincipalName))"
Write-Host "Groups to add ($($groupsToAdd.Count)):"
foreach ($grp in $groupsToAdd) {
    Write-Host "  - $grp" -ForegroundColor Cyan
}

$confirm = Read-Host "`nDo you want to add the user to these groups? [Y] Yes [N] No"
if ($confirm -notmatch "[yY]") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# Add user to groups
$successCount = 0
$failCount = 0
$failedGroups = @()
$currentCount = 0

Write-Host "`nAdding user to groups..." -ForegroundColor Cyan

foreach ($groupIdentifier in $groupsToAdd) {
    $currentCount++
    Write-Progress -Activity "Adding user to groups" -Status "Processing $currentCount of $($groupsToAdd.Count)" -PercentComplete (($currentCount / $groupsToAdd.Count) * 100)
    
    try {
        # Try to find group by DisplayName or Mail
        $group = Get-MgGroup -Filter "displayName eq '$groupIdentifier' or mail eq '$groupIdentifier'" -ErrorAction Stop | Select-Object -First 1
        
        if (-not $group) {
            $failCount++
            $failedGroups += [PSCustomObject]@{
                GroupName = $groupIdentifier
                Error = "Group not found"
            }
            continue
        }
        
        # Add user to group
        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop
        $successCount++
        
    } catch {
        $failCount++
        $errorMessage = $_.Exception.Message
        
        # Check if user is already a member
        if ($errorMessage -like "*already exist*" -or $errorMessage -like "*already a member*") {
            $errorMessage = "User is already a member of this group"
        }
        
        $failedGroups += [PSCustomObject]@{
            GroupName = $groupIdentifier
            Error = $errorMessage
        }
    }
}

Write-Progress -Activity "Adding user to groups" -Completed

# Summary
Write-Host "`n=== GROUP MEMBERSHIP SUMMARY ===" -ForegroundColor Cyan
Write-Host "Total groups processed: $($groupsToAdd.Count)" -ForegroundColor White
Write-Host "Successfully added: $successCount" -ForegroundColor Green
Write-Host "Failed: $failCount" -ForegroundColor Red

# List failed groups
if ($failCount -gt 0) {
    Write-Host "`n=== FAILED GROUPS LIST ===" -ForegroundColor Yellow
    foreach ($failed in $failedGroups) {
        Write-Host "  - $($failed.GroupName)" -ForegroundColor Red
        Write-Host "    Reason: $($failed.Error)" -ForegroundColor DarkGray
    }
}

# Calculate total execution time
$scriptEndTime = Get-Date
$executionTime = $scriptEndTime - $scriptStartTime
$minutes = [math]::Floor($executionTime.TotalMinutes)
$seconds = $executionTime.Seconds

# Disconnect from Microsoft Graph
Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected successfully." -ForegroundColor Green

# Display execution time
Write-Host "`n=== SCRIPT COMPLETED ===" -ForegroundColor Cyan
if ($minutes -gt 0) {
    Write-Host "Total execution time: $minutes minutes and $seconds seconds" -ForegroundColor White
} else {
    Write-Host "Total execution time: $seconds seconds" -ForegroundColor White
}
Start-Sleep -Seconds 3
