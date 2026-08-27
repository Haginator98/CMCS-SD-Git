# Remove one or more users from one or more Entra ID groups
# Use the newest version installed for any required submodule, and install it for the
# others if missing, so all three load the same assembly version in this session
$requiredGraphModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Groups")

$allVersions = $requiredGraphModules | ForEach-Object {
    Get-Module -ListAvailable -Name $_ | ForEach-Object { $_.Version.ToString() }
}
$graphModuleVersion = $allVersions | Sort-Object { [version]$_ } -Descending | Select-Object -First 1

if (-not $graphModuleVersion) {
    Write-Host "Error: None of the required modules are installed: $($requiredGraphModules -join ', ')" -ForegroundColor Red
    Exit
}

Write-Host "Using Microsoft.Graph module version $graphModuleVersion." -ForegroundColor Cyan

foreach ($moduleName in $requiredGraphModules) {
    $hasVersion = Get-Module -ListAvailable -Name $moduleName | Where-Object { $_.Version.ToString() -eq $graphModuleVersion }
    if (-not $hasVersion) {
        Write-Host "`n$moduleName version $graphModuleVersion is not installed." -ForegroundColor Yellow
        $confirm = Read-Host "Install $moduleName $graphModuleVersion now so all modules match? [Y] Yes [N] No"
        if ($confirm -notmatch "[yY]") {
            Write-Host "Cannot continue without matching module versions. Exiting..." -ForegroundColor Red
            Exit
        }
        try {
            Install-Module -Name $moduleName -RequiredVersion $graphModuleVersion -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "Installed $moduleName $graphModuleVersion." -ForegroundColor Green
        } catch {
            Write-Host "Error: Failed to install $moduleName $graphModuleVersion. $($_.Exception.Message)" -ForegroundColor Red
            Exit
        }
    }
}

Import-Module Microsoft.Graph.Authentication -RequiredVersion $graphModuleVersion -ErrorAction Stop
Import-Module Microsoft.Graph.Users -RequiredVersion $graphModuleVersion -ErrorAction Stop
Import-Module Microsoft.Graph.Groups -RequiredVersion $graphModuleVersion -ErrorAction Stop

$scriptStartTime = Get-Date

Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All", "GroupMember.ReadWrite.All" -NoWelcome

Write-Host "This script will remove one or more users from one or more groups in Entra ID." -ForegroundColor Cyan
Start-Sleep -Seconds 1

# --- Select users to remove ---
Write-Host "`nHow do you want to specify the users to remove?" -ForegroundColor Yellow
Write-Host "[1] Enter user emails/UPNs manually, separated by comma"
Write-Host "[2] Import from CSV file"
$userInputMethod = Read-Host "Choose option (1 or 2)"

$userIdentifiers = @()

if ($userInputMethod -eq "1") {
    $userInput = Read-Host "`nEnter user emails/UPNs separated by comma (e.g. user1@domain.com, user2@domain.com)"
    $userIdentifiers = $userInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    if ($userIdentifiers.Count -eq 0) {
        Write-Host "No users entered. Exiting..." -ForegroundColor Yellow
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Exit
    }

} elseif ($userInputMethod -eq "2") {
    $Desktop = [Environment]::GetFolderPath("Desktop")
    $usersCsvValid = $false

    while (-not $usersCsvValid) {
        $usersCsvFile = Read-Host "Enter the full path to the users CSV file (or press Enter to browse from Desktop)"
        if ([string]::IsNullOrWhiteSpace($usersCsvFile)) {
            $usersCsvFileName = Read-Host "Enter the filename on your Desktop (e.g., Remove-EntraGroupMembers_Users_Template.csv)"
            $usersCsvFile = Join-Path $Desktop $usersCsvFileName
        }

        if (-not (Test-Path $usersCsvFile)) {
            Write-Host "Error: CSV file not found at: $usersCsvFile" -ForegroundColor Red
            $retry = Read-Host "Do you want to try another file? [Y] Yes [N] No, cancel"
            if ($retry -notmatch "[yY]") {
                Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
                Disconnect-MgGraph | Out-Null
                Exit
            }
            continue
        }

        try {
            $usersCsvData = Import-Csv $usersCsvFile -Encoding UTF8 -ErrorAction Stop
            $usersCsvColumns = $usersCsvData[0].PSObject.Properties.Name

            # Accept either an 'Email' or a 'UserPrincipalName' column
            if ($usersCsvColumns -contains "Email") {
                $usersColumnName = "Email"
            } elseif ($usersCsvColumns -contains "UserPrincipalName") {
                $usersColumnName = "UserPrincipalName"
            } else {
                Write-Host "Error: CSV must contain an 'Email' or 'UserPrincipalName' column." -ForegroundColor Red
                Write-Host "Columns found: $($usersCsvColumns -join ', ')" -ForegroundColor Yellow
                $retry = Read-Host "Do you want to try another file? [Y] Yes [N] No, cancel"
                if ($retry -notmatch "[yY]") {
                    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
                    Disconnect-MgGraph | Out-Null
                    Exit
                }
                continue
            }

            $userIdentifiers = @()
            foreach ($row in $usersCsvData) {
                if ($row.$usersColumnName) {
                    $userIdentifiers += $row.$usersColumnName.Trim()
                }
            }

            if ($userIdentifiers.Count -eq 0) {
                Write-Host "No users found in the '$usersColumnName' column." -ForegroundColor Red
                $retry = Read-Host "Do you want to try another file? [Y] Yes [N] No, cancel"
                if ($retry -notmatch "[yY]") {
                    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
                    Disconnect-MgGraph | Out-Null
                    Exit
                }
                continue
            }

            Write-Host "Found $($userIdentifiers.Count) users in CSV file (column: $usersColumnName)." -ForegroundColor Green
            $usersCsvValid = $true
        } catch {
            Write-Host "Error: Failed to read CSV file. $($_.Exception.Message)" -ForegroundColor Red
            $retry = Read-Host "Do you want to try another file? [Y] Yes [N] No, cancel"
            if ($retry -notmatch "[yY]") {
                Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
                Disconnect-MgGraph | Out-Null
                Exit
            }
        }
    }
} else {
    Write-Host "Invalid option. Exiting..." -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# --- Select groups to remove from ---
Write-Host "`nHow do you want to specify the groups to remove the users from?" -ForegroundColor Yellow
Write-Host "[1] Enter group names/emails/object IDs manually, separated by comma"
Write-Host "[2] Import from CSV file"
$groupInputMethod = Read-Host "Choose option (1 or 2)"

$groupIdentifiers = @()

if ($groupInputMethod -eq "1") {
    $groupInput = Read-Host "`nEnter group names/emails/object IDs separated by comma (e.g. Group1, Group2@domain.com, 3f2504e0-4f89-11d3-9a0c-0305e82c3301)"
    $groupIdentifiers = $groupInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    if ($groupIdentifiers.Count -eq 0) {
        Write-Host "No groups entered. Exiting..." -ForegroundColor Yellow
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Exit
    }

} elseif ($groupInputMethod -eq "2") {
    $Desktop = [Environment]::GetFolderPath("Desktop")
    $groupsCsvValid = $false

    while (-not $groupsCsvValid) {
        $groupsCsvFile = Read-Host "Enter the full path to the groups CSV file (or press Enter to browse from Desktop)"
        if ([string]::IsNullOrWhiteSpace($groupsCsvFile)) {
            $groupsCsvFileName = Read-Host "Enter the filename on your Desktop (e.g., Remove-EntraGroupMembers_Groups_Template.csv)"
            $groupsCsvFile = Join-Path $Desktop $groupsCsvFileName
        }

        if (-not (Test-Path $groupsCsvFile)) {
            Write-Host "Error: CSV file not found at: $groupsCsvFile" -ForegroundColor Red
            $retry = Read-Host "Do you want to try another file? [Y] Yes [N] No, cancel"
            if ($retry -notmatch "[yY]") {
                Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
                Disconnect-MgGraph | Out-Null
                Exit
            }
            continue
        }

        try {
            $groupsCsvData = Import-Csv $groupsCsvFile -Encoding UTF8 -ErrorAction Stop
            $groupsCsvColumns = $groupsCsvData[0].PSObject.Properties.Name

            if ($groupsCsvColumns -notcontains "GroupName") {
                Write-Host "Error: CSV must contain a 'GroupName' column." -ForegroundColor Red
                Write-Host "Columns found: $($groupsCsvColumns -join ', ')" -ForegroundColor Yellow
                $retry = Read-Host "Do you want to try another file? [Y] Yes [N] No, cancel"
                if ($retry -notmatch "[yY]") {
                    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
                    Disconnect-MgGraph | Out-Null
                    Exit
                }
                continue
            }

            $groupIdentifiers = @()
            foreach ($row in $groupsCsvData) {
                if ($row.GroupName) {
                    $groupIdentifiers += $row.GroupName.Trim()
                }
            }

            if ($groupIdentifiers.Count -eq 0) {
                Write-Host "No groups found in the 'GroupName' column." -ForegroundColor Red
                $retry = Read-Host "Do you want to try another file? [Y] Yes [N] No, cancel"
                if ($retry -notmatch "[yY]") {
                    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
                    Disconnect-MgGraph | Out-Null
                    Exit
                }
                continue
            }

            Write-Host "Found $($groupIdentifiers.Count) groups in CSV file." -ForegroundColor Green
            $groupsCsvValid = $true
        } catch {
            Write-Host "Error: Failed to read CSV file. $($_.Exception.Message)" -ForegroundColor Red
            $retry = Read-Host "Do you want to try another file? [Y] Yes [N] No, cancel"
            if ($retry -notmatch "[yY]") {
                Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
                Disconnect-MgGraph | Out-Null
                Exit
            }
        }
    }
} else {
    Write-Host "Invalid option. Exiting..." -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# --- Resolve users against Entra ID ---
Write-Host "`nResolving users in Entra ID..." -ForegroundColor Cyan
$resolvedUsers = @()
$unresolvedUsers = @()

foreach ($identifier in $userIdentifiers) {
    try {
        $user = Get-MgUser -Filter "userPrincipalName eq '$identifier' or mail eq '$identifier'" -ConsistencyLevel eventual -CountVariable userCount -ErrorAction Stop | Select-Object -First 1
        if ($user) {
            $resolvedUsers += $user
        } else {
            $unresolvedUsers += "$identifier (not found)"
        }
    } catch {
        $unresolvedUsers += "$identifier ($($_.Exception.Message))"
    }
}

if ($unresolvedUsers.Count -gt 0) {
    Write-Host "`nWarning: The following users could not be found and will be skipped:" -ForegroundColor Yellow
    foreach ($u in $unresolvedUsers) {
        Write-Host "  - $u" -ForegroundColor Yellow
    }
}

if ($resolvedUsers.Count -eq 0) {
    Write-Host "`nNo valid users found. Exiting..." -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# --- Resolve groups against Entra ID ---
Write-Host "`nResolving groups in Entra ID..." -ForegroundColor Cyan
$resolvedGroups = @()
$unresolvedGroups = @()

foreach ($identifier in $groupIdentifiers) {
    try {
        $group = $null

        # Object IDs are GUIDs and must be looked up directly, not via $filter
        if ($identifier -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
            $group = Get-MgGroup -GroupId $identifier -ErrorAction SilentlyContinue
        }

        if (-not $group) {
            # ConsistencyLevel/CountVariable are required for 'or' filters spanning different properties (e.g. displayName/mail)
            $group = Get-MgGroup -Filter "displayName eq '$identifier' or mail eq '$identifier'" -ConsistencyLevel eventual -CountVariable groupCount -ErrorAction Stop | Select-Object -First 1
        }

        if ($group) {
            $resolvedGroups += $group
        } else {
            $unresolvedGroups += "$identifier (not found)"
        }
    } catch {
        $unresolvedGroups += "$identifier ($($_.Exception.Message))"
    }
}

if ($unresolvedGroups.Count -gt 0) {
    Write-Host "`nWarning: The following groups could not be found and will be skipped:" -ForegroundColor Yellow
    foreach ($g in $unresolvedGroups) {
        Write-Host "  - $g" -ForegroundColor Yellow
    }
}

if ($resolvedGroups.Count -eq 0) {
    Write-Host "`nNo valid groups found. Exiting..." -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# --- Confirmation ---
$totalOperations = $resolvedUsers.Count * $resolvedGroups.Count

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "REMOVAL SUMMARY" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "`nUsers to remove ($($resolvedUsers.Count)):"
foreach ($u in $resolvedUsers) {
    Write-Host "  - $($u.DisplayName) ($($u.UserPrincipalName))" -ForegroundColor Cyan
}
Write-Host "`nGroups to remove from ($($resolvedGroups.Count)):"
foreach ($g in $resolvedGroups) {
    Write-Host "  - $($g.DisplayName)" -ForegroundColor Cyan
}
Write-Host "`nTotal removal operations: $totalOperations" -ForegroundColor Yellow

$confirm = Read-Host "`nDo you want to proceed with removing these users from these groups? [Y] Yes [N] No"
if ($confirm -notmatch "[yY]") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# --- Remove users from groups ---
$successCount = 0
$failCount = 0
$failedOperations = @()
$currentCount = 0

Write-Host "`nRemoving users from groups..." -ForegroundColor Cyan

foreach ($group in $resolvedGroups) {
    foreach ($user in $resolvedUsers) {
        $currentCount++
        Write-Progress -Activity "Removing users from groups" -Status "Processing $currentCount of $totalOperations" -PercentComplete (($currentCount / $totalOperations) * 100)

        try {
            Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop
            $successCount++
        } catch {
            $failCount++
            $errorMessage = $_.Exception.Message

            if ($errorMessage -like "*does not exist*" -or $errorMessage -like "*not a member*") {
                $errorMessage = "User is not a member of this group"
            }

            $failedOperations += [PSCustomObject]@{
                User  = $user.UserPrincipalName
                Group = $group.DisplayName
                Error = $errorMessage
            }
        }
    }
}

Write-Progress -Activity "Removing users from groups" -Completed

# --- Summary ---
Write-Host "`n=== GROUP REMOVAL SUMMARY ===" -ForegroundColor Cyan
Write-Host "Total operations processed: $totalOperations" -ForegroundColor White
Write-Host "Successfully removed: $successCount" -ForegroundColor Green
Write-Host "Failed: $failCount" -ForegroundColor Red

if ($failCount -gt 0) {
    Write-Host "`n=== FAILED OPERATIONS LIST ===" -ForegroundColor Yellow
    foreach ($failed in $failedOperations) {
        Write-Host "  - $($failed.User) -> $($failed.Group)" -ForegroundColor Red
        Write-Host "    Reason: $($failed.Error)" -ForegroundColor DarkGray
    }

    $exportChoice = Read-Host "`nDo you want to export the list of failed users to a CSV file on your Desktop? [Y] Yes [N] No"
    if ($exportChoice -match "[yY]") {
        $Desktop = [Environment]::GetFolderPath("Desktop")
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $exportPath = Join-Path $Desktop "Remove-EntraGroupMembers_Failed_$timestamp.csv"
        try {
            $failedOperations | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            Write-Host "Failed operations exported to: $exportPath" -ForegroundColor Green
        } catch {
            Write-Host "Error: Failed to export CSV. $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# --- Execution time ---
$scriptEndTime = Get-Date
$executionTime = $scriptEndTime - $scriptStartTime
$minutes = [math]::Floor($executionTime.TotalMinutes)
$seconds = $executionTime.Seconds

Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected successfully." -ForegroundColor Green

Write-Host "`n=== SCRIPT COMPLETED ===" -ForegroundColor Cyan
if ($minutes -gt 0) {
    Write-Host "Total execution time: $minutes minutes and $seconds seconds" -ForegroundColor White
} else {
    Write-Host "Total execution time: $seconds seconds" -ForegroundColor White
}
Start-Sleep -Seconds 3
