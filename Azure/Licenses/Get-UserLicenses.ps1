# Check which Entra ID / Microsoft 365 licenses one or more users have assigned
$requiredGraphModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "Microsoft.Graph.Identity.DirectoryManagement")

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
Import-Module Microsoft.Graph.Identity.DirectoryManagement -RequiredVersion $graphModuleVersion -ErrorAction Stop

$scriptStartTime = Get-Date

Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All", "Organization.Read.All" -NoWelcome

Write-Host "This script checks which licenses one or more users have assigned in Entra ID." -ForegroundColor Cyan
Start-Sleep -Seconds 1

# --- Build SKU ID -> friendly name lookup ---
Write-Host "`nRetrieving available license SKUs in tenant..." -ForegroundColor Cyan
$skuLookup = @{}
try {
    $subscribedSkus = Get-MgSubscribedSku -ErrorAction Stop
    foreach ($sku in $subscribedSkus) {
        $skuLookup[$sku.SkuId] = $sku.SkuPartNumber
    }
} catch {
    Write-Host "Warning: Failed to retrieve SKU list. License names will show as raw SKU IDs. $($_.Exception.Message)" -ForegroundColor Yellow
}

# --- Select users to check ---
Write-Host "`nHow do you want to specify the users to check?" -ForegroundColor Yellow
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
            $usersCsvFileName = Read-Host "Enter the filename on your Desktop (e.g., Get-UserLicenses_Users_Template.csv)"
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

# --- Resolve users and read their licenses ---
Write-Host "`nChecking licenses for $($userIdentifiers.Count) user(s)..." -ForegroundColor Cyan
$results = @()
$unresolvedUsers = @()
$currentCount = 0

foreach ($identifier in $userIdentifiers) {
    $currentCount++
    Write-Progress -Activity "Checking user licenses" -Status "Processing $currentCount of $($userIdentifiers.Count)" -PercentComplete (($currentCount / $userIdentifiers.Count) * 100)

    try {
        $user = Get-MgUser -Filter "userPrincipalName eq '$identifier' or mail eq '$identifier'" -Property Id, DisplayName, UserPrincipalName, AssignedLicenses -ConsistencyLevel eventual -CountVariable userCount -ErrorAction Stop | Select-Object -First 1

        if (-not $user) {
            $unresolvedUsers += "$identifier (not found)"
            continue
        }

        if ($user.AssignedLicenses.Count -eq 0) {
            $results += [PSCustomObject]@{
                DisplayName       = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Licenses          = "no licenses assigned"
            }
        } else {
            $licenseNames = foreach ($license in $user.AssignedLicenses) {
                $skuName = $skuLookup[$license.SkuId]
                if ([string]::IsNullOrEmpty($skuName)) { $license.SkuId } else { $skuName }
            }
            $results += [PSCustomObject]@{
                DisplayName       = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Licenses          = ($licenseNames -join ", ")
            }
        }
    } catch {
        $unresolvedUsers += "$identifier ($($_.Exception.Message))"
    }
}

Write-Progress -Activity "Checking user licenses" -Completed

# --- Display results ---
Write-Host "`n=== LICENSE REPORT ===" -ForegroundColor Cyan
foreach ($result in $results) {
    Write-Host "`n$($result.DisplayName) ($($result.UserPrincipalName))" -ForegroundColor White
    Write-Host "  Licenses: $($result.Licenses)" -ForegroundColor $(if ($result.Licenses -eq "no licenses assigned") { "Yellow" } else { "Green" })
}

if ($unresolvedUsers.Count -gt 0) {
    Write-Host "`nWarning: The following users could not be found and were skipped:" -ForegroundColor Yellow
    foreach ($u in $unresolvedUsers) {
        Write-Host "  - $u" -ForegroundColor Yellow
    }
}

Write-Host "`nChecked $($userIdentifiers.Count) user(s): $($results.Count) found, $($unresolvedUsers.Count) not found." -ForegroundColor Cyan

# --- Optional export ---
if ($results.Count -gt 0) {
    $exportChoice = Read-Host "`nDo you want to export this license report to a CSV file on your Desktop? [Y] Yes [N] No"
    if ($exportChoice -match "[yY]") {
        $Desktop = [Environment]::GetFolderPath("Desktop")
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $exportPath = Join-Path $Desktop "UserLicenseReport_$timestamp.csv"
        try {
            $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            Write-Host "License report exported to: $exportPath" -ForegroundColor Green
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
