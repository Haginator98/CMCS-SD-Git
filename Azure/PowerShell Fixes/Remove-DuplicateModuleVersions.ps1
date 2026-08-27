# Find PowerShell modules with multiple installed versions and remove all but the newest
$scriptStartTime = Get-Date

Write-Host "This script scans for PowerShell modules with multiple installed versions and removes all but the newest version." -ForegroundColor Cyan
Write-Host "This helps avoid 'assembly already loaded' errors caused by mismatched module versions (e.g. Microsoft.Graph.*)." -ForegroundColor Cyan

$moduleFilter = Read-Host "`nEnter a module name or wildcard pattern to scan (e.g. Microsoft.Graph*, or press Enter for all modules)"
if ([string]::IsNullOrWhiteSpace($moduleFilter)) {
    $moduleFilter = "*"
}

Write-Host "`nScanning installed modules matching '$moduleFilter'..." -ForegroundColor Cyan
try {
    # -AllVersions only works against a single, exact module name (no wildcard),
    # so first resolve the distinct matching module names, then query each one individually
    $moduleNames = Get-InstalledModule -Name $moduleFilter -ErrorAction Stop | Select-Object -ExpandProperty Name -Unique

    $allModules = @()
    foreach ($name in $moduleNames) {
        $allModules += Get-InstalledModule -Name $name -AllVersions -ErrorAction Stop
    }
} catch {
    Write-Host "Error: Failed to query installed modules. $($_.Exception.Message)" -ForegroundColor Red
    Exit
}

if (-not $allModules) {
    Write-Host "No installed modules found matching '$moduleFilter'." -ForegroundColor Yellow
    Exit
}

# Group by module name and keep only modules that have more than one installed version
$duplicateGroups = $allModules | Group-Object Name | Where-Object { $_.Count -gt 1 }

if ($duplicateGroups.Count -eq 0) {
    Write-Host "`nNo duplicate versions found. All modules matching '$moduleFilter' have only one version installed." -ForegroundColor Green
    Exit
}

# Build the list of older versions to remove (everything except the newest per module)
$toRemove = @()
foreach ($group in $duplicateGroups) {
    $sorted = $group.Group | Sort-Object { [version]$_.Version } -Descending
    $newest = $sorted[0]
    $older = $sorted | Select-Object -Skip 1

    Write-Host "`nModule: $($group.Name)" -ForegroundColor Cyan
    Write-Host "  Keeping newest version: $($newest.Version)" -ForegroundColor Green
    foreach ($old in $older) {
        Write-Host "  Will remove version: $($old.Version)" -ForegroundColor Yellow
        $toRemove += $old
    }
}

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "SUMMARY: $($toRemove.Count) old module version(s) will be uninstalled across $($duplicateGroups.Count) module(s)." -ForegroundColor Yellow
Write-Host "========================================`n" -ForegroundColor Yellow

$confirm = Read-Host "Do you want to proceed with uninstalling these older versions? [Y] Yes [N] No"
if ($confirm -notmatch "[yY]") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    Exit
}

$successCount = 0
$failCount = 0
$failedRemovals = @()
$currentCount = 0

Write-Host "`nUninstalling old module versions..." -ForegroundColor Cyan

foreach ($old in $toRemove) {
    $currentCount++
    Write-Progress -Activity "Uninstalling old module versions" -Status "Processing $currentCount of $($toRemove.Count)" -PercentComplete (($currentCount / $toRemove.Count) * 100)

    try {
        # A version can't be uninstalled while it's loaded in the current session
        Get-Module -Name $old.Name -All | Where-Object { $_.Version -eq $old.Version } | Remove-Module -Force -ErrorAction SilentlyContinue

        Uninstall-Module -Name $old.Name -RequiredVersion $old.Version -Force -ErrorAction Stop
        $successCount++
        Write-Host "Removed: $($old.Name) $($old.Version)" -ForegroundColor Green
    } catch {
        $failCount++
        $failedRemovals += [PSCustomObject]@{
            Module  = $old.Name
            Version = $old.Version
            Error   = $_.Exception.Message
        }
    }
}

Write-Progress -Activity "Uninstalling old module versions" -Completed

# Summary
Write-Host "`n=== CLEANUP SUMMARY ===" -ForegroundColor Cyan
Write-Host "Successfully removed: $successCount" -ForegroundColor Green
Write-Host "Failed: $failCount" -ForegroundColor Red

if ($failCount -gt 0) {
    Write-Host "`n=== FAILED REMOVALS ===" -ForegroundColor Yellow
    foreach ($failed in $failedRemovals) {
        Write-Host "  - $($failed.Module) $($failed.Version)" -ForegroundColor Red
        Write-Host "    Reason: $($failed.Error)" -ForegroundColor DarkGray
    }
    Write-Host "`nTip: If removal fails due to permission or 'in use' errors, close all PowerShell/VS Code terminal sessions and run this script again, or run it as administrator." -ForegroundColor Yellow
}

# Calculate total execution time
$scriptEndTime = Get-Date
$executionTime = $scriptEndTime - $scriptStartTime
$minutes = [math]::Floor($executionTime.TotalMinutes)
$seconds = $executionTime.Seconds

Write-Host "`n=== SCRIPT COMPLETED ===" -ForegroundColor Cyan
if ($minutes -gt 0) {
    Write-Host "Total execution time: $minutes minutes and $seconds seconds" -ForegroundColor White
} else {
    Write-Host "Total execution time: $seconds seconds" -ForegroundColor White
}
Start-Sleep -Seconds 3
