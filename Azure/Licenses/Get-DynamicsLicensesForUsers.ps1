# Find users with duplicate Dynamics base licenses that could be optimized with attach licenses
# Requires: Microsoft.Graph PowerShell module

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "$([Environment]::GetFolderPath('Desktop'))\DynamicsLicenseDuplicates_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Define Dynamics base licenses that should not be combined
# These are the actual SKU PartNumbers from your tenant
$DynamicsBaseLicenses = @{
    'DYN365_FINANCE' = 'Dynamics 365 Finance'
    'DYN365_SCM' = 'Dynamics 365 Supply Chain Management'
    'DYN365_ENTERPRISE_SALES' = 'Dynamics 365 Sales Enterprise'
    'DYN365_ENTERPRISE_CUSTOMER_SERVICE' = 'Dynamics 365 Customer Service Enterprise'
    'DYN365_PROJECT_OPERATIONS' = 'Dynamics 365 Project Operations'
    'DYN365_ENTERPRISE_PLAN1' = 'Dynamics 365 Plan 1'
}

# Define common optimization scenarios
$OptimizationScenarios = @(
    @{
        Name = "Finance + Sales Enterprise"
        Licenses = @('DYN365_FINANCE', 'DYN365_ENTERPRISE_SALES')
        Recommendation = "Keep Finance (base) + Sales Enterprise Attach"
    },
    @{
        Name = "Sales Enterprise + Customer Service Enterprise"
        Licenses = @('DYN365_ENTERPRISE_SALES', 'DYN365_ENTERPRISE_CUSTOMER_SERVICE')
        Recommendation = "Keep Sales Enterprise (base) + Customer Service Enterprise Attach"
    },
    @{
        Name = "Finance + Supply Chain Management"
        Licenses = @('DYN365_FINANCE', 'DYN365_SCM')
        Recommendation = "Keep Finance (base) + Supply Chain Management Attach"
    },
    @{
        Name = "Finance + Project Operations"
        Licenses = @('DYN365_FINANCE', 'DYN365_PROJECT_OPERATIONS')
        Recommendation = "Keep Finance (base) + Project Operations Attach"
    },
    @{
        Name = "Supply Chain + Sales Enterprise"
        Licenses = @('DYN365_SCM', 'DYN365_ENTERPRISE_SALES')
        Recommendation = "Keep Supply Chain (base) + Sales Enterprise Attach"
    },
    @{
        Name = "Supply Chain + Project Operations"
        Licenses = @('DYN365_SCM', 'DYN365_PROJECT_OPERATIONS')
        Recommendation = "Keep Supply Chain (base) + Project Operations Attach"
    },
    @{
        Name = "Multiple Dynamics 365 Base Licenses (3+)"
        Licenses = @() # Will catch any user with 2+ base licenses
        Recommendation = "Keep the most used application as base + convert others to attach licenses"
    }
)

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Dynamics 365 License Cost Optimization" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Import required modules (installed via Tools.ps1)
Write-Host "Importing required modules..." -ForegroundColor Yellow
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
try {
    Connect-MgGraph -Scopes "User.Read.All", "Organization.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "Successfully connected to Microsoft Graph`n" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    exit
}

# Get all available SKUs in the tenant
Write-Host "Retrieving available licenses in tenant..." -ForegroundColor Yellow
$availableSkus = Get-MgSubscribedSku -All
Write-Host "Found $($availableSkus.Count) license types`n" -ForegroundColor Green

# Create a lookup table for SKU IDs to names
$skuLookup = @{}
foreach ($sku in $availableSkus) {
    $skuLookup[$sku.SkuId] = $sku.SkuPartNumber
}

# Ask if user wants to check one user or all users
$scopeChoice = ""
while ($scopeChoice -ne "1" -and $scopeChoice -ne "2") {
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "SCOPE SELECTION" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "[1] Check single user" -ForegroundColor White
    Write-Host "[2] Check all users`n" -ForegroundColor White

    $scopeChoice = Read-Host "Select option [1] or [2]"
    
    if ($scopeChoice -ne "1" -and $scopeChoice -ne "2") {
        Write-Host "Invalid option. Please select 1 or 2.`n" -ForegroundColor Red
    }
}

if ($scopeChoice -eq "1") {
    # Single user mode
    $userFound = $false
    
    while (-not $userFound) {
        $userPrincipalName = Read-Host "`nEnter UserPrincipalName (e.g., user@domain.com)"
        
        Write-Host "Retrieving user: $userPrincipalName..." -ForegroundColor Yellow
        try {
            $users = @(Get-MgUser -UserId $userPrincipalName -Property Id,DisplayName,UserPrincipalName,AssignedLicenses -ErrorAction Stop)
            
            # Check if user has any licenses
            if ($users[0].AssignedLicenses.Count -eq 0) {
                Write-Host "User found, but has no assigned licenses." -ForegroundColor Yellow
                Write-Host "This user cannot be checked for duplicate Dynamics licenses.`n" -ForegroundColor Yellow
                
                $retry = Read-Host "[1] Try another user [2] Exit"
                if ($retry -ne "1") {
                    Write-Host "`nExiting script." -ForegroundColor Yellow
                    Disconnect-MgGraph | Out-Null
                    exit
                }
            }
            else {
                Write-Host "User found`n" -ForegroundColor Green
                $userCount = 1
                $userFound = $true
            }
        }
        catch {
            Write-Host "User not found or error occurred: $_" -ForegroundColor Red
            Write-Host "Did you spell the UserPrincipalName correctly?" -ForegroundColor Yellow
            
            $retry = Read-Host "`n[1] Try again [2] Exit"
            if ($retry -ne "1") {
                Write-Host "`nExiting script." -ForegroundColor Yellow
                Disconnect-MgGraph | Out-Null
                exit
            }
        }
    }
}
elseif ($scopeChoice -eq "2") {
    # All users mode
    Write-Host "`nRetrieving all users with assigned licenses..." -ForegroundColor Yellow
    $users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AssignedLicenses -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable userCount
    Write-Host "Found $userCount licensed users`n" -ForegroundColor Green
}

# For single user mode, show all licenses before analyzing
if ($scopeChoice -eq "1") {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "USER LICENSE DETAILS" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "User: $($users[0].DisplayName) ($($users[0].UserPrincipalName))" -ForegroundColor White
    Write-Host "Total licenses assigned: $($users[0].AssignedLicenses.Count)`n" -ForegroundColor White
    
    # Show all licenses
    foreach ($license in $users[0].AssignedLicenses) {
        $skuPartNumber = $skuLookup[$license.SkuId]
        if ($skuPartNumber) {
            $isDynamicsBase = $DynamicsBaseLicenses.ContainsKey($skuPartNumber)
            $color = if ($isDynamicsBase) { "Yellow" } else { "Gray" }
            $marker = if ($isDynamicsBase) { " [DYNAMICS BASE LICENSE]" } else { "" }
            Write-Host "  - $skuPartNumber$marker" -ForegroundColor $color
        }
    }
    Write-Host ""
}

# For single user mode, analyze immediately without confirmation
if ($scopeChoice -eq "1") {
    Write-Host "Analyzing for cost optimization opportunities..." -ForegroundColor Yellow
    
    # Get user's license SKU PartNumbers
    $userLicenseSKUs = @()
    $userLicenseNames = @()
    
    foreach ($license in $users[0].AssignedLicenses) {
        $skuPartNumber = $skuLookup[$license.SkuId]
        if ($skuPartNumber) {
            $userLicenseSKUs += $skuPartNumber
            if ($DynamicsBaseLicenses.ContainsKey($skuPartNumber)) {
                $userLicenseNames += $DynamicsBaseLicenses[$skuPartNumber]
            }
        }
    }
    
    # Check if user has multiple Dynamics base licenses
    $dynamicsBaseLicensesFound = $userLicenseSKUs | Where-Object { $DynamicsBaseLicenses.ContainsKey($_) }
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "ANALYSIS RESULT" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    if ($dynamicsBaseLicensesFound.Count -ge 2) {
        Write-Host "COST OPTIMIZATION OPPORTUNITY FOUND!" -ForegroundColor Red
        Write-Host "`nThis user has $($dynamicsBaseLicensesFound.Count) Dynamics base licenses:" -ForegroundColor Yellow
        foreach ($lic in $userLicenseNames) {
            Write-Host "  - $lic" -ForegroundColor Yellow
        }
        
        # Determine best base license to keep (priority order)
        $basePriority = @('DYN365_FINANCE', 'DYN365_SCM', 'DYN365_ENTERPRISE_SALES', 'DYN365_ENTERPRISE_CUSTOMER_SERVICE', 'DYN365_PROJECT_OPERATIONS', 'DYN365_ENTERPRISE_PLAN1')
        $recommendedBase = $null
        foreach ($priority in $basePriority) {
            if ($dynamicsBaseLicensesFound -contains $priority) {
                $recommendedBase = $DynamicsBaseLicenses[$priority]
                break
            }
        }
        if (-not $recommendedBase) { $recommendedBase = $userLicenseNames[0] }
        
        # Find which optimization scenario applies
        $recommendation = "Keep $recommendedBase (base) + convert remaining $($dynamicsBaseLicensesFound.Count - 1) to attach licenses"
        $scenarioFound = $false
        
        foreach ($optScenario in $OptimizationScenarios) {
            if ($optScenario.Licenses.Count -eq 2) {
                $hasAll = $true
                foreach ($reqLicense in $optScenario.Licenses) {
                    if ($reqLicense -notin $dynamicsBaseLicensesFound) {
                        $hasAll = $false
                        break
                    }
                }
                if ($hasAll -and $dynamicsBaseLicensesFound.Count -eq 2) {
                    $recommendation = $optScenario.Recommendation
                    $scenarioFound = $true
                    break
                }
            }
        }
        
        Write-Host "`nRECOMMENDATION:" -ForegroundColor Green
        Write-Host "$recommendation" -ForegroundColor White
        Write-Host "`nPOTENTIAL SAVINGS:" -ForegroundColor Green
        Write-Host "By using attach licenses, you could save the cost of $(($dynamicsBaseLicensesFound.Count - 1)) base license(s).`n" -ForegroundColor White
    }
    elseif ($dynamicsBaseLicensesFound.Count -eq 1) {
        Write-Host "License allocation is optimized." -ForegroundColor Green
        Write-Host "User has only 1 Dynamics base license: $($userLicenseNames -join ', ')" -ForegroundColor White
        Write-Host "No cost optimization needed.`n" -ForegroundColor White
    }
    else {
        Write-Host "No Dynamics 365 base licenses found for this user." -ForegroundColor Green
        Write-Host "No cost optimization opportunities.`n" -ForegroundColor White
    }
    
    # Disconnect and exit for single user mode
    Disconnect-MgGraph | Out-Null
    Write-Host "Disconnected from Microsoft Graph`n" -ForegroundColor Gray
    exit
}

# Ask for confirmation before analyzing (all users mode only)
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "READY TO ANALYZE" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "This will analyze $userCount users for cost optimization opportunities." -ForegroundColor White
Write-Host "Looking for users with multiple Dynamics base licenses that could use attach licenses." -ForegroundColor White
Write-Host "Results will be exported to: $ExportPath`n" -ForegroundColor White

$confirmation = Read-Host "Do you want to continue with the analysis? [Y] Yes [N] No"
if ($confirmation -notmatch "[yY]") {
    Write-Host "`nAnalysis cancelled by user." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit
}
Write-Host "" # Empty line for readability

# Analyze users for duplicate base licenses
Write-Host "Analyzing users for cost optimization opportunities..." -ForegroundColor Yellow
$results = @()
$userCounter = 0

foreach ($user in $users) {
    $userCounter++
    Write-Progress -Activity "Analyzing users" -Status "Processing $userCounter of $userCount" -PercentComplete (($userCounter / $userCount) * 100)
    
    # Get user's license SKU PartNumbers
    $userLicenseSKUs = @()
    $userLicenseNames = @()
    
    foreach ($license in $user.AssignedLicenses) {
        $skuPartNumber = $skuLookup[$license.SkuId]
        if ($skuPartNumber) {
            $userLicenseSKUs += $skuPartNumber
            if ($DynamicsBaseLicenses.ContainsKey($skuPartNumber)) {
                $userLicenseNames += $DynamicsBaseLicenses[$skuPartNumber]
            }
        }
    }
    
    # Check if user has multiple Dynamics base licenses
    $dynamicsBaseLicensesFound = $userLicenseSKUs | Where-Object { $DynamicsBaseLicenses.ContainsKey($_) }
    
    if ($dynamicsBaseLicensesFound.Count -ge 2) {
        # Determine best base license to keep (priority order)
        $basePriority = @('DYN365_FINANCE', 'DYN365_SCM', 'DYN365_ENTERPRISE_SALES', 'DYN365_ENTERPRISE_CUSTOMER_SERVICE', 'DYN365_PROJECT_OPERATIONS', 'DYN365_ENTERPRISE_PLAN1')
        $recommendedBase = $null
        foreach ($priority in $basePriority) {
            if ($dynamicsBaseLicensesFound -contains $priority) {
                $recommendedBase = $DynamicsBaseLicenses[$priority]
                break
            }
        }
        if (-not $recommendedBase) { $recommendedBase = $userLicenseNames[0] }
        
        # Find which optimization scenario applies
        $scenario = "Multiple Dynamics Base Licenses Detected"
        $recommendation = "Keep $recommendedBase (base) + convert remaining $($dynamicsBaseLicensesFound.Count - 1) to attach licenses"
        
        foreach ($optScenario in $OptimizationScenarios) {
            if ($optScenario.Licenses.Count -eq 2) {
                $hasAll = $true
                foreach ($reqLicense in $optScenario.Licenses) {
                    if ($reqLicense -notin $dynamicsBaseLicensesFound) {
                        $hasAll = $false
                        break
                    }
                }
                if ($hasAll -and $dynamicsBaseLicensesFound.Count -eq 2) {
                    $scenario = $optScenario.Name
                    $recommendation = $optScenario.Recommendation
                    break
                }
            }
        }
        
        $results += [PSCustomObject]@{
            DisplayName = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            ConflictingLicenses = ($userLicenseNames -join "; ")
            Recommendation = $recommendation
        }
    }
}

Write-Progress -Activity "Analyzing users" -Completed

# Display summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Total licensed users analyzed: " -NoNewline
Write-Host $userCount -ForegroundColor Yellow

Write-Host "Users with cost optimization opportunities: " -NoNewline
Write-Host $results.Count -ForegroundColor $(if ($results.Count -gt 0) { "Red" } else { "Green" })

if ($results.Count -gt 0) {
    Write-Host "`nBreakdown by scenario:" -ForegroundColor Cyan
    $results | Group-Object Scenario | ForEach-Object {
        Write-Host "  - $($_.Name): " -NoNewline -ForegroundColor White
        Write-Host "$($_.Count) users" -ForegroundColor Yellow
    }
    
    # Ask for confirmation before exporting
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "EXPORT RESULTS" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Export path: $ExportPath" -ForegroundColor White
    Write-Host "Total records: $($results.Count)`n" -ForegroundColor White
    
    $exportConfirmation = Read-Host "Do you want to export the results to CSV? [Y] Yes [N] No"
    if ($exportConfirmation -match "[yY]") {
        Write-Host "`nExporting results to CSV..." -ForegroundColor Yellow
        $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $ExportPath" -ForegroundColor Green
    }
    else {
        Write-Host "`nExport cancelled by user." -ForegroundColor Yellow
    }
    
    # Show sample results
    Write-Host "`nSample results (first 5 users):" -ForegroundColor Cyan
    $results | Select-Object -First 5 DisplayName, UserPrincipalName, ConflictingLicenses, Recommendation | Format-Table -AutoSize
    
    if ($results.Count -gt 5) {
        Write-Host "... and $($results.Count - 5) more. See CSV for full list.`n" -ForegroundColor Yellow
    }
    
    Write-Host "`nPOTENTIAL COST SAVINGS:" -ForegroundColor Green
    Write-Host "By converting these users to attach licenses, you could potentially save" -ForegroundColor White
    Write-Host "one base license cost per user (typically significant savings per user/month).`n" -ForegroundColor White
}
else {
    Write-Host "`nNo cost optimization opportunities found. License allocation looks optimized!`n" -ForegroundColor Green
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected from Microsoft Graph`n" -ForegroundColor Gray
