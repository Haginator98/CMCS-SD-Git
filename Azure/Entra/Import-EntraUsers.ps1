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

# Write-Host "Importing Microsoft.Graph module..." -ForegroundColor Yellow
# Import-Module Microsoft.Graph

# Start timer
$scriptStartTime = Get-Date

Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome

Write-Host "This script will bulk import users to Entra ID from a CSV file." -ForegroundColor Cyan
Write-Host "NOTE: This script expects Azure AD CSV format (exported from Entra ID)." -ForegroundColor Yellow
Start-Sleep -Seconds 1

# Prompt for CSV file
$Desktop = [Environment]::GetFolderPath("Desktop")
$csvFile = Read-Host "Enter the full path to the CSV file (or press Enter to browse from Desktop)"
if ([string]::IsNullOrWhiteSpace($csvFile)) {
    $csvFileName = Read-Host "Enter the filename on your Desktop (e.g., users.csv)"
    $csvFile = Join-Path $Desktop $csvFileName
}

# Verify file exists
if (-not (Test-Path $csvFile)) {
    Write-Host "Error: CSV file not found at: $csvFile" -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# Import CSV
Write-Host "Importing users from: $csvFile" -ForegroundColor Cyan
try {
    # Read CSV and skip first line if it contains "version:v1.0"
    $allLines = Get-Content $csvFile
    if ($allLines[0] -match "version:v1.0") {
        $csvData = $allLines | Select-Object -Skip 1 | ConvertFrom-Csv -ErrorAction Stop
    } else {
        $csvData = Import-Csv $csvFile -ErrorAction Stop
    }
    
    # Map Azure AD CSV format to standard format and filter out empty rows
    $users = $csvData | Where-Object {
        # Only include rows with actual DisplayName data
        -not [string]::IsNullOrWhiteSpace($_.'Name [displayName] Required')
    } | ForEach-Object {
        # Extract mailNickname from UPN (before @)
        $mailNickname = if ($_.'User name [userPrincipalName] Required') {
            ($_.'User name [userPrincipalName] Required' -split '@')[0]
        } else { "" }
        
        [PSCustomObject]@{
            DisplayName = $_.'Name [displayName] Required'
            UserPrincipalName = $_.'User name [userPrincipalName] Required'
            MailNickname = $mailNickname
            Password = $_.'Initial password [passwordProfile] Required'
            AccountEnabled = if ($_.'Block sign in (Yes/No) [accountEnabled] Required' -eq 'Yes') { $false } else { $true }
            GivenName = $_.'First name [givenName]'
            Surname = $_.'Last name [surname]'
            JobTitle = $_.'Job title [jobTitle]'
            Department = $_.'Department [department]'
            UsageLocation = $_.'Usage location [usageLocation]'
            StreetAddress = $_.'Street address [streetAddress]'
            State = $_.'State or province [state]'
            Country = $_.'Country or region [country]'
            OfficeLocation = $_.'Office [physicalDeliveryOfficeName]'
            City = $_.'City [city]'
            PostalCode = $_.'ZIP or postal code [postalCode]'
            BusinessPhones = $_.'Office phone [telephoneNumber]'
            MobilePhone = $_.'Mobile phone [mobile]'
        }
    }
    
    Write-Host "Found $($users.Count) users in CSV file (empty rows filtered out)." -ForegroundColor Green
} catch {
    Write-Host "Error: Failed to read CSV file. $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# Display expected CSV format
Write-Host "`nExpected Azure AD CSV format with columns:" -ForegroundColor Yellow
Write-Host "  - Name [displayName] Required"
Write-Host "  - User name [userPrincipalName] Required"
Write-Host "  - Initial password [passwordProfile] Required"
Write-Host "  - Block sign in (Yes/No) [accountEnabled] Required"
Write-Host "  - First name [givenName] (optional)"
Write-Host "  - Last name [surname] (optional)"
Write-Host "  - Job title [jobTitle] (optional)"
Write-Host "  - Department [department] (optional)"
Write-Host "  - Usage location [usageLocation] (optional)"
Write-Host "  - Street address [streetAddress] (optional)"
Write-Host "  - State or province [state] (optional)"
Write-Host "  - Country or region [country] (optional)"
Write-Host "  - Office [physicalDeliveryOfficeName] (optional)"
Write-Host "  - City [city] (optional)"
Write-Host "  - ZIP or postal code [postalCode] (optional)"
Write-Host "  - Office phone [telephoneNumber] (optional)"
Write-Host "  - Mobile phone [mobile] (optional)"

# Confirm import with option to preview
Write-Host "`nDo you want to proceed with importing $($users.Count) users?" -ForegroundColor Yellow
Write-Host "[Y] Yes"
Write-Host "[P] Preview users first"
Write-Host "[N] No"
$confirm = Read-Host "Enter your choice"

if ($confirm -match "[pP]") {
    Write-Host "`n=== USERS TO BE IMPORTED ===" -ForegroundColor Cyan
    foreach ($user in $users) {
        Write-Host "  - $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor White
    }
    
    $confirm = Read-Host "`nDo you want to proceed with importing? [Y] Yes [N] No"
}

if ($confirm -notmatch "[yY]") {
    Write-Host "Import cancelled by user." -ForegroundColor Yellow
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Exit
}

# Import users
$successCount = 0
$failCount = 0
$failedUsers = @()
$currentCount = 0

Write-Host "`nStarting user import..." -ForegroundColor Cyan

foreach ($user in $users) {
    $currentCount++
    Write-Progress -Activity "Importing users" -Status "Processing $currentCount of $($users.Count)" -PercentComplete (($currentCount / $users.Count) * 100)
    
    # Validate required fields
    if ([string]::IsNullOrWhiteSpace($user.DisplayName) -or 
        [string]::IsNullOrWhiteSpace($user.UserPrincipalName) -or 
        [string]::IsNullOrWhiteSpace($user.MailNickname) -or
        [string]::IsNullOrWhiteSpace($user.Password)) {
        $failCount++
        $failedUsers += [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            Error = "Missing required fields (DisplayName, UserPrincipalName, MailNickname, or Password)"
        }
        continue
    }
    
    # Build user parameters
    $params = @{
        AccountEnabled = $user.AccountEnabled
        DisplayName = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        MailNickname = $user.MailNickname
        PasswordProfile = @{
            Password = $user.Password
            ForceChangePasswordNextSignIn = $true
        }
    }
    
    # Add optional fields if present
    if (-not [string]::IsNullOrWhiteSpace($user.GivenName)) { $params.GivenName = $user.GivenName }
    if (-not [string]::IsNullOrWhiteSpace($user.Surname)) { $params.Surname = $user.Surname }
    if (-not [string]::IsNullOrWhiteSpace($user.JobTitle)) { $params.JobTitle = $user.JobTitle }
    if (-not [string]::IsNullOrWhiteSpace($user.Department)) { $params.Department = $user.Department }
    if (-not [string]::IsNullOrWhiteSpace($user.UsageLocation)) { $params.UsageLocation = $user.UsageLocation }
    if (-not [string]::IsNullOrWhiteSpace($user.StreetAddress)) { $params.StreetAddress = $user.StreetAddress }
    if (-not [string]::IsNullOrWhiteSpace($user.State)) { $params.State = $user.State }
    if (-not [string]::IsNullOrWhiteSpace($user.Country)) { $params.Country = $user.Country }
    if (-not [string]::IsNullOrWhiteSpace($user.OfficeLocation)) { $params.OfficeLocation = $user.OfficeLocation }
    if (-not [string]::IsNullOrWhiteSpace($user.City)) { $params.City = $user.City }
    if (-not [string]::IsNullOrWhiteSpace($user.PostalCode)) { $params.PostalCode = $user.PostalCode }
    if (-not [string]::IsNullOrWhiteSpace($user.BusinessPhones)) { $params.BusinessPhones = @($user.BusinessPhones) }
    if (-not [string]::IsNullOrWhiteSpace($user.MobilePhone)) { $params.MobilePhone = $user.MobilePhone }
    
    # Create user
    try {
        New-MgUser -BodyParameter $params -ErrorAction Stop | Out-Null
        $successCount++
    } catch {
        $failCount++
        $failedUsers += [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            Error = $_.Exception.Message
        }
    }
}

Write-Progress -Activity "Importing users" -Completed

# Summary
Write-Host "`n=== USER IMPORT SUMMARY ===" -ForegroundColor Cyan
Write-Host "Total users processed: $($users.Count)" -ForegroundColor White
Write-Host "Successfully created: $successCount" -ForegroundColor Green
Write-Host "Failed: $failCount" -ForegroundColor Red

# List failed users
if ($failCount -gt 0) {
    Write-Host "`n=== FAILED USERS LIST ===" -ForegroundColor Yellow
    foreach ($failed in $failedUsers) {
        Write-Host "  - $($failed.UserPrincipalName)" -ForegroundColor Red
        Write-Host "    Reason: $($failed.Error)" -ForegroundColor DarkGray
    }
}
start-sleep -Seconds 2
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
