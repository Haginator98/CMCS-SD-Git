# NOTE: This script requires both Graph and Exchange modules
# Import required modules
Write-Host "Loading required modules..." -ForegroundColor Cyan

# Import Microsoft Graph modules
try {
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Write-Host "✓ Microsoft Graph modules loaded" -ForegroundColor Green
} catch {
    Write-Host "Failed to import Microsoft Graph modules: $_" -ForegroundColor Red
    Write-Host "Install with: Install-Module Microsoft.Graph -Scope CurrentUser -Force" -ForegroundColor Yellow
    Exit
}

# Import Exchange Online module
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "✓ Exchange Online module loaded" -ForegroundColor Green
} catch {
    Write-Host "Failed to import Exchange Online module: $_" -ForegroundColor Red
    Write-Host "Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force" -ForegroundColor Yellow
    Exit
}

# Connect to Microsoft Graph FIRST
Write-Host "`nSigning in to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "✓ Successfully connected to Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    Write-Host "`nTry updating your modules with:" -ForegroundColor Yellow
    Write-Host "  Update-Module Microsoft.Graph -Force" -ForegroundColor Cyan
    Write-Host "  Update-Module ExchangeOnlineManagement -Force" -ForegroundColor Cyan
    Exit
}

# Connect to Exchange Online (after Graph to avoid conflicts)
Write-Host "`nSigning in to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "✓ Successfully connected to Exchange Online" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Exchange Online: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    Exit
}

# Export all users with their managers to a CSV file
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "User Export Script - Entra ID" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Ask if user wants to filter by UPN domain
Write-Host "Do you want to filter users by UPN domain?" -ForegroundColor Yellow
Write-Host "Examples: .dk, .no, .com, @test.pl, @company.com" -ForegroundColor Gray
$filterChoice = Read-Host "Enter domain filter (or press Enter to export all users)"

$filterSuffix = ""
if ($filterChoice -ne "") {
    $filterSuffix = $filterChoice
    Write-Host "Filtering users with UPN ending in: $filterSuffix" -ForegroundColor Green
} else {
    Write-Host "No filter applied - exporting all users" -ForegroundColor Green
}

# Ask if user wants to include shared mailboxes and distribution lists
Write-Host "`nDo you want to include shared mailboxes and distribution lists?" -ForegroundColor Yellow
$includeSharedAndDL = Read-Host "[Y] Yes [N] No (default: No)"
$shouldIncludeSharedAndDL = $includeSharedAndDL -match "[yY]"

if ($shouldIncludeSharedAndDL) {
    Write-Host "Will include shared mailboxes and distribution lists" -ForegroundColor Green
} else {
    Write-Host "Will exclude shared mailboxes and distribution lists (standard)" -ForegroundColor Green
}

# Ask if user wants to exclude external users
Write-Host "`nDo you want to exclude external/guest users?" -ForegroundColor Yellow
$excludeExternalUsers = Read-Host "[Y] Yes [N] No (default: No)"
$shouldExcludeExternal = $excludeExternalUsers -match "[yY]"

if ($shouldExcludeExternal) {
    Write-Host "Will exclude external/guest users" -ForegroundColor Green
} else {
    Write-Host "Will include external/guest users" -ForegroundColor Green
}

# Ask about data scope
Write-Host "`nExport field options:" -ForegroundColor Yellow
Write-Host "[1] Standard fields (recommended)" -ForegroundColor Cyan
Write-Host "    - Display Name, UPN, Account Enabled, User Type, On-premises synced," -ForegroundColor Gray
Write-Host "      Company Name, Department, Office Location, Job Title, Phone number, Manager" -ForegroundColor Gray
Write-Host "[2] Custom fields (select specific fields)" -ForegroundColor Cyan
Write-Host "[3] All fields (complete user object)`n" -ForegroundColor Cyan

$fieldChoice = Read-Host "Enter your choice (1/2/3)"

$exportAllProperties = $false
$customFields = @()

switch ($fieldChoice) {
    '1' {
        Write-Host "Will export standard fields" -ForegroundColor Green
    }
    '2' {
        Write-Host "`nAvailable fields:" -ForegroundColor Yellow
        Write-Host "  1. DisplayName" -ForegroundColor Gray
        Write-Host "  2. UserPrincipalName" -ForegroundColor Gray
        Write-Host "  3. Object ID (Azure AD Id)" -ForegroundColor Gray
        Write-Host "  4. AccountEnabled" -ForegroundColor Gray
        Write-Host "  5. UserType" -ForegroundColor Gray
        Write-Host "  6. OnPremisesSyncEnabled" -ForegroundColor Gray
        Write-Host "  7. CompanyName" -ForegroundColor Gray
        Write-Host "  8. Department" -ForegroundColor Gray
        Write-Host "  9. OfficeLocation" -ForegroundColor Gray
        Write-Host " 10. JobTitle" -ForegroundColor Gray
        Write-Host " 11. BusinessPhones / MobilePhone" -ForegroundColor Gray
        Write-Host " 12. Manager" -ForegroundColor Gray
        Write-Host " 13. Mail" -ForegroundColor Gray
        Write-Host " 14. EmployeeId" -ForegroundColor Gray
        Write-Host " 15. Country" -ForegroundColor Gray
        Write-Host " 16. City" -ForegroundColor Gray
        Write-Host " 17. StreetAddress" -ForegroundColor Gray
        Write-Host " 18. PostalCode" -ForegroundColor Gray
        Write-Host " 19. PreferredLanguage" -ForegroundColor Gray
        Write-Host " 20. CreatedDateTime" -ForegroundColor Gray
        Write-Host " 21. LastPasswordChangeDateTime" -ForegroundColor Gray
        Write-Host " 22. AssignedLicenses" -ForegroundColor Gray
        Write-Host " 23. ProxyAddresses" -ForegroundColor Gray
        Write-Host " 24. MailNickname (Exchange Alias)" -ForegroundColor Gray
        
        Write-Host "`nEnter field numbers separated by comma (e.g., 1,2,3,7,11)" -ForegroundColor Yellow
        Write-Host "Or type 'all' to include all fields above" -ForegroundColor Gray
        $fieldSelection = Read-Host "Your selection"
        
        if ($fieldSelection -eq 'all') {
            $customFields = @('DisplayName', 'UserPrincipalName', 'Id', 'AccountEnabled', 'UserType', 'OnPremisesSyncEnabled', 
                            'CompanyName', 'Department', 'OfficeLocation', 'JobTitle', 'BusinessPhones', 'MobilePhone', 
                            'Manager', 'Mail', 'EmployeeId', 'Country', 'City', 'StreetAddress', 'PostalCode', 
                            'PreferredLanguage', 'CreatedDateTime', 'LastPasswordChangeDateTime', 'AssignedLicenses', 'ProxyAddresses', 'MailNickname')
        } else {
            $selectedNumbers = $fieldSelection -split ',' | ForEach-Object { $_.Trim() }
            $fieldMap = @{
                '1' = 'DisplayName'; '2' = 'UserPrincipalName'; '3' = 'Id'; '4' = 'AccountEnabled'; '5' = 'UserType'
                '6' = 'OnPremisesSyncEnabled'; '7' = 'CompanyName'; '8' = 'Department'; '9' = 'OfficeLocation'
                '10' = 'JobTitle'; '11' = 'Phone'; '12' = 'Manager'; '13' = 'Mail'; '14' = 'EmployeeId'
                '15' = 'Country'; '16' = 'City'; '17' = 'StreetAddress'; '18' = 'PostalCode'
                '19' = 'PreferredLanguage'; '20' = 'CreatedDateTime'; '21' = 'LastPasswordChangeDateTime'
                '22' = 'AssignedLicenses'; '23' = 'ProxyAddresses'; '24' = 'MailNickname'
            }
            foreach ($num in $selectedNumbers) {
                if ($fieldMap.ContainsKey($num)) {
                    $customFields += $fieldMap[$num]
                }
            }
        }
        Write-Host "Will export selected custom fields: $($customFields -join ', ')" -ForegroundColor Green
    }
    '3' {
        $exportAllProperties = $true
        Write-Host "Will export ALL user properties" -ForegroundColor Green
    }
    default {
        Write-Host "Invalid choice. Using standard fields." -ForegroundColor Yellow
    }
}

Start-Sleep -Seconds 1

# Retrieve users with all necessary properties
Write-Host "`nRetrieving users from Entra ID..." -ForegroundColor Cyan

if ($exportAllProperties) {
    $users = Get-MgUser -All
} elseif ($customFields.Count -gt 0) {
    # Build property list for custom fields
    $propertiesToFetch = @('DisplayName', 'UserPrincipalName', 'Id', 'UserType')
    $additionalProps = @('CompanyName', 'Department', 'OfficeLocation', 'JobTitle', 'BusinessPhones', 'MobilePhone', 
                        'AccountEnabled', 'Mail', 'MailNickname', 'OnPremisesSyncEnabled', 'EmployeeId', 
                        'Country', 'City', 'StreetAddress', 'PostalCode', 'PreferredLanguage', 'CreatedDateTime', 
                        'LastPasswordChangeDateTime', 'AssignedLicenses', 'ProxyAddresses')
    foreach ($prop in $additionalProps) {
        if ($prop -notin $propertiesToFetch) {
            $propertiesToFetch += $prop
        }
    }
    $users = Get-MgUser -All -Property $propertiesToFetch
} else {
    $users = Get-MgUser -All -Property DisplayName, UserPrincipalName, Id, UserType, OnPremisesSyncEnabled, CompanyName, Department, OfficeLocation, JobTitle, BusinessPhones, MobilePhone, AccountEnabled, Mail, MailNickname
}

# Filter users if a domain filter was specified
if ($filterSuffix -ne "") {
    $users = $users | Where-Object { $_.UserPrincipalName -like "*$filterSuffix" }
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "FILTER RESULT: Found $($users.Count) objects matching '$filterSuffix'" -ForegroundColor Yellow
    Write-Host "========================================`n" -ForegroundColor Yellow
} else {
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "TOTAL OBJECTS: Found $($users.Count) objects (no filter applied)" -ForegroundColor Yellow
    Write-Host "========================================`n" -ForegroundColor Yellow
}

# Filter out external/guest users if requested
if ($shouldExcludeExternal) {
    $beforeCount = $users.Count
    $users = $users | Where-Object { $_.UserType -ne 'Guest' }
    $removedCount = $beforeCount - $users.Count
    Write-Host "Removed $removedCount external/guest users" -ForegroundColor Yellow
}

# Filter out shared mailboxes and distribution lists by default (unless user chose to include them)
if (-not $shouldIncludeSharedAndDL) {
    Write-Host "Filtering out shared mailboxes and distribution lists using Exchange Online..." -ForegroundColor Cyan
    $beforeCount = $users.Count
    
    # Get all recipients from Exchange Online to check RecipientTypeDetails
    Write-Host "Retrieving recipient information from Exchange Online (this may take a moment)..." -ForegroundColor Gray
    $allRecipients = Get-EXORecipient -ResultSize Unlimited -Properties RecipientTypeDetails, PrimarySmtpAddress
    
    # Create a hashtable for fast lookup of recipient types
    $recipientTypes = @{}
    foreach ($recipient in $allRecipients) {
        if ($recipient.PrimarySmtpAddress) {
            $recipientTypes[$recipient.PrimarySmtpAddress.ToLower()] = $recipient.RecipientTypeDetails
        }
    }
    
    # Filter out shared mailboxes and distribution lists
    $users = $users | Where-Object {
        $upn = $_.UserPrincipalName.ToLower()
        $recipientType = $recipientTypes[$upn]
        
        # Keep only if it's a UserMailbox or MailUser (regular users)
        # Exclude: SharedMailbox, RoomMailbox, EquipmentMailbox, MailUniversalDistributionGroup, etc.
        $recipientType -in @('UserMailbox', 'MailUser', 'GuestMailUser') -or $null -eq $recipientType
    }
    
    $removedCount = $beforeCount - $users.Count
    Write-Host "Removed $removedCount shared mailboxes and distribution lists" -ForegroundColor Yellow
}

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "FINAL COUNT: $($users.Count) users will be processed" -ForegroundColor Yellow
Write-Host "========================================`n" -ForegroundColor Yellow

if ($users.Count -eq 0) {
    Write-Host "No users found. Exiting script." -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Exit
}

# Ask for confirmation before processing
$confirmation = Read-Host "Do you want to continue processing these $($users.Count) users? [Y] Yes [N] No"
if ($confirmation -notmatch "[yY]") {
    Write-Host "Export cancelled by user." -ForegroundColor Yellow
    Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    Exit
}

$output = @()
$total = $users.Count
$counter = 0

foreach ($user in $users) {
    $counter++
    Write-Progress -Activity "Processing users" -Status "Processing $($user.DisplayName) ($counter of $total)" -PercentComplete (($counter / $total) * 100)

    # Get manager information
    $manager = Get-MgUserManager -UserId $user.Id -ErrorAction SilentlyContinue
    if ($manager) {
        # Try different properties to get manager name
        $managerName = if ($manager.AdditionalProperties.displayName) { 
            $manager.AdditionalProperties.displayName 
        } elseif ($manager.DisplayName) { 
            $manager.DisplayName 
        } else { 
            "no manager" 
        }
    } else {
        $managerName = "no manager"
    }

    if ($exportAllProperties) {
        # Export all properties
        $userObject = [PSCustomObject]@{
            DisplayName           = $user.DisplayName
            UserPrincipalName     = $user.UserPrincipalName
            Manager               = $managerName
        }
        
        # Add all other properties dynamically
        $user.PSObject.Properties | Where-Object { $_.Name -notin @('DisplayName', 'UserPrincipalName') } | ForEach-Object {
            $userObject | Add-Member -MemberType NoteProperty -Name $_.Name -Value $_.Value -Force
        }
        
        $output += $userObject
    } elseif ($customFields.Count -gt 0) {
        # Export custom selected fields
        $userObject = [PSCustomObject]@{}
        
        foreach ($field in $customFields) {
            switch ($field) {
                'DisplayName' { $userObject | Add-Member -NotePropertyName 'DisplayName' -NotePropertyValue ($user.DisplayName ?? "no display name") }
                'UserPrincipalName' { $userObject | Add-Member -NotePropertyName 'UPN' -NotePropertyValue ($user.UserPrincipalName ?? "no UPN") }
                'Id' { $userObject | Add-Member -NotePropertyName 'ObjectID' -NotePropertyValue ($user.Id ?? "no object id") }
                'AccountEnabled' { $userObject | Add-Member -NotePropertyName 'AccountEnabled' -NotePropertyValue ($user.AccountEnabled ?? "no account status") }
                'UserType' { $userObject | Add-Member -NotePropertyName 'UserType' -NotePropertyValue ($user.UserType ?? "no user type") }
                'OnPremisesSyncEnabled' { $userObject | Add-Member -NotePropertyName 'OnPremisesSynced' -NotePropertyValue ($user.OnPremisesSyncEnabled ?? "no sync info") }
                'CompanyName' { $userObject | Add-Member -NotePropertyName 'CompanyName' -NotePropertyValue ($user.CompanyName ?? "no company name") }
                'Department' { $userObject | Add-Member -NotePropertyName 'Department' -NotePropertyValue ($user.Department ?? "no department") }
                'OfficeLocation' { $userObject | Add-Member -NotePropertyName 'OfficeLocation' -NotePropertyValue ($user.OfficeLocation ?? "no office location") }
                'JobTitle' { $userObject | Add-Member -NotePropertyName 'JobTitle' -NotePropertyValue ($user.JobTitle ?? "no job title") }
                'Phone' { 
                    $phoneNumber = if ($user.BusinessPhones -and $user.BusinessPhones.Count -gt 0) { $user.BusinessPhones[0] } elseif ($user.MobilePhone) { $user.MobilePhone } else { "no phone number" }
                    $userObject | Add-Member -NotePropertyName 'PhoneNumber' -NotePropertyValue $phoneNumber
                }
                'Manager' { $userObject | Add-Member -NotePropertyName 'Manager' -NotePropertyValue $managerName }
                'Mail' { $userObject | Add-Member -NotePropertyName 'Mail' -NotePropertyValue ($user.Mail ?? "no mail") }
                'EmployeeId' { $userObject | Add-Member -NotePropertyName 'EmployeeId' -NotePropertyValue ($user.EmployeeId ?? "no employee id") }
                'Country' { $userObject | Add-Member -NotePropertyName 'Country' -NotePropertyValue ($user.Country ?? "no country") }
                'City' { $userObject | Add-Member -NotePropertyName 'City' -NotePropertyValue ($user.City ?? "no city") }
                'StreetAddress' { $userObject | Add-Member -NotePropertyName 'StreetAddress' -NotePropertyValue ($user.StreetAddress ?? "no street address") }
                'PostalCode' { $userObject | Add-Member -NotePropertyName 'PostalCode' -NotePropertyValue ($user.PostalCode ?? "no postal code") }
                'PreferredLanguage' { $userObject | Add-Member -NotePropertyName 'PreferredLanguage' -NotePropertyValue ($user.PreferredLanguage ?? "no language") }
                'CreatedDateTime' { $userObject | Add-Member -NotePropertyName 'CreatedDateTime' -NotePropertyValue ($user.CreatedDateTime ?? "no creation date") }
                'LastPasswordChangeDateTime' { $userObject | Add-Member -NotePropertyName 'LastPasswordChangeDateTime' -NotePropertyValue ($user.LastPasswordChangeDateTime ?? "no password change date") }
                'AssignedLicenses' { 
                    $licenses = if ($user.AssignedLicenses.Count -gt 0) { ($user.AssignedLicenses | ForEach-Object { $_.SkuId }) -join '; ' } else { "no licenses" }
                    $userObject | Add-Member -NotePropertyName 'AssignedLicenses' -NotePropertyValue $licenses
                }
                'ProxyAddresses' { 
                    $proxies = if ($user.ProxyAddresses.Count -gt 0) { $user.ProxyAddresses -join '; ' } else { "no proxy addresses" }
                    $userObject | Add-Member -NotePropertyName 'ProxyAddresses' -NotePropertyValue $proxies
                }
                'MailNickname' { $userObject | Add-Member -NotePropertyName 'MailNickname' -NotePropertyValue ($user.MailNickname ?? "no mail nickname") }
            }
        }
        
        $output += $userObject
    } else {
        # Export standard fields only
        $phoneNumber = if ($user.BusinessPhones -and $user.BusinessPhones.Count -gt 0) { 
            $user.BusinessPhones[0] 
        } elseif ($user.MobilePhone) { 
            $user.MobilePhone 
        } else { 
            "no phone number" 
        }
        
        $output += [PSCustomObject]@{
            DisplayName         = if ($user.DisplayName) { $user.DisplayName } else { "no display name" }
            UPN                 = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { "no UPN" }
            AccountEnabled      = if ($null -ne $user.AccountEnabled) { $user.AccountEnabled } else { "no account status" }
            UserType            = if ($user.UserType) { $user.UserType } else { "no user type" }
            OnPremisesSynced    = if ($null -ne $user.OnPremisesSyncEnabled) { $user.OnPremisesSyncEnabled } else { "no sync info" }
            CompanyName         = if ($user.CompanyName) { $user.CompanyName } else { "no company name" }
            Department          = if ($user.Department) { $user.Department } else { "no department" }
            OfficeLocation      = if ($user.OfficeLocation) { $user.OfficeLocation } else { "no office location" }
            JobTitle            = if ($user.JobTitle) { $user.JobTitle } else { "no job title" }
            PhoneNumber         = $phoneNumber
            Manager             = $managerName
        }
    }
}

Write-Progress -Activity "Processing users" -Completed

# Create filename based on filter
$desktopPath = [Environment]::GetFolderPath("Desktop")
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$exportType = if ($exportAllProperties) { 
    "AllInfo" 
} elseif ($customFields.Count -gt 0) { 
    "CustomFields" 
} else { 
    "StandardFields" 
}

if ($filterSuffix -ne "") {
    $cleanFilter = $filterSuffix -replace '[^a-zA-Z0-9]', ''
    $exportPath = Join-Path -Path $desktopPath -ChildPath "EntraUsers_$exportType`_Filter_$cleanFilter`_$timestamp.csv"
} else {
    $exportPath = Join-Path -Path $desktopPath -ChildPath "EntraUsers_$exportType`_$timestamp.csv"
}

# Export to CSV with UTF8 encoding to ensure proper display of Norwegian characters (æ, ø, å) in Excel
$output | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
Write-Host "`n========================================" -ForegroundColor Green
Write-Host "Export completed successfully!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "File location: $exportPath" -ForegroundColor Cyan
Write-Host "Total users exported: $($output.Count)" -ForegroundColor Cyan
Start-Sleep -Seconds 3

# Disconnect from Microsoft Graph and Exchange Online
Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Write-Host "Disconnected. Script finished." -ForegroundColor Green
Start-Sleep -Seconds 2
