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
        Write-Host "Microsoft.Graph module is required. Please install it using Install-Module Microsoft.Graph cmdlet."
        Exit
    }
}

# Check if ExchangeOnlineManagement module is installed
$ExoModule = Get-Module -Name ExchangeOnlineManagement -ListAvailable
if ($ExoModule.Count -eq 0) {
    Write-Host "ExchangeOnlineManagement module is not available." -ForegroundColor Yellow
    $Confirm = Read-Host "This module is needed for accurate shared mailbox filtering. Install it? [Y] Yes [N] No"
    if ($Confirm -match "[yY]") {
        Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Cyan
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
        if ($?) {
            Write-Host "ExchangeOnlineManagement module installed successfully." -ForegroundColor Green
        } else {
            Write-Host "Failed to install ExchangeOnlineManagement module." -ForegroundColor Red
            Exit
        }
    } else {
        Write-Host "ExchangeOnlineManagement module is required for accurate filtering. Exiting..." -ForegroundColor Yellow
        Exit
    }
}

# Connect to Microsoft Graph
Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome

# Connect to Exchange Online
Write-Host "Signing in to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowBanner:$false

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

# Ask about data scope
Write-Host "`nStandard fields that will be exported:" -ForegroundColor Yellow
Write-Host "  - Display Name" -ForegroundColor Gray
Write-Host "  - UPN" -ForegroundColor Gray
Write-Host "  - Account Enabled" -ForegroundColor Gray
Write-Host "  - User Type" -ForegroundColor Gray
Write-Host "  - On-premises synced" -ForegroundColor Gray
Write-Host "  - Company Name" -ForegroundColor Gray
Write-Host "  - Department" -ForegroundColor Gray
Write-Host "  - Office Location" -ForegroundColor Gray
Write-Host "  - Job Title" -ForegroundColor Gray
Write-Host "  - Phone number" -ForegroundColor Gray
Write-Host "  - Manager`n" -ForegroundColor Gray

$gatherAllInfo = Read-Host "Type 'gather all info' to export all user properties, or press Enter for standard fields"
$exportAllProperties = $gatherAllInfo -eq "gather all info"

if ($exportAllProperties) {
    Write-Host "Will export ALL user properties" -ForegroundColor Green
} else {
    Write-Host "Will export standard fields only" -ForegroundColor Green
}

Start-Sleep -Seconds 1

# Retrieve users with all necessary properties
Write-Host "`nRetrieving users from Entra ID..." -ForegroundColor Cyan

if ($exportAllProperties) {
    $users = Get-MgUser -All
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

if ($filterSuffix -ne "") {
    $cleanFilter = $filterSuffix -replace '[^a-zA-Z0-9]', ''
    if ($exportAllProperties) {
        $exportPath = Join-Path -Path $desktopPath -ChildPath "EntraUsers_AllInfo_Filter_$cleanFilter`_$timestamp.csv"
    } else {
        $exportPath = Join-Path -Path $desktopPath -ChildPath "EntraUsers_StandardFields_Filter_$cleanFilter`_$timestamp.csv"
    }
} else {
    if ($exportAllProperties) {
        $exportPath = Join-Path -Path $desktopPath -ChildPath "EntraUsers_AllInfo_$timestamp.csv"
    } else {
        $exportPath = Join-Path -Path $desktopPath -ChildPath "EntraUsers_StandardFields_$timestamp.csv"
    }
}

# Export to CSV with BOM to ensure proper display of Norwegian characters (æ, ø, å) in Excel
$output | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8BOM
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
