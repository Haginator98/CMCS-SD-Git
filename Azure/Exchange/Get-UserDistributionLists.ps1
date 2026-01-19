# Check if ExchangeOnlineManagement module is installed
$Module = Get-Module -Name ExchangeOnlineManagement -ListAvailable
if ($Module.Count -eq 0) {
    Write-Host "ExchangeOnlineManagement module is not available" -ForegroundColor Yellow
    $Confirm = Read-Host "Do you want to install the module? [Y] Yes [N] No"
    if ($Confirm -match "[yY]") {
        Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Cyan
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
        Write-Host "Module installed successfully!" -ForegroundColor Green
    } else {
        Write-Host "ExchangeOnlineManagement module is required. Please install it using: Install-Module ExchangeOnlineManagement" -ForegroundColor Red
        Exit
    }
}

# Import module
Write-Host "Importing ExchangeOnlineManagement module..." -ForegroundColor Cyan
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online!" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
    Exit
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Get Distribution Lists for User" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Main loop
$continue = $true
while ($continue) {
    
    $UserEmail = Read-Host "Enter the user's email address"
    
    if ([string]::IsNullOrWhiteSpace($UserEmail)) {
        Write-Host "Email address cannot be empty." -ForegroundColor Red
        continue
    }
    
    # Try to get the user
    try {
        Write-Host "`nValidating user..." -ForegroundColor Yellow
        $userInfo = Get-Recipient -Identity $UserEmail -ErrorAction Stop
        
        Write-Host "User found: $($userInfo.DisplayName) ($($userInfo.PrimarySmtpAddress))" -ForegroundColor Green
        Write-Host "`nSearching for Distribution Lists... This may take a few minutes." -ForegroundColor Yellow
        
        # Initialize results array
        $results = @()
        
        # Get all distribution groups
        $allDLs = Get-DistributionGroup -ResultSize Unlimited
        $totalDLs = $allDLs.Count
        $counter = 0
        
        foreach ($dl in $allDLs) {
            $counter++
            Write-Progress -Activity "Checking Distribution Lists" -Status "Processing $($dl.DisplayName)" -PercentComplete (($counter / $totalDLs) * 100)
            
            # Check if user is a member
            $members = Get-DistributionGroupMember -Identity $dl.Identity -ResultSize Unlimited
            if ($members.PrimarySmtpAddress -contains $userInfo.PrimarySmtpAddress) {
                $results += [PSCustomObject]@{
                    DisplayName         = $dl.DisplayName
                    PrimarySmtpAddress  = $dl.PrimarySmtpAddress
                    Alias              = $dl.Alias
                    GroupType          = $dl.GroupType
                    RecipientTypeDetails = $dl.RecipientTypeDetails
                    ManagedBy          = ($dl.ManagedBy -join "; ")
                    HiddenFromAddressListsEnabled = $dl.HiddenFromAddressListsEnabled
                    RequireSenderAuthenticationEnabled = $dl.RequireSenderAuthenticationEnabled
                }
            }
        }
        
        Write-Progress -Activity "Checking Distribution Lists" -Completed
        
        if ($results.Count -gt 0) {
            # Prepare export
            $Desktop = [Environment]::GetFolderPath("Desktop")
            $Date = Get-Date -Format "yyyy-MM-dd_HHmm"
            $SafeAlias = $userInfo.Alias -replace '[^a-zA-Z0-9]', '_'
            $OutputFile = Join-Path $Desktop "UserDL_Memberships_${SafeAlias}_$Date.csv"
            
            # Export to CSV
            $results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
            
            # Display results
            Write-Host "`n========================================" -ForegroundColor Green
            Write-Host "RESULTS" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "Found $($results.Count) Distribution List(s) for $($userInfo.DisplayName):`n" -ForegroundColor Green
            
            $results | Select-Object DisplayName, PrimarySmtpAddress, GroupType | Format-Table -AutoSize
            
            Write-Host "Full results exported to:" -ForegroundColor Green
            Write-Host $OutputFile -ForegroundColor Cyan
        } else {
            Write-Host "`n========================================" -ForegroundColor Yellow
            Write-Host "No Distribution Lists found for user: $($userInfo.DisplayName)" -ForegroundColor Yellow
            Write-Host "========================================" -ForegroundColor Yellow
        }
        
        # Ask if user wants to check another email
        Write-Host "`n"
        $retry = Read-Host "Do you want to check another user? [Y] Yes [N] No"
        if ($retry -notmatch "[yY]") {
            $continue = $false
        }
        
    } catch {
        Write-Host "`n========================================" -ForegroundColor Red
        Write-Host "ERROR" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "Could not find user with email address: $UserEmail" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "========================================`n" -ForegroundColor Red
        
        $retry = Read-Host "Would you like to try another email address? [Y] Yes [N] No"
        if ($retry -notmatch "[yY]") {
            $continue = $false
        }
    }
}

# Disconnect from Exchange Online
Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false

Write-Host "`nScript completed successfully!" -ForegroundColor Green
