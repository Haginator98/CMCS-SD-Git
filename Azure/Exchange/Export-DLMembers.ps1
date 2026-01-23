# Import required modules (installed via Tools.ps1)
Import-Module ExchangeOnlineManagement -ErrorAction Stop

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline | Out-Null
Write-Host "This script will export all members of a Distribution List (DL) to a CSV file on your desktop." -ForegroundColor Cyan

# Loop
$success = $false
while (-not $success) {
    
    $DL = Read-Host "Enter the Distribution List email address"
    
    # Try to get the distribution group
    try {
        $dlInfo = Get-DistributionGroup -Identity $DL -ErrorAction Stop
        
        $Desktop = [Environment]::GetFolderPath("Desktop")
        $Date = Get-Date -Format "yyyy-MM-dd"
        $MembersFile = Join-Path $Desktop "$($dlInfo.DisplayName)_Members_$Date.csv"
        $OwnersFile = Join-Path $Desktop "$($dlInfo.DisplayName)_Owners_$Date.csv"

        # Get and Export direct members
        Get-DistributionGroupMember -Identity $DL -ResultSize Unlimited |
            Select-Object PrimarySmtpAddress |
            Export-Csv -Path $MembersFile -NoTypeInformation -Encoding UTF8

        # Get and Export owners
        $owners = $dlInfo.ManagedBy
        if ($owners.Count -gt 0) {
            $ownerDetails = $owners | ForEach-Object {
                try {
                    Get-Recipient -Identity $_ -ErrorAction Stop | Select-Object PrimarySmtpAddress
                } catch {
                    Write-Warning "Could not retrieve details for owner: $_"
                }
            }
            $ownerDetails | Export-Csv -Path $OwnersFile -NoTypeInformation -Encoding UTF8
            Write-Host "Members exported to: $MembersFile" -ForegroundColor Green
            Write-Host "Owners exported to: $OwnersFile" -ForegroundColor Green
        } else {
            Write-Host "Members exported to: $MembersFile" -ForegroundColor Green
            Write-Host "No owners found for this DL." -ForegroundColor Yellow
        }
        
        Start-Sleep -Seconds 2
        $success = $true
        
    } catch {
        Write-Host "Error: Distribution List '$DL' not found." -ForegroundColor Red
        $retry = Read-Host "Do you want to try again? [Y] Yes [N] No"
        
        if ($retry -notmatch "[yY]") {
            Write-Host "Exiting script..." -ForegroundColor Yellow
            break
        }
    }
}

# Disconnect from Exchange Online
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected successfully." -ForegroundColor Green
Start-Sleep -Seconds 2
