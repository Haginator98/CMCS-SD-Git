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
# Connect to Microsoft Graph
Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -NoWelcome

# Export all users with their managers to a CSV file
Write-Host "This script will export all users along with their managers from Entra ID (Azure AD) to a CSV file." -ForegroundColor Cyan
Start-Sleep -Seconds 1
$output = @()
$users = Get-MgUser -All -Property DisplayName, UserPrincipalName, Id

$total = $users.Count
$counter = 0

foreach ($user in $users) {
    $counter++
    Write-Host "Processing $($user.DisplayName) ($counter of $total)" -ForegroundColor Cyan
    Write-Progress -Activity "Exporting users and managers" -Status "$counter of $total users" -PercentComplete (($counter / $total) * 100)

    $manager = Get-MgUserManager -UserId $user.Id -ErrorAction SilentlyContinue
    $managerDisplayName = if ($manager) { $manager.DisplayName } else { "N/A" }
    $managerEmail = if ($manager) { $manager.UserPrincipalName } else { "N/A" }
    
    $output += [PSCustomObject]@{
        DisplayName        = $user.DisplayName
        UserPrincipalName  = $user.UserPrincipalName
        ManagerDisplayName = $managerDisplayName
        ManagerEmail       = $managerEmail
    }
}

# Get a cross-platform Desktop path
$desktopPath = [Environment]::GetFolderPath("Desktop")
$exportPath = Join-Path -Path $desktopPath -ChildPath "EntraUsersWithManagers.csv"

$output | Export-Csv -Path $exportPath -NoTypeInformation
Write-Host "Export completed. The file is saved at $exportPath" -ForegroundColor Green
Start-Sleep -Seconds 4

# Disconnect from Microsoft Graph
Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected. Script finished." -ForegroundColor Green
Start-Sleep -Seconds 2
