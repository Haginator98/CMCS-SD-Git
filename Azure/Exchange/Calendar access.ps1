Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
Import-Module ExchangeOnlineManagement

# Check if ExchangeOnlineManagement module is installed
$Module = Get-Module -Name ExchangeOnlineManagement -ListAvailable
if ($Module.Count -eq 0) {
    Write-Host ExchangeOnlineManagement module is not available -ForegroundColor yellow
    $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No"
    if ($Confirm -match "[yY]") {
        Install-Module ExchangeOnlineManagement
    } else {
        Write-Host ExchangeOnlineManagement module is required. Please install module using Install-Module ExchangeOnlineManagement cmdlet.
        Exit
    }
}
Write-Host Importing ExchangeOnlineManagement module... -ForegroundColor Yellow
Import-Module ExchangeOnlineManagement

#---------------------------------------------------------------------------#



Connect-ExchangeOnline 

# Define target users (calendar owners)
$targetUsers = @(
    "userA@domain.com",
    "userB@domain.com"
)
 
# Define delegates (users who need access)
$delegates = @(
    "delegate1@domain.com",
    "delegate2@domain.com"
)
 
# Loop through each combination and apply permission
foreach ($target in $targetUsers) {
    foreach ($delegate in $delegates) {
        $calendar = "${target}:\Calendar"
        Write-Host "Granting Editor access to $delegate on $calendar"
        Add-MailboxFolderPermission -Identity $calendar -User $delegate -AccessRights Editor
    }
}