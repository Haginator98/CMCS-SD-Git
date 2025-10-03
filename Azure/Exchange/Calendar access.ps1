Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
Import-Module ExchangeOnlineManagement
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