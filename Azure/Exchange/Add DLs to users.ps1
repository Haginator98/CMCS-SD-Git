# Dette er en test fil, egenskrevet kode, blandet med noe AI 
# Koden her er ikke testet, og kan inneholde feil.

#Bruk denne dersom ExchangeOnlineManagement ikke er installert
Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser

#Dersom den er installert, bruk disse to linjene
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName epostadresse@dittdomene.no

# Define Users to receive access
$Users = @(
    "userA@domain.com",
    "userB@domain.com"
)
 
# Define Distribution lists (DLs) to grant access to
$DistributionLists = @(
    "DL@domain.com",
    "DL@domain.com"
)
 
# Loop through each combination of user and distribution list
foreach ($user in $Users) {
    foreach ($DL in $DistributionLists) {
        Write-Host "Granting $user access to $DL on"
        try {
            Add-DistributionGroupMember -Identity $DL -Member $user
            Write-Host "Successfully added $user to $DL"
        }
        catch {
            Write-Host -ForegroundColor Red "Error: Could not add $user to $DL"
            $error
            $error.clear
        }
    }
}