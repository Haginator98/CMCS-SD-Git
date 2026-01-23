#This script is used to choose what script to run in Azure AD / Exchange environment
#Made by Mr. Hagen - 2025

# Organize scripts by category
$scriptCategories = @{
    "Entra" = @(
        @{ Name = "Change Department"; Path = "$PSScriptRoot\Entra\Change department.ps1" }
        @{ Name = "Change Street Address"; Path = "$PSScriptRoot\Entra\change street address Entra.ps1" }
        @{ Name = "Change Dep. Office based on Street Address"; Path = "$PSScriptRoot\Entra\change dep.office based on street address.ps1" }
        @{ Name = "Bulk import CSV"; Path = "$PSScriptRoot\Entra\Bulk import users.ps1" }
        @{ Name = "Update users from CSV"; Path = "$PSScriptRoot\Entra\Update users from CSV.ps1" }
        @{ Name = "Export Filtered Entra Users"; Path = "$PSScriptRoot\Entra\Export-FilteredEntraUsers.ps1" }
        @{ Name = "Set Manager by Street Address"; Path = "$PSScriptRoot\Entra\Set-ManagerByStreetAddress.ps1" }
    )
    "Exchange" = @(
        @{ Name = "Find Alias Mailbox"; Path = "$PSScriptRoot\Exchange\Find-Aliasmailbox.ps1" }
        @{ Name = "Get Shared Mailboxes from User"; Path = "$PSScriptRoot\Exchange\Get shared mailboxes from user.ps1" }
        @{ Name = "Export All DL Members"; Path = "$PSScriptRoot\Exchange\Export all DL members.ps1" }
        @{ Name = "Import DL Members from CSV"; Path = "$PSScriptRoot\Exchange\Import DL members.ps1" }
        @{ Name = "Import DL Members from CSV (Alternative)"; Path = "$PSScriptRoot\Exchange\Import-DLMembersFromCSV.ps1" }
        @{ Name = "Import Contacts from CSV"; Path = "$PSScriptRoot\Exchange\Import-ContactsFromCSV.ps1" }
        @{ Name = "Get User Distribution Lists"; Path = "$PSScriptRoot\Exchange\Get-UserDistributionLists.ps1" }
    )
    "Licenses" = @(
        @{ Name = "Get Dynamics Licenses for Users"; Path = "$PSScriptRoot\Licenses\Get-DynamicsLicensesForUsers.ps1" }
        @{ Name = "Remove Direct User License"; Path = "$PSScriptRoot\Licenses\Remove-DirectUserLicense.ps1" }
    )
#    "Teams" = @(
#        @{ Name = "Teams Reports"; Path = "$PSScriptRoot\Teams\TeamsReports.ps1" }
#    )
}

while ($true) {
    Clear-Host
    Write-Host "
░██████╗███████╗██████╗░██╗░░░██╗██╗░█████╗░███████╗██████╗░███████╗░██████╗██╗░░██╗
██╔════╝██╔════╝██╔══██╗██║░░░██║██║██╔══██╗██╔════╝██╔══██╗██╔════╝██╔════╝██║░██╔╝
╚█████╗░█████╗░░██████╔╝╚██╗░██╔╝██║██║░░╚═╝█████╗░░██║░░██║█████╗░░╚█████╗░█████═╝░
░╚═══██╗██╔══╝░░██╔══██╗░╚████╔╝░██║██║░░██╗██╔══╝░░██║░░██║██╔══╝░░░╚═══██╗██╔═██╗░
██████╔╝███████╗██║░░██║░░╚██╔╝░░██║╚█████╔╝███████╗██████╔╝███████╗██████╔╝██║░╚██╗
╚═════╝░╚══════╝╚═╝░░╚═╝░░░╚═╝░░░╚═╝░╚════╝░╚══════╝╚═════╝░╚══════╝╚═════╝░╚═╝░░╚═╝" -ForegroundColor Cyan
    Write-Host "Welcome to Servicedesk Tools! - Made by Mr.Hagen" -ForegroundColor Green 
    Write-Host "Following scripts have been tested, let me know if there is a bug, issue with something or if you got any ideas for scripts to add." -ForegroundColor Yellow
    Write-Host "Remember that you need PIM activated" -ForegroundColor Red
    Write-Host ""
    Write-Host "Select a category:" -ForegroundColor Yellow
    Write-Host "1: Entra ID"
    Write-Host "2: Exchange"
    Write-Host "3: Licenses"
#    Write-Host "4: Teams"
    Write-Host "0: Exit"

    $categoryChoice = Read-Host "Enter the number of the category"
    
    if ($categoryChoice -eq '0') { break }
    if ($categoryChoice -eq 'exit') { break }

    $selectedCategory = switch ($categoryChoice) {
        '1' { "Entra" }
        '2' { "Exchange" }
        '3' { "Licenses" }
    #    '4' { "Teams" }
        default { $null }
    }

    if ($selectedCategory -and $scriptCategories.ContainsKey($selectedCategory)) {
        $scripts = $scriptCategories[$selectedCategory]
        
        while ($true) {
            Clear-Host
            Write-Host "=== $selectedCategory Scripts ===" -ForegroundColor Cyan
            Write-Host ""
            for ($i = 0; $i -lt $scripts.Count; $i++) {
                Write-Host "$($i+1): $($scripts[$i].Name)"
            }
            Write-Host "0: Back to main menu"

            $choice = Read-Host "Enter the number of the script to run"
            
            if ($choice -eq '0') { break }
            if ($choice -eq 'back') { break }
            
            if ($choice -match '^[1-9][0-9]*$' -and $choice -le $scripts.Count) {
                $scriptToRun = $scripts[$choice-1].Path
                Write-Host "Running $($scripts[$choice-1].Name)..." -ForegroundColor Cyan
                & $scriptToRun
                Write-Host "`nScript finished. Press any key to return to menu..." -ForegroundColor Yellow
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            } else {
                Write-Host "Invalid selection. Try again." -ForegroundColor Red
                Start-Sleep -Seconds 2
            }
        }
    } else {
        Write-Host "Invalid selection. Try again." -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

Write-Host "Exiting Servicedesk Tools. Goodbye!" -ForegroundColor Green
Start-Sleep -Seconds 2
Clear-Host
