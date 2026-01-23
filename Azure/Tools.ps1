#This script is used to choose what script to run in Azure AD / Exchange environment
#Made by Mr. Hagen - 2025

# Organize scripts by category
$scriptCategories = @{
    "Entra" = @(
        @{ Name = "Set Department"; Path = "$PSScriptRoot\Entra\Set-Department.ps1" }
        @{ Name = "Set Street Address"; Path = "$PSScriptRoot\Entra\Set-StreetAddress.ps1" }
        @{ Name = "Set Department Office by Street Address"; Path = "$PSScriptRoot\Entra\Set-DepartmentOfficeByStreetAddress.ps1" }
        @{ Name = "Import Entra Users"; Path = "$PSScriptRoot\Entra\Import-EntraUsers.ps1" }
        @{ Name = "Update Entra Users from CSV"; Path = "$PSScriptRoot\Entra\Update-EntraUsersFromCSV.ps1" }
        @{ Name = "Export Filtered Entra Users"; Path = "$PSScriptRoot\Entra\Export-FilteredEntraUsers.ps1" }
        @{ Name = "Set Manager by Street Address"; Path = "$PSScriptRoot\Entra\Set-ManagerByStreetAddress.ps1" }
    )
    "Exchange" = @(
        @{ Name = "Find Alias Mailbox"; Path = "$PSScriptRoot\Exchange\Find-AliasMailbox.ps1" }
        @{ Name = "Get Shared Mailboxes from User"; Path = "$PSScriptRoot\Exchange\Get-SharedMailboxesFromUser.ps1" }
        @{ Name = "Export DL Members"; Path = "$PSScriptRoot\Exchange\Export-DLMembers.ps1" }
        @{ Name = "Import DL Members"; Path = "$PSScriptRoot\Exchange\Import-DLMembers.ps1" }
        @{ Name = "Import DL Members from CSV"; Path = "$PSScriptRoot\Exchange\Import-DLMembersFromCSV.ps1" }
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
    Write-Host "======================================" -ForegroundColor Cyan
    Write-Host "    SERVICEDESK TOOLS - V.0" -ForegroundColor Cyan
    Write-Host "======================================" -ForegroundColor Cyan
    Write-Host "Welcome to Servicedesk Tools - An idea by Servicedesk, made by Mr.Hagen - 2025/2026" -ForegroundColor Green 
    Write-Host "All scripts has been tested. Please let Alexander Hagen know if there are any issues." -ForegroundColor Yellow
    Write-Host "Remember that you should probably have User/Exchange PIM activated" -ForegroundColor Red
    Write-Host ""
    Write-Host "Select a category:" -ForegroundColor Yellow
    Write-Host "1: Entra ID (Mostly user management)"
    Write-Host "2: Exchange (Mailboxes and Distribution Lists)"
    Write-Host "3: Licenses (Manage user licenses)"
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
