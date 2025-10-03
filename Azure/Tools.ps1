#This script is used to choose what script to run in Azure AD / Entra ID
# List available scripts
$scripts = @(
    @{ Name = "Entra ID - Change Department"; Path = ".\\Entra\\change department.ps1" }
    @{ Name = "Entra ID - Change Street Address"; Path = ".\\Entra\\change street address Entra.ps1" }
    # @{ Name = "Entra ID - Change Dep. Office based on Streed Address"; Path = ".\\Entra\\change dep.office based on street address.ps1" } 
    @{ Name = "Exchange - Find Alias Mailbox"; Path = ".\\Exchange\\Find-Aliasmailbox.ps1" }
    # @{ Name = "Exchange - Get shared mailboxes from user"; Path = ".\\Exchange\\Get shared mailboxes from user.ps1" }
    # @{ Name = "Exchange - Calendar acces" ; Path = ".\\Exchange\\Calendar Access.ps1" }
    # @{ Name = "Exchange - Add Distribution Lists to users"; Path = ".\\Exchange\\Add DLs to users.ps1" }
)

while ($true) {
    Clear-Host
    # SERVICEDESK - Matrix Style Logo
    Write-Host "
░██████╗███████╗██████╗░██╗░░░██╗██╗░█████╗░███████╗██████╗░███████╗░██████╗██╗░░██╗
██╔════╝██╔════╝██╔══██╗██║░░░██║██║██╔══██╗██╔════╝██╔══██╗██╔════╝██╔════╝██║░██╔╝
╚█████╗░█████╗░░██████╔╝╚██╗░██╔╝██║██║░░╚═╝█████╗░░██║░░██║█████╗░░╚█████╗░█████═╝░
░╚═══██╗██╔══╝░░██╔══██╗░╚████╔╝░██║██║░░██╗██╔══╝░░██║░░██║██╔══╝░░░╚═══██╗██╔═██╗░
██████╔╝███████╗██║░░██║░░╚██╔╝░░██║╚█████╔╝███████╗██████╔╝███████╗██████╔╝██║░╚██╗
╚═════╝░╚══════╝╚═╝░░╚═╝░░░╚═╝░░░╚═╝░╚════╝░╚══════╝╚═════╝░╚══════╝╚═════╝░╚═╝░░╚═╝" -ForegroundColor Cyan
    Write-Host "Welcome to Servicedesk Tools!"-ForegroundColor Cyan
    Write-Host "Remember, not all scripts are tested or fully functional! Use at your own risk!" -ForegroundColor Yellow
    Write-Host "Some scripts requires you to already be logged in" -ForegroundColor Yellow
    Write-Host "Select a script to run:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $scripts.Count; $i++) {
        Write-Host "$($i+1): $($scripts[$i].Name)"
    }
    Write-Host "0: Exit"

    $choice = Read-Host "Enter the number of the script to run"
    if ($choice -eq '0') { break }
    if ($choice -eq 'exit') { break }
    if ($choice -match '^[1-9][0-9]*$' -and $choice -le $scripts.Count) {
        $scriptToRun = $scripts[$choice-1].Path
        Write-Host "Running $($scripts[$choice-1].Name)..." -ForegroundColor Cyan
        & $scriptToRun
        Write-Host "`nScript finished. Returning to menu..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
    } else {
        Write-Host "Invalid selection. Try again." -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

