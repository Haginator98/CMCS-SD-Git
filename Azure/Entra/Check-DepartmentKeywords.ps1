Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All" -NoWelcome

Write-Host "Fetching users (DisplayName, UPN, Department)..." -ForegroundColor Cyan
$allUsers = Get-MgUser -All -Property "displayName,userPrincipalName,department,id"

$inputTerms = Read-Host "Enter letters/words to search for in Department (comma separated, e.g. it,sales,nor)"

if ([string]::IsNullOrWhiteSpace($inputTerms)) {
    Write-Host "No search terms provided. Exiting script." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    return
}

$terms = $inputTerms -split "," |
    ForEach-Object { $_.Trim() } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Select-Object -Unique

if ($terms.Count -eq 0) {
    Write-Host "No valid search terms found. Exiting script." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    return
}

$matches = foreach ($user in $allUsers) {
    $department = [string]$user.Department

    if ([string]::IsNullOrWhiteSpace($department)) {
        continue
    }

    $matchedTerms = $terms | Where-Object { $department -like "*$_*" }

    if ($matchedTerms.Count -gt 0) {
        [PSCustomObject]@{
            DisplayName        = $user.DisplayName
            UserPrincipalName  = $user.UserPrincipalName
            Department         = $department
            MatchedTerms       = ($matchedTerms -join ", ")
        }
    }
}

if (-not $matches -or $matches.Count -eq 0) {
    Write-Host "No users found where Department contains the provided letters/words." -ForegroundColor Yellow
} else {
    Write-Host "Found $($matches.Count) matching users:" -ForegroundColor Green
    $matches | Sort-Object Department, DisplayName | Format-Table -AutoSize

    $exportChoice = Read-Host "Export result to CSV? (y/n)"
    if ($exportChoice -eq "y") {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $outputPath = Join-Path -Path $PSScriptRoot -ChildPath "DepartmentKeywordMatches_$timestamp.csv"
        $matches | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Exported to: $outputPath" -ForegroundColor Cyan
    }
}

Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
Write-Host "Disconnected." -ForegroundColor Green
