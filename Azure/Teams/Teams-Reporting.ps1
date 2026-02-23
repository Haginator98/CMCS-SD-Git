param(
    [int]$Action
)

$ErrorActionPreference = 'Stop'

function Get-SafeValue {
    param(
        [object]$Value,
        [string]$Fallback
    )

    if ($null -eq $Value) {
        return $Fallback
    }

    $text = "$Value".Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $Fallback
    }

    return $text
}

function Confirm-Export {
    param(
        [string]$ReportName,
        [int]$ObjectCount
    )

    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "REPORT READY: $ReportName" -ForegroundColor Yellow
    Write-Host "Objects to export: $ObjectCount" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow

    $confirmation = Read-Host "Do you want to export this report? [Y] Yes [N] No"
    return ($confirmation -match '^[yY]$')
}

function Export-Report {
    param(
        [System.Collections.IEnumerable]$Data,
        [string]$BaseFileName
    )

    $items = @($Data)
    if ($items.Count -eq 0) {
        Write-Host "No data found. Nothing to export." -ForegroundColor Yellow
        return
    }

    if (-not (Confirm-Export -ReportName $BaseFileName -ObjectCount $items.Count)) {
        Write-Host "Export cancelled by user." -ForegroundColor Yellow
        return
    }

    $desktopPath = [Environment]::GetFolderPath('Desktop')
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $safeFileName = ($BaseFileName -replace '[\\/:*?"<>|]', '_')
    $exportPath = Join-Path $desktopPath ("$safeFileName`_$timestamp.csv")

    $items | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nReport exported to: $exportPath" -ForegroundColor Green
}

function Show-Menu {
    Write-Host ""
    Write-Host "`nMicrosoft Teams Reporting" -ForegroundColor Yellow
    Write-Host "    1.All Teams in organization" -ForegroundColor Cyan
    Write-Host "    2.All Teams members and owners report" -ForegroundColor Cyan
    Write-Host "    3.Specific Teams' members and Owners report" -ForegroundColor Cyan
    Write-Host "    4.All Teams' owners report" -ForegroundColor Cyan
    Write-Host "    5.Specific Teams' owners report" -ForegroundColor Cyan
    Write-Host "`nTeams Channel Reporting" -ForegroundColor Yellow
    Write-Host "    6.All channels in organization" -ForegroundColor Cyan
    Write-Host "    7.All channels in specific Team" -ForegroundColor Cyan
    Write-Host "    8.Members and Owners Report of Single Channel" -ForegroundColor Cyan
    Write-Host "    0.Exit" -ForegroundColor Cyan
    Write-Host ""
}

function Get-TeamByDisplayName {
    param(
        [string]$DisplayName
    )

    $team = Get-Team -DisplayName $DisplayName -ErrorAction SilentlyContinue
    if ($null -eq $team) {
        Write-Host "Team '$DisplayName' was not found." -ForegroundColor Red
        return $null
    }

    return $team
}

$connected = $false

try {
    $module = Get-Module -Name MicrosoftTeams -ListAvailable
    if ($module.Count -eq 0) {
        Write-Host "MicrosoftTeams module is not available." -ForegroundColor Yellow
        $installConfirm = Read-Host "Do you want to install MicrosoftTeams module now? [Y] Yes [N] No"
        if ($installConfirm -match '^[yY]$') {
            Install-Module MicrosoftTeams -Scope CurrentUser -Force -AllowClobber
        }
        else {
            Write-Host "MicrosoftTeams module is required. Script cancelled." -ForegroundColor Red
            return
        }
    }

    Write-Host "Importing Microsoft Teams module..." -ForegroundColor Yellow
    Import-Module MicrosoftTeams -ErrorAction Stop

    Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Yellow
    Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
    $connected = $true
    Write-Host "Successfully connected to Microsoft Teams." -ForegroundColor Green

    do {
        if ($null -eq $Action) {
            Show-Menu
            $choice = Read-Host 'Please choose the action to continue'
        }
        else {
            $choice = "$Action"
        }

        switch ($choice) {
            '1' {
                Write-Host "Collecting all Teams in organization..." -ForegroundColor Cyan
                $teams = @(Get-Team -ErrorAction SilentlyContinue)

                $results = @()
                $count = 0
                foreach ($team in $teams) {
                    $count++
                    Write-Progress -Activity "Processing Teams" -Status "Team $count of $($teams.Count): $($team.DisplayName)" -PercentComplete (($count / [Math]::Max($teams.Count, 1)) * 100)

                    $channels = @(Get-TeamChannel -GroupId $team.GroupId -ErrorAction SilentlyContinue)
                    $users = @(Get-TeamUser -GroupId $team.GroupId -ErrorAction SilentlyContinue)

                    $results += [PSCustomObject]@{
                        'Teams Name'         = Get-SafeValue -Value $team.DisplayName -Fallback 'no team name'
                        'Team Type'          = Get-SafeValue -Value $team.Visibility -Fallback 'no team type'
                        'Mail Nick Name'     = Get-SafeValue -Value $team.MailNickName -Fallback 'no mail nickname'
                        'Description'        = Get-SafeValue -Value $team.Description -Fallback 'no description'
                        'Archived Status'    = Get-SafeValue -Value $team.Archived -Fallback 'no archived status'
                        'Channel Count'      = $channels.Count
                        'Team Members Count' = $users.Count
                        'Team Owners Count'  = (@($users | Where-Object { $_.Role -eq 'Owner' })).Count
                    }
                }

                Write-Progress -Activity "Processing Teams" -Completed
                Export-Report -Data $results -BaseFileName 'All Teams Report'
            }

            '2' {
                Write-Host "Collecting all Teams members and owners..." -ForegroundColor Cyan
                $teams = @(Get-Team -ErrorAction SilentlyContinue)

                $results = @()
                $count = 0
                foreach ($team in $teams) {
                    $count++
                    Write-Progress -Activity "Processing Teams" -Status "Team $count of $($teams.Count): $($team.DisplayName)" -PercentComplete (($count / [Math]::Max($teams.Count, 1)) * 100)

                    $users = @(Get-TeamUser -GroupId $team.GroupId -ErrorAction SilentlyContinue)
                    foreach ($user in $users) {
                        $results += [PSCustomObject]@{
                            'Teams Name'  = Get-SafeValue -Value $team.DisplayName -Fallback 'no team name'
                            'Member Name' = Get-SafeValue -Value $user.Name -Fallback 'no member name'
                            'Member Mail' = Get-SafeValue -Value $user.User -Fallback 'no member mail'
                            'Role'        = Get-SafeValue -Value $user.Role -Fallback 'no role'
                        }
                    }
                }

                Write-Progress -Activity "Processing Teams" -Completed
                Export-Report -Data $results -BaseFileName 'All Teams Members and Owners Report'
            }

            '3' {
                $teamName = Read-Host "Enter Teams name to get members and owners report (Case sensitive)"
                $team = Get-TeamByDisplayName -DisplayName $teamName
                if ($null -eq $team) {
                    break
                }

                Write-Host "Collecting members and owners for '$($team.DisplayName)'..." -ForegroundColor Cyan
                $users = @(Get-TeamUser -GroupId $team.GroupId -ErrorAction SilentlyContinue)

                $results = foreach ($user in $users) {
                    [PSCustomObject]@{
                        'Member Name' = Get-SafeValue -Value $user.Name -Fallback 'no member name'
                        'Member Mail' = Get-SafeValue -Value $user.User -Fallback 'no member mail'
                        'Role'        = Get-SafeValue -Value $user.Role -Fallback 'no role'
                    }
                }

                Export-Report -Data $results -BaseFileName "Members and Owners - $($team.DisplayName)"
            }

            '4' {
                Write-Host "Collecting all Teams owners..." -ForegroundColor Cyan
                $teams = @(Get-Team -ErrorAction SilentlyContinue)

                $results = @()
                $count = 0
                foreach ($team in $teams) {
                    $count++
                    Write-Progress -Activity "Processing Teams" -Status "Team $count of $($teams.Count): $($team.DisplayName)" -PercentComplete (($count / [Math]::Max($teams.Count, 1)) * 100)

                    $owners = @(Get-TeamUser -GroupId $team.GroupId -ErrorAction SilentlyContinue | Where-Object { $_.Role -eq 'Owner' })
                    foreach ($owner in $owners) {
                        $results += [PSCustomObject]@{
                            'Teams Name' = Get-SafeValue -Value $team.DisplayName -Fallback 'no team name'
                            'Owner Name' = Get-SafeValue -Value $owner.Name -Fallback 'no owner name'
                            'Owner Mail' = Get-SafeValue -Value $owner.User -Fallback 'no owner mail'
                        }
                    }
                }

                Write-Progress -Activity "Processing Teams" -Completed
                Export-Report -Data $results -BaseFileName 'All Teams Owners Report'
            }

            '5' {
                $teamName = Read-Host "Enter Teams name to get owners report (Case sensitive)"
                $team = Get-TeamByDisplayName -DisplayName $teamName
                if ($null -eq $team) {
                    break
                }

                Write-Host "Collecting owners for '$($team.DisplayName)'..." -ForegroundColor Cyan
                $owners = @(Get-TeamUser -GroupId $team.GroupId -ErrorAction SilentlyContinue | Where-Object { $_.Role -eq 'Owner' })

                $results = foreach ($owner in $owners) {
                    [PSCustomObject]@{
                        'Owner Name' = Get-SafeValue -Value $owner.Name -Fallback 'no owner name'
                        'Owner Mail' = Get-SafeValue -Value $owner.User -Fallback 'no owner mail'
                    }
                }

                Export-Report -Data $results -BaseFileName "Owners - $($team.DisplayName)"
            }

            '6' {
                Write-Host "Collecting all channels in organization..." -ForegroundColor Cyan
                $teams = @(Get-Team -ErrorAction SilentlyContinue)

                $results = @()
                $count = 0
                foreach ($team in $teams) {
                    $count++
                    Write-Progress -Activity "Processing Teams" -Status "Team $count of $($teams.Count): $($team.DisplayName)" -PercentComplete (($count / [Math]::Max($teams.Count, 1)) * 100)

                    $channels = @(Get-TeamChannel -GroupId $team.GroupId -ErrorAction SilentlyContinue)
                    foreach ($channel in $channels) {
                        $channelUsers = @(Get-TeamChannelUser -GroupId $team.GroupId -DisplayName $channel.DisplayName -ErrorAction SilentlyContinue)

                        $results += [PSCustomObject]@{
                            'Teams Name'          = Get-SafeValue -Value $team.DisplayName -Fallback 'no team name'
                            'Channel Name'        = Get-SafeValue -Value $channel.DisplayName -Fallback 'no channel name'
                            'Membership Type'     = Get-SafeValue -Value $channel.MembershipType -Fallback 'no membership type'
                            'Description'         = Get-SafeValue -Value $channel.Description -Fallback 'no description'
                            'Owners Count'        = (@($channelUsers | Where-Object { $_.Role -eq 'Owner' })).Count
                            'Total Members Count' = $channelUsers.Count
                        }
                    }
                }

                Write-Progress -Activity "Processing Teams" -Completed
                Export-Report -Data $results -BaseFileName 'All Channels Report'
            }

            '7' {
                $teamName = Read-Host "Enter Teams name (Case Sensitive)"
                $team = Get-TeamByDisplayName -DisplayName $teamName
                if ($null -eq $team) {
                    break
                }

                Write-Host "Collecting channels for '$($team.DisplayName)'..." -ForegroundColor Cyan
                $channels = @(Get-TeamChannel -GroupId $team.GroupId -ErrorAction SilentlyContinue)

                $results = @()
                $count = 0
                foreach ($channel in $channels) {
                    $count++
                    Write-Progress -Activity "Processing Channels" -Status "Channel $count of $($channels.Count): $($channel.DisplayName)" -PercentComplete (($count / [Math]::Max($channels.Count, 1)) * 100)

                    $channelUsers = @(Get-TeamChannelUser -GroupId $team.GroupId -DisplayName $channel.DisplayName -ErrorAction SilentlyContinue)
                    $results += [PSCustomObject]@{
                        'Teams Name'          = Get-SafeValue -Value $team.DisplayName -Fallback 'no team name'
                        'Channel Name'        = Get-SafeValue -Value $channel.DisplayName -Fallback 'no channel name'
                        'Membership Type'     = Get-SafeValue -Value $channel.MembershipType -Fallback 'no membership type'
                        'Description'         = Get-SafeValue -Value $channel.Description -Fallback 'no description'
                        'Owners Count'        = (@($channelUsers | Where-Object { $_.Role -eq 'Owner' })).Count
                        'Total Members Count' = $channelUsers.Count
                    }
                }

                Write-Progress -Activity "Processing Channels" -Completed
                Export-Report -Data $results -BaseFileName "Channels - $($team.DisplayName)"
            }

            '8' {
                $teamName = Read-Host "Enter Teams name in which channel resides (Case sensitive)"
                $channelName = Read-Host "Enter Channel name"

                $team = Get-TeamByDisplayName -DisplayName $teamName
                if ($null -eq $team) {
                    break
                }

                Write-Host "Collecting members and owners for channel '$channelName' in '$($team.DisplayName)'..." -ForegroundColor Cyan
                $channelUsers = @(Get-TeamChannelUser -GroupId $team.GroupId -DisplayName $channelName -ErrorAction SilentlyContinue)

                $results = foreach ($user in $channelUsers) {
                    [PSCustomObject]@{
                        'Teams Name'   = Get-SafeValue -Value $team.DisplayName -Fallback 'no team name'
                        'Channel Name' = Get-SafeValue -Value $channelName -Fallback 'no channel name'
                        'Member Name'  = Get-SafeValue -Value $user.Name -Fallback 'no member name'
                        'Member Mail'  = Get-SafeValue -Value $user.User -Fallback 'no member mail'
                        'Role'         = Get-SafeValue -Value $user.Role -Fallback 'no role'
                    }
                }

                Export-Report -Data $results -BaseFileName "Channel Members and Owners - $($team.DisplayName) - $channelName"
            }

            '0' {
                Write-Host "Operation cancelled by user." -ForegroundColor Yellow
                if ($connected) {
                    Write-Host "Disconnecting from Microsoft Teams..." -ForegroundColor Yellow
                }
                break
            }

            default {
                Write-Host "Invalid selection. Please choose 0-8." -ForegroundColor Red
            }
        }

        if ($null -ne $Action) {
            break
        }
    }
    while ($true)
}
catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Script stopped because the operation could not continue." -ForegroundColor Red
    if ($connected) {
        Write-Host "Disconnecting from Microsoft Teams due to error..." -ForegroundColor Yellow
    }
}
finally {
    if ($connected) {
        Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Disconnected from Microsoft Teams." -ForegroundColor Green
    }
}
