import os
import shutil
import sys
from pathlib import Path

# Common install locations for PowerShell on macOS and Windows
_PWSH_SEARCH_PATHS = [
    "/usr/local/bin/pwsh",
    "/opt/homebrew/bin/pwsh",
    "/usr/local/microsoft/powershell/7/pwsh",
    os.path.expanduser("~/.dotnet/tools/pwsh"),
    r"C:\Program Files\PowerShell\7\pwsh.exe",
    r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
]


def find_pwsh() -> str | None:
    """Find the PowerShell executable, checking PATH and known install locations."""
    found = shutil.which("pwsh")
    if found:
        return found
    for path in _PWSH_SEARCH_PATHS:
        if os.path.isfile(path) and os.access(path, os.X_OK):
            return path
    return None


# Scripts root is always the GitHub-synced cache
_SCRIPTS_CACHE = Path.home() / ".servicedesk-tools" / "scripts" / "Azure"


def get_azure_scripts_root() -> Path:
    """Return the cached scripts directory (synced from GitHub)."""
    return _SCRIPTS_CACHE


SCRIPT_CATEGORIES: dict[str, list[dict[str, str]]] = {
    "Entra": [
        {"name": "Replace Department", "script": "Entra/Replace-Department.ps1"},
        {"name": "Replace Street Address", "script": "Entra/Replace-StreetAddress.ps1"},
        {"name": "Replace Office Location", "script": "Entra/Replace-OfficeLocation.ps1"},
        {"name": "Set Department/Office by Street Address", "script": "Entra/Set-DepartmentOfficeByStreetAddress.ps1"},
        {"name": "Set Manager by Street Address", "script": "Entra/Set-ManagerByStreetAddress.ps1"},
        {"name": "Import Entra Users", "script": "Entra/Import-EntraUsers.ps1"},
        {"name": "Update Entra Users from CSV", "script": "Entra/Update-EntraUsersFromCSV.ps1"},
        {"name": "Export Filtered Entra Users", "script": "Entra/Export-FilteredEntraUsers.ps1"},
        {"name": "Check Department Keywords", "script": "Entra/Check-DepartmentKeywords.ps1"},
    ],
    "Exchange": [
        {"name": "Find Alias Mailbox", "script": "Exchange/Find-Aliasmailbox.ps1"},
        {"name": "Export DL Members", "script": "Exchange/Export-DLMembers.ps1"},
        {"name": "Import DL Members", "script": "Exchange/Import-DLMembers.ps1"},
        {"name": "Import DL Members from CSV", "script": "Exchange/Import-DLMembersFromCSV.ps1"},
        {"name": "Import Contacts from CSV", "script": "Exchange/Import-ContactsFromCSV.ps1"},
        {"name": "Get User Distribution Lists", "script": "Exchange/Get-UserDistributionLists.ps1"},
        {"name": "Get Room Mailbox Settings", "script": "Exchange/Get-RoomMailboxSettings.ps1"},
        {"name": "Get Shared Mailboxes from User", "script": "Exchange/Get-SharedMailboxesFromUser.ps1"},
        {"name": "Compare UPN and Primary Email", "script": "Exchange/Compare-UPNandPrimaryEmail.ps1"},
        {"name": "New Dynamic Distribution List", "script": "Exchange/New-DynamicDistributionList.ps1"},
    ],
    "Licenses": [
        {"name": "Get Dynamics Licenses for Users", "script": "Licenses/Get-DynamicsLicensesForUsers.ps1"},
        {"name": "Remove Direct User License", "script": "Licenses/Remove-DirectUserLicense.ps1"},
    ],
    "Teams": [
        {"name": "Teams Reporting Tool", "script": "Teams/Teams-Reporting.ps1"},
    ],
    "On-Prem": [
        {"name": "Convert On-Prem DL to Cloud", "script": "On-Prem/Convert-OnPremDLToCloud.ps1"},
    ],
}

CATEGORY_DESCRIPTIONS: dict[str, str] = {
    "Entra": "Entra ID – User management",
    "Exchange": "Exchange – Mailboxes & Distribution Lists",
    "Licenses": "Licenses – Manage user licenses",
    "Teams": "Teams – Groups & channels reporting",
    "On-Prem": "On-Prem – On-premises Exchange tools",
}

CATEGORY_ICONS: dict[str, str] = {
    "Entra": "👤",
    "Exchange": "📧",
    "Licenses": "🔑",
    "Teams": "👥",
    "On-Prem": "🏢",
}

REQUIRED_MODULES: list[str] = [
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Identity.DirectoryManagement",
    "Microsoft.Graph.Authentication",
    "ExchangeOnlineManagement",
    "MicrosoftTeams",
]


def get_script_path(relative_script: str) -> Path:
    return get_azure_scripts_root() / relative_script


def script_exists(relative_script: str) -> bool:
    return get_script_path(relative_script).is_file()
