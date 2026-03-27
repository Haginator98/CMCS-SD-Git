import os
from pathlib import Path


def get_azure_scripts_root() -> Path:
    """Resolve the Azure scripts root directory (sibling to ServicedeskTools-GUI)."""
    gui_dir = Path(__file__).resolve().parent
    repo_root = gui_dir.parent
    azure_root = repo_root / "Azure"
    return azure_root


SCRIPT_CATEGORIES: dict[str, list[dict[str, str]]] = {
    "Entra": [
        {"name": "Replace Department", "script": "Entra/Replace-Department.ps1"},
        {"name": "Replace Street Address", "script": "Entra/Replace-StreetAddress.ps1"},
        {"name": "Set Department/Office by Street Address", "script": "Entra/Set-DepartmentOfficeByStreetAddress.ps1"},
        {"name": "Set Manager by Street Address", "script": "Entra/Set-ManagerByStreetAddress.ps1"},
        {"name": "Import Entra Users", "script": "Entra/Import-EntraUsers.ps1"},
        {"name": "Update Entra Users from CSV", "script": "Entra/Update-EntraUsersFromCSV.ps1"},
        {"name": "Export Filtered Entra Users", "script": "Entra/Export-FilteredEntraUsers.ps1"},
        {"name": "Check Department Keywords", "script": "Entra/Check-DepartmentKeywords.ps1"},
    ],
    "Exchange": [
        {"name": "Find Alias Mailbox", "script": "Exchange/Find-AliasMailbox.ps1"},
        {"name": "Export DL Members", "script": "Exchange/Export-DLMembers.ps1"},
        {"name": "Import DL Members", "script": "Exchange/Import-DLMembers.ps1"},
        {"name": "Import DL Members from CSV", "script": "Exchange/Import-DLMembersFromCSV.ps1"},
        {"name": "Import Contacts from CSV", "script": "Exchange/Import-ContactsFromCSV.ps1"},
        {"name": "Get User Distribution Lists", "script": "Exchange/Get-UserDistributionLists.ps1"},
        {"name": "Get Room Mailbox Settings", "script": "Exchange/Get-RoomMailboxSettings.ps1"},
        {"name": "Get Shared Mailboxes from User", "script": "Exchange/Get-SharedMailboxesFromUser.ps1"},
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
