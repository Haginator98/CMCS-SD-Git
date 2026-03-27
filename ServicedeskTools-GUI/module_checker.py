import subprocess
import threading
import shutil
from dataclasses import dataclass, field

import customtkinter as ctk

from script_registry import REQUIRED_MODULES


@dataclass
class ModuleStatus:
    name: str
    installed: bool = False
    version: str = ""
    update_available: bool = False
    online_version: str = ""


class ModuleChecker:
    """Check and manage required PowerShell modules."""

    def __init__(self):
        self.pwsh_path: str | None = shutil.which("pwsh")
        self.modules: list[ModuleStatus] = [
            ModuleStatus(name=m) for m in REQUIRED_MODULES
        ]

    @property
    def pwsh_available(self) -> bool:
        return self.pwsh_path is not None

    def check_modules(self, callback=None):
        """Check all modules in a background thread. Calls callback(modules) when done."""
        def _run():
            if not self.pwsh_available:
                if callback:
                    callback(self.modules)
                return

            for mod in self.modules:
                try:
                    result = subprocess.run(
                        [self.pwsh_path, "-NoProfile", "-Command",
                         f"(Get-Module -Name '{mod.name}' -ListAvailable | "
                         f"Select-Object -First 1).Version.ToString()"],
                        capture_output=True, text=True, timeout=30,
                    )
                    version = result.stdout.strip()
                    if version and not version.startswith("Error"):
                        mod.installed = True
                        mod.version = version
                    else:
                        mod.installed = False
                except Exception:
                    mod.installed = False

            if callback:
                callback(self.modules)

        thread = threading.Thread(target=_run, daemon=True)
        thread.start()

    def install_module(self, module_name: str, output_callback=None):
        """Install a module in a background thread."""
        def _run():
            if not self.pwsh_available:
                if output_callback:
                    output_callback(f"❌ PowerShell (pwsh) not found.\n")
                return

            if output_callback:
                output_callback(f"Installing {module_name}...\n")

            try:
                result = subprocess.run(
                    [self.pwsh_path, "-NoProfile", "-Command",
                     f"Install-Module -Name '{module_name}' -Scope CurrentUser "
                     f"-Force -AllowClobber -ErrorAction Stop"],
                    capture_output=True, text=True, timeout=300,
                )
                if result.returncode == 0:
                    if output_callback:
                        output_callback(f"✅ {module_name} installed successfully.\n")
                else:
                    if output_callback:
                        output_callback(f"❌ Failed: {result.stderr.strip()}\n")
            except Exception as e:
                if output_callback:
                    output_callback(f"❌ Error: {e}\n")

        thread = threading.Thread(target=_run, daemon=True)
        thread.start()
