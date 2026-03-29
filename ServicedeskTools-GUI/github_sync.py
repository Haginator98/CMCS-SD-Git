import os
import json
import shutil
import threading
import urllib.request
import urllib.error
from pathlib import Path

GITHUB_REPO = "Haginator98/CMCS-SD-Git"
GITHUB_BRANCH = "main"
GITHUB_API_BASE = f"https://api.github.com/repos/{GITHUB_REPO}"
GITHUB_RAW_BASE = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}"
AZURE_SUBDIR = "Azure"


def get_scripts_cache_dir() -> Path:
    """Return the local cache directory for downloaded scripts."""
    cache_base = Path.home() / ".servicedesk-tools" / "scripts"
    cache_base.mkdir(parents=True, exist_ok=True)
    return cache_base


def _api_get(path: str) -> list | dict:
    """Make a GET request to the GitHub API."""
    url = f"{GITHUB_API_BASE}/{path}"
    req = urllib.request.Request(url, headers={"Accept": "application/vnd.github.v3+json"})
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read().decode())


def _download_file(raw_path: str, dest: Path):
    """Download a single file from GitHub raw."""
    url = f"{GITHUB_RAW_BASE}/{raw_path}"
    dest.parent.mkdir(parents=True, exist_ok=True)
    urllib.request.urlretrieve(url, str(dest))


def _list_files_recursive(repo_path: str) -> list[str]:
    """List all files under a GitHub repo path (recursive)."""
    files = []
    try:
        items = _api_get(f"contents/{repo_path}?ref={GITHUB_BRANCH}")
        for item in items:
            if item["type"] == "file" and item["name"].endswith(".ps1"):
                files.append(item["path"])
            elif item["type"] == "dir":
                files.extend(_list_files_recursive(item["path"]))
    except Exception:
        pass
    return files


class GitHubSyncer:
    """Syncs PowerShell scripts from GitHub to a local cache."""

    def __init__(self):
        self.cache_dir = get_scripts_cache_dir()
        self.azure_dir = self.cache_dir / AZURE_SUBDIR

    @property
    def scripts_root(self) -> Path:
        return self.azure_dir

    def is_cached(self) -> bool:
        """Check if we have any cached scripts."""
        return self.azure_dir.is_dir() and any(self.azure_dir.rglob("*.ps1"))

    def sync(self, progress_callback=None, done_callback=None):
        """Sync scripts from GitHub in a background thread.

        progress_callback(message: str) — called with status updates
        done_callback(success: bool, message: str) — called when done
        """
        def _run():
            try:
                if progress_callback:
                    progress_callback("🔄 Connecting to GitHub...\n")

                ps1_files = _list_files_recursive(AZURE_SUBDIR)

                if not ps1_files:
                    if done_callback:
                        done_callback(False, "No .ps1 scripts found in repo.")
                    return

                if progress_callback:
                    progress_callback(f"📥 Found {len(ps1_files)} scripts. Downloading...\n")

                for i, file_path in enumerate(ps1_files, 1):
                    # file_path is like "Azure/Entra/Replace-Department.ps1"
                    local_path = self.cache_dir / file_path
                    if progress_callback:
                        name = Path(file_path).name
                        progress_callback(f"  [{i}/{len(ps1_files)}] {name}\n")
                    _download_file(file_path, local_path)

                # Write a sync marker
                marker = self.cache_dir / ".last_sync"
                marker.write_text(f"Synced {len(ps1_files)} files from {GITHUB_REPO}")

                if done_callback:
                    done_callback(True, f"✅ Synced {len(ps1_files)} scripts from GitHub.")

            except urllib.error.URLError as e:
                if done_callback:
                    done_callback(False, f"❌ Network error: {e.reason}")
            except Exception as e:
                if done_callback:
                    done_callback(False, f"❌ Sync failed: {e}")

        thread = threading.Thread(target=_run, daemon=True)
        thread.start()

    def clear_cache(self):
        """Remove all cached scripts."""
        if self.cache_dir.exists():
            shutil.rmtree(self.cache_dir)
            self.cache_dir.mkdir(parents=True, exist_ok=True)
