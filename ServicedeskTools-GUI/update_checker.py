import json
import threading
import urllib.request
import urllib.error
import webbrowser

GITHUB_REPO = "Haginator98/CMCS-SD-Git"
RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
RELEASES_PAGE = f"https://github.com/{GITHUB_REPO}/releases/latest"


def _parse_version(tag: str) -> tuple[int, ...]:
    """Parse a version tag like 'v1.2.0' or '1.2' into a comparable tuple."""
    clean = tag.lstrip("vV").strip()
    parts = []
    for p in clean.split("."):
        try:
            parts.append(int(p))
        except ValueError:
            break
    return tuple(parts) if parts else (0,)


def check_for_update(current_version: str, callback=None):
    """Check GitHub Releases for a newer version.

    callback(has_update: bool, latest_tag: str, download_url: str, message: str)
    """
    def _run():
        try:
            req = urllib.request.Request(
                RELEASES_API,
                headers={"Accept": "application/vnd.github.v3+json"},
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode())

            latest_tag = data.get("tag_name", "")
            release_name = data.get("name", latest_tag)
            html_url = data.get("html_url", RELEASES_PAGE)

            current = _parse_version(current_version)
            latest = _parse_version(latest_tag)

            if latest > current:
                if callback:
                    callback(
                        True,
                        latest_tag,
                        html_url,
                        f"🆕 New version available: {release_name} ({latest_tag})",
                    )
            else:
                if callback:
                    callback(False, latest_tag, html_url, "")

        except urllib.error.HTTPError as e:
            if e.code == 404:
                # No releases yet
                if callback:
                    callback(False, "", "", "")
            else:
                if callback:
                    callback(False, "", "", f"Update check failed: {e}")
        except Exception:
            # Network error — silently ignore
            if callback:
                callback(False, "", "", "")

    thread = threading.Thread(target=_run, daemon=True)
    thread.start()


def open_releases_page():
    """Open the GitHub releases page in the default browser."""
    webbrowser.open(RELEASES_PAGE)
