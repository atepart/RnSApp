from __future__ import annotations

import json
import logging
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass
from typing import Optional

import requests


class UpdateError(RuntimeError):
    def __init__(self, message: str, *, url: str | None = None, status: int | None = None, body: str | None = None):
        super().__init__(message)
        self.url = url
        self.status = status
        self.body = body

    def __str__(self) -> str:
        parts = [super().__str__()]
        if self.url:
            parts.append(f"URL: {self.url}")
        if self.status is not None:
            parts.append(f"HTTP: {self.status}")
        if self.body:
            snippet = self.body
            if len(snippet) > 800:
                snippet = snippet[:800] + "…"
            parts.append(f"Body: {snippet}")
        return "\n".join(parts)


@dataclass
class ReleaseAsset:
    name: str
    download_url: str
    size: int | None = None


@dataclass
class ReleaseInfo:
    tag: str
    published_at: str | None
    asset: Optional[ReleaseAsset]
    prerelease: bool | None = None


logger = logging.getLogger(__name__)

_TAG_RE = re.compile(r"^(?P<prefix>new)(?P<num>\d+)(?:b(?P<beta>\d+))?$")


def parse_tag(tag: str) -> tuple[int, int | None]:
    """Parse version tag like 'new103b1' into a tuple (103, 1).

    Returns (version_number, beta_number or None).
    Raises ValueError if tag does not match expected format.
    """
    m = _TAG_RE.match(tag.strip())
    if not m:
        raise ValueError(f"Unsupported tag format: {tag}")
    num = int(m.group("num"))
    beta = m.group("beta")
    return num, (int(beta) if beta is not None else None)


def compare_tags(a: str, b: str) -> int:
    """Compare two tags by version semantics.

    Returns -1 if a<b, 0 if equal, 1 if a>b.
    Stable (no beta) is considered newer than any beta of the same number.
    """
    try:
        a_num, a_beta = parse_tag(a)
        b_num, b_beta = parse_tag(b)
    except ValueError:
        # Fallback to lexicographic if unparsable
        return (a > b) - (a < b)

    if a_num != b_num:
        return (a_num > b_num) - (a_num < b_num)
    # Same base version: stable (None) > beta
    if a_beta is None and b_beta is None:
        return 0
    if a_beta is None:
        return 1
    if b_beta is None:
        return -1
    return (a_beta > b_beta) - (a_beta < b_beta)


def _request_headers() -> dict[str, str]:
    # Mimic curl for compatibility (as user's curl works reliably)
    hdrs = {
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "curl/8.0.1",
    }
    token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    if token:
        hdrs["Authorization"] = f"Bearer {token}"
    return hdrs


def _http_get_json(url: str) -> list[dict] | dict:
    logger.info(f"GET {url}")
    headers = _request_headers()
    last_error: UpdateError | None = None
    try:
        # Connect timeout 10s, read timeout 10s
        resp = requests.get(url, headers=headers, timeout=(5, 10))
        status = resp.status_code
        logger.debug(f"Response status: {status}")
        if status >= 400:
            body = None
            try:
                j = resp.json()
                body = j.get("message") if isinstance(j, dict) else str(j)
            except Exception:
                body = resp.text
            raise UpdateError("GitHub API error", url=url, status=status, body=body)
        return resp.json()
    except requests.Timeout as e:
        logger.error(f"GitHub API timeout: {e}")
        last_error = UpdateError("GitHub API timeout (connect 5s, read 10s)", url=url)
    except requests.SSLError as e:
        logger.error(f"GitHub SSL error: {e}")
        last_error = UpdateError("GitHub SSL error", url=url, body=str(e))
    except requests.ConnectionError as e:
        logger.error(f"GitHub connection error: {e}")
        last_error = UpdateError("GitHub connection error", url=url, body=str(e))
    except requests.RequestException as e:
        logger.error(f"GitHub request error: {e}")
        last_error = UpdateError("GitHub request error", url=url, body=str(e))
    except ValueError as e:
        logger.error(f"GitHub API returned invalid JSON: {e}")
        last_error = UpdateError("GitHub API returned invalid JSON", url=url, body=str(e))

    if last_error is None:
        raise

    # Fallback: try curl (uses system certificates; often succeeds where requests fails)
    try:
        return _http_get_json_via_curl(url, headers=headers)
    except Exception as fallback_err:
        logger.error(f"Curl fallback failed: {fallback_err}")
        if isinstance(fallback_err, UpdateError):
            raise fallback_err
        raise last_error from fallback_err


def _http_get_json_via_curl(url: str, headers: dict[str, str]) -> list[dict] | dict:
    """Fetch JSON using curl as a fallback (better system CA compatibility)."""
    if not shutil.which("curl"):
        raise UpdateError("curl not found for GitHub API fallback", url=url)

    cmd = [
        "curl",
        "-sSL",
        "--connect-timeout",
        "5",
        "--max-time",
        "10",
    ]
    for k, v in (headers or {}).items():
        cmd.extend(["-H", f"{k}: {v}"])
    # Write HTTP status to the end of stdout for parsing
    cmd.extend(["-w", "\n%{http_code}\n", url])

    proc = subprocess.run(cmd, capture_output=True, text=True)
    stdout = proc.stdout or ""
    stderr = proc.stderr or ""
    if proc.returncode != 0:
        raise UpdateError("GitHub API error (curl)", url=url, body=stderr or stdout)

    body = stdout
    status_code: int | None = None
    if "\n" in stdout:
        # Split off the trailing status line added by -w
        body, _, status_line = stdout.rstrip("\n").rpartition("\n")
        try:
            status_code = int(status_line.strip())
        except ValueError:
            body = stdout
            status_code = None

    if status_code and status_code >= 400:
        raise UpdateError("GitHub API error (curl)", url=url, status=status_code, body=body.strip())

    try:
        return json.loads(body)
    except json.JSONDecodeError as e:
        raise UpdateError("GitHub API returned invalid JSON (curl)", url=url, body=body[:800]) from e


def _http_download(url: str, dest_path: str, progress_cb=None, should_cancel=None):
    logger.info(f"Downloading asset: {url}")
    try:
        # Strict timeouts: connect 10s, read 10s
        with requests.get(url, stream=True, timeout=(10, 10)) as r:
            status = r.status_code
            if status >= 400:
                raise UpdateError("Asset download error", url=url, status=status, body=r.text)
            total = int(r.headers.get("Content-Length", "0") or 0)
            downloaded = 0
            chunk = 64 * 1024
            with open(dest_path, "wb") as f:
                for part in r.iter_content(chunk_size=chunk):
                    if not part:
                        continue
                    if should_cancel and should_cancel():
                        raise UpdateError("Download cancelled by user")
                    f.write(part)
                    downloaded += len(part)
                    if progress_cb:
                        try:
                            progress_cb(downloaded, total)
                        except Exception:
                            pass
    except requests.Timeout as e:
        logger.error(f"Download timeout: {e}")
        raise UpdateError("Download timeout (10s)", url=url) from e
    except requests.ConnectionError as e:
        logger.error(f"Download connection error: {e}")
        raise UpdateError("Download connection error", url=url) from e
    except requests.RequestException as e:
        logger.error(f"Download failed: {e}")
        raise UpdateError("Download request error", url=url) from e


def _is_frozen() -> bool:
    return getattr(sys, "frozen", False) is True


def _app_dir() -> str:
    if _is_frozen():
        return os.path.dirname(sys.executable)
    # Dev mode: use project root main directory
    return os.path.abspath(os.path.dirname(sys.argv[0]))


def _normalize_arch() -> str:
    m = platform.machine().lower()
    # Common normalizations
    if m in ("x86_64", "amd64", "x64"):
        return "x64"
    if m in ("i386", "i686", "x86"):
        return "x86"
    if m in ("arm64", "aarch64"):
        return "arm64"
    return m or "unknown"


def _select_asset_for_current_platform(tag: str, assets: list[dict]) -> Optional[ReleaseAsset]:
    """Select a release asset that matches current OS/arch with flexible naming.

    Tries multiple heuristics to match typical naming patterns. Prefers assets whose
    name contains the tag and ends with .zip. Falls back to any OS/arch match.
    """
    system = platform.system().lower()
    arch = _normalize_arch()

    def is_zip(name: str) -> bool:
        return name.lower().endswith(".zip")

    def match_os(name: str) -> bool:
        n = name.lower()
        if system == "windows":
            return any(k in n for k in ("windows", "win"))
        if system == "darwin":
            return any(k in n for k in ("macos", "darwin", "osx", "mac"))
        return False

    def match_arch(name: str) -> bool:
        n = name.lower()
        if arch == "x64":
            return any(k in n for k in ("x64", "x86_64", "amd64", "win64", "64bit")) or not any(
                k in n for k in ("x86", "i386", "win32", "32bit")
            )
        if arch == "x86":
            return any(k in n for k in ("x86", "i386", "win32", "32bit"))
        if arch == "arm64":
            return any(k in n for k in ("arm64", "aarch64"))
        # Unknown arch: don't filter out
        return True

    candidates: list[ReleaseAsset] = []
    for a in assets or []:
        if not isinstance(a, dict):
            continue
        name = a.get("name", "") or ""
        url = a.get("browser_download_url", "") or ""
        if not name or not url or not is_zip(name):
            continue
        if not match_os(name):
            continue
        if not match_arch(name):
            continue
        candidates.append(
            ReleaseAsset(
                name=name,
                download_url=url,
                size=a.get("size"),
            )
        )

    if not candidates:
        return None

    # Prefer names containing the tag; then longer names (more specific)
    def score(asset: ReleaseAsset) -> tuple[int, int]:
        n = asset.name.lower()
        return (1 if tag.lower() in n else 0, len(n))

    candidates.sort(key=score, reverse=True)
    return candidates[0]

    # Note: function moved above with improved heuristics


def find_latest_release(repo_slug: str) -> Optional[ReleaseInfo]:
    """Fetch releases and return the latest by tag semantics (then date)."""
    url = f"https://api.github.com/repos/{repo_slug}/releases"
    try:
        releases = _http_get_json(url)
    except Exception:
        return None
    if not isinstance(releases, list):
        return None

    # Keep only tags matching our scheme, prepare tuples for robust sorting
    prepared: list[tuple[tuple[int, int, int], dict]] = []
    for r in releases:
        tag = r.get("tag_name") or r.get("name") or ""
        if not _TAG_RE.match(tag):
            continue
        try:
            num, beta = parse_tag(tag)
        except ValueError:
            continue
        stable_flag = 1 if beta is None else 0  # stable preferred
        beta_val = beta or 0
        prepared.append(((num, stable_flag, beta_val), r))

    if not prepared:
        return None

    prepared.sort(key=lambda it: (it[0][0], it[0][1], it[0][2], (it[1].get("published_at") or "")), reverse=True)
    latest = prepared[0][1]
    tag = latest.get("tag_name") or latest.get("name")
    published_at = latest.get("published_at")
    asset = _select_asset_for_current_platform(tag, latest.get("assets", []) or [])
    return ReleaseInfo(tag=tag, published_at=published_at, asset=asset, prerelease=latest.get("prerelease"))


def list_releases(repo_slug: str, limit: int = 10) -> list[ReleaseInfo]:
    """Return up to `limit` latest releases as-is from GitHub.

    No tag-format filtering; we rely on published order and platform asset matching.
    """
    url = f"https://api.github.com/repos/{repo_slug}/releases?per_page={max(1, min(limit, 100))}"
    logger.info(f"Fetching releases: {url}")
    releases = _http_get_json(url)
    if not isinstance(releases, list):
        raise UpdateError("Unexpected GitHub API response (not a list)", url=url, body=str(releases))

    items: list[ReleaseInfo] = []
    for r in releases:
        tag = r.get("tag_name") or r.get("name") or ""
        ri = ReleaseInfo(
            tag=tag,
            published_at=r.get("published_at"),
            asset=_select_asset_for_current_platform(tag, r.get("assets", []) or []),
            prerelease=r.get("prerelease"),
        )
        items.append(ri)

    # Sort by published_at desc as primary order (GitHub already returns latest first)
    items.sort(key=lambda ri: (ri.published_at or ""), reverse=True)
    logger.info(f"Releases fetched: {len(items)}")
    return items[:limit]


def is_newer(tag_remote: str, tag_local: str) -> bool:
    return compare_tags(tag_remote, tag_local) > 0


def stage_update_zip(asset_url: str, progress_cb=None, should_cancel=None) -> tuple[str, str]:
    """Download zip to temp and extract; return (zip_path, extracted_root)."""
    tmp_dir = tempfile.mkdtemp(prefix="rns_update_")
    zip_path = os.path.join(tmp_dir, "update.zip")
    _http_download(asset_url, zip_path, progress_cb=progress_cb, should_cancel=should_cancel)

    extracted_root = os.path.join(tmp_dir, "extracted")
    os.makedirs(extracted_root, exist_ok=True)
    logger.info(f"Extracting archive to: {extracted_root}")
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(extracted_root)
        names = zf.namelist()

    # Determine top-level folder inside zip (zip uses '/' as separator)
    def split_zip_path(p: str) -> list[str]:
        return [part for part in p.replace("\\", "/").strip("/").split("/") if part]

    tops: set[str] = set()
    for name in names:
        parts = split_zip_path(name)
        if parts:
            tops.add(parts[0])

    top_dir = None
    # Prefer bundle/exe container if present
    for name in names:
        parts = split_zip_path(name)
        if not parts:
            continue
        if parts[-1].lower() == "rnsapp.exe" and len(parts) >= 2:
            cand = os.path.join(extracted_root, *parts[:-1])
            if os.path.isdir(cand):
                top_dir = cand
                break
        if parts[0].lower().endswith(".app"):
            cand = os.path.join(extracted_root, parts[0])
            if os.path.isdir(cand):
                top_dir = cand
                break

    if not top_dir:
        if "RnSApp" in tops:
            cand = os.path.join(extracted_root, "RnSApp")
            top_dir = cand if os.path.isdir(cand) else None
        if not top_dir and len(tops) == 1:
            only = next(iter(tops))
            cand = os.path.join(extracted_root, only)
            top_dir = cand if os.path.isdir(cand) else None
    if not top_dir:
        # Fallback: if contents extracted flat or cannot determine
        top_dir = extracted_root
    logger.info(f"Update staged. zip={zip_path}, top={top_dir}")
    return zip_path, top_dir


def write_windows_updater_script(src_dir: str, dst_dir: str, exe_name: str = "RnSApp.exe") -> str:
    """Create a .bat script that waits for app to exit, swaps directories, and relaunches."""
    script = f"""
@echo off
setlocal ENABLEDELAYEDEXPANSION
set "SRC={src_dir}"
set "DST={dst_dir}"
set "BACKUP=%DST%_old_%RANDOM%"

echo Starting updater...
rem Wait for running app to exit
for /l %%i in (1,1,60) do (
  >nul 2>&1 (move /Y "%DST%" "%BACKUP%") && goto :moved
  timeout /t 2 /nobreak >nul
)
echo Failed to move old version. Opening folders for manual update.
start "" explorer "%SRC%"
start "" explorer "%DST%"
exit /b 1

:moved
echo Old version moved to %BACKUP%
>nul 2>&1 (move /Y "%SRC%" "%DST%")
if errorlevel 1 (
  echo Move failed; trying copy with robocopy
  robocopy "%SRC%" "%DST%" /MIR /R:5 /W:2 >nul
  if errorlevel 8 (
    echo Copy failed. Opening folders for manual update.
    start "" explorer "%SRC%"
    start "" explorer "%DST%"
    exit /b 1
  )
)

if exist "%DST%\{exe_name}" (
  start "" "%DST%\{exe_name}"
)
rem Cleanup backup
rmdir /s /q "%BACKUP%" >nul 2>&1
endlocal
exit /b 0
""".strip()
    tmp = tempfile.mkdtemp(prefix="rns_updater_")
    bat_path = os.path.join(tmp, "update.bat")
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(script)
    return bat_path


def launch_updater_and_quit(new_dir: str):
    if platform.system() != "Windows":
        raise RuntimeError("Auto-update launch is only implemented for Windows.")
    dst = _app_dir()
    bat = write_windows_updater_script(src_dir=new_dir, dst_dir=dst)
    # Launch detached (ensure script path is quoted in case of spaces)
    comspec = os.environ.get("COMSPEC", r"C:\\Windows\\System32\\cmd.exe")
    try:
        os.spawnl(os.P_NOWAIT, comspec, "cmd.exe", "/c", f'"{bat}"')
    except Exception:
        # Fallback: try without quoting
        os.spawnl(os.P_NOWAIT, comspec, "cmd.exe", "/c", bat)
    # Give a moment for the batch to start
    time.sleep(0.5)
    # Exit current app
    os._exit(0)


def get_app_dir() -> str:
    """Public accessor for current application directory."""
    return _app_dir()


def _mac_bundle_root_from_executable(exe_path: str) -> Optional[str]:
    # PyInstaller executable is inside RnSApp.app/Contents/MacOS/RnSApp
    p = os.path.abspath(exe_path)
    # ascend to .app
    for _ in range(3):
        p = os.path.dirname(p)
    if p.endswith(".app") and os.path.isdir(p):
        return p
    # Fallback: search upwards for *.app
    cur = os.path.abspath(exe_path)
    while cur and cur != "/":
        if cur.endswith(".app") and os.path.isdir(cur):
            return cur
        cur = os.path.dirname(cur)
    return None


def write_macos_updater_script(src_app: str, dst_app: str) -> str:
    script = f"""#!/bin/bash
set -e
SRC="{src_app}"
DST="{dst_app}"
BACKUP="${{DST}}_old_$$"

echo "[Updater] Starting macOS updater..."
# Ensure paths exist
if [ ! -e "$SRC" ]; then
  echo "Source app not found: $SRC" >&2
  open "$SRC" 2>/dev/null || true
  open -R "$DST" 2>/dev/null || true
  exit 1
fi

# Small delay to let the main app exit
sleep 1

set +e
mv "$DST" "$BACKUP" 2>/dev/null
MV1=$?
set -e

if [ $MV1 -ne 0 ]; then
  echo "[Updater] Move old app failed; trying rsync copy into DST"
  mkdir -p "$DST"
  if command -v rsync >/dev/null 2>&1; then
    rsync -a --delete "$SRC/" "$DST/" || RSYNC_FAIL=1
  else
    cp -R "$SRC/" "$DST/" || CP_FAIL=1
  fi
  if [ "$RSYNC_FAIL$CP_FAIL" != "" ]; then
    echo "[Updater] Copy failed. Opening folders for manual update."
    open "$SRC" 2>/dev/null || true
    open -R "$DST" 2>/dev/null || true
    exit 1
  fi
else
  set +e
  mv "$SRC" "$DST"
  MV2=$?
  set -e
  if [ $MV2 -ne 0 ]; then
    echo "[Updater] Move new app failed; trying rsync"
    if command -v rsync >/dev/null 2>&1; then
      rsync -a --delete "$SRC/" "$DST/" || RSYNC_FAIL=1
    else
      cp -R "$SRC/" "$DST/" || CP_FAIL=1
    fi
    if [ "$RSYNC_FAIL$CP_FAIL" != "" ]; then
      echo "[Updater] Copy failed. Opening folders for manual update."
      open "$SRC" 2>/dev/null || true
      open -R "$DST" 2>/dev/null || true
      exit 1
    fi
  fi
  rm -rf "$BACKUP" >/dev/null 2>&1 || true
fi

# Remove quarantine attribute if present
xattr -dr com.apple.quarantine "$DST" >/dev/null 2>&1 || true

# Relaunch
open "$DST" || true
exit 0
""".strip()
    tmp = tempfile.mkdtemp(prefix="rns_updater_mac_")
    sh_path = os.path.join(tmp, "update.sh")
    with open(sh_path, "w", encoding="utf-8") as f:
        f.write(script)
    os.chmod(sh_path, 0o755)
    return sh_path


def launch_macos_updater_and_quit(new_app_path: str):
    if platform.system() != "Darwin":
        raise RuntimeError("macOS updater can be launched only on macOS.")
    # Determine current bundle root
    if _is_frozen():
        dst = _mac_bundle_root_from_executable(sys.executable)
    else:
        # Dev mode: cannot auto-replace; open folders for manual copy
        dst = None
    if not dst:
        # Open folders to help manual update
        try:
            if os.path.isdir(new_app_path):
                os.spawnlp(os.P_NOWAIT, "open", "open", new_app_path)
        except Exception:
            pass
        return

    sh_path = write_macos_updater_script(src_app=new_app_path, dst_app=dst)
    # Launch detached
    try:
        os.spawnl(os.P_NOWAIT, "/bin/sh", "sh", sh_path)
    except Exception:
        # Fallback manual
        try:
            os.spawnlp(os.P_NOWAIT, "open", "open", new_app_path)
            os.spawnlp(os.P_NOWAIT, "open", "open", dst)
        except Exception:
            pass
    # Exit current app
    os._exit(0)
