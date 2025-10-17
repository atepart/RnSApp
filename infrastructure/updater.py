from __future__ import annotations

import json
import logging
import os
import platform
import re
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass
from typing import Optional

import requests


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
    hdrs = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "RnSApp-Updater/1.0",
    }
    token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    if token:
        hdrs["Authorization"] = f"Bearer {token}"
    return hdrs


def _http_get_json(url: str) -> list[dict] | dict:
    logger.debug(f"HTTP GET JSON: {url}")
    try:
        resp = requests.get(url, headers=_request_headers(), timeout=(5, 8))
        resp.raise_for_status()
    except requests.RequestException as e:
        logger.warning(f"GitHub API request failed: {e}")
        raise
    return resp.json()


def _http_download(url: str, dest_path: str, progress_cb=None):
    logger.info(f"Downloading asset: {url}")
    try:
        with requests.get(url, stream=True, timeout=(5, 30)) as r:
            r.raise_for_status()
            total = int(r.headers.get("Content-Length", "0") or 0)
            downloaded = 0
            chunk = 64 * 1024
            with open(dest_path, "wb") as f:
                for part in r.iter_content(chunk_size=chunk):
                    if not part:
                        continue
                    f.write(part)
                    downloaded += len(part)
                    if progress_cb:
                        try:
                            progress_cb(downloaded, total)
                        except Exception:
                            pass
    except requests.RequestException as e:
        logger.error(f"Download failed: {e}")
        raise


def _is_frozen() -> bool:
    return getattr(sys, "frozen", False) is True


def _app_dir() -> str:
    if _is_frozen():
        return os.path.dirname(sys.executable)
    # Dev mode: use project root main directory
    return os.path.abspath(os.path.dirname(sys.argv[0]))


def _platform_asset_suffix(tag: str) -> Optional[str]:
    system = platform.system()
    if system == "Windows":
        arch = "x86"  # current CI builds
        return f"Windows_{arch}_{tag}.zip"
    if system == "Darwin":
        m = platform.machine().lower()
        arch = "arm64" if m in ("arm64", "aarch64") else "x64"
        return f"macOS_{arch}_{tag}.zip"
    # For other platforms return None to indicate unsupported
    return None


def _select_asset_for_current_platform(tag: str, assets: list[dict]) -> Optional[ReleaseAsset]:
    suffix = _platform_asset_suffix(tag)
    if not suffix:
        return None
    for a in assets or []:
        if isinstance(a, dict) and a.get("name", "").endswith(suffix):
            return ReleaseAsset(
                name=a.get("name", ""),
                download_url=a.get("browser_download_url", ""),
                size=a.get("size"),
            )
    return None


def find_latest_release(repo_slug: str) -> Optional[ReleaseInfo]:
    """Fetch releases and return the latest by tag semantics (then date)."""
    url = f"https://api.github.com/repos/{repo_slug}/releases"
    try:
        releases = _http_get_json(url)
    except Exception:
        return None
    if not isinstance(releases, list):
        return None

    # Filter tags that match expected format
    filtered: list[dict] = []
    for r in releases:
        tag = r.get("tag_name") or r.get("name") or ""
        if _TAG_RE.match(tag):
            filtered.append(r)
    if not filtered:
        return None

    filtered.sort(key=lambda r: (parse_tag(r.get("tag_name") or r.get("name")), r.get("published_at") or ""))
    latest = filtered[-1]
    tag = latest.get("tag_name") or latest.get("name")
    published_at = latest.get("published_at")
    asset = _select_asset_for_current_platform(tag, latest.get("assets", []) or [])
    return ReleaseInfo(tag=tag, published_at=published_at, asset=asset, prerelease=latest.get("prerelease"))


def list_releases(repo_slug: str, limit: int = 10) -> list[ReleaseInfo]:
    url = f"https://api.github.com/repos/{repo_slug}/releases?per_page={max(1, min(limit, 100))}"
    logger.info(f"Fetching releases: {url}")
    try:
        releases = _http_get_json(url)
    except Exception as e:
        logger.warning(f"Failed to fetch releases: {e}")
        return []
    if not isinstance(releases, list):
        return []

    # Keep only tags matching our scheme
    items: list[ReleaseInfo] = []
    for r in releases:
        tag = r.get("tag_name") or r.get("name") or ""
        if not _TAG_RE.match(tag):
            continue
        ri = ReleaseInfo(
            tag=tag,
            published_at=r.get("published_at"),
            asset=_select_asset_for_current_platform(tag, r.get("assets", []) or []),
            prerelease=r.get("prerelease"),
        )
        items.append(ri)

    # Sort by semantic version then date
    def sort_key(ri: ReleaseInfo):
        num, beta = parse_tag(ri.tag)
        stable_flag = 1 if beta is None else 0  # stable first
        beta_val = beta or 0
        return (num, stable_flag, beta_val, ri.published_at or "")

    items.sort(key=sort_key, reverse=True)
    logger.debug(f"Total releases after filter: {len(items)}")
    return items[:limit]


def is_newer(tag_remote: str, tag_local: str) -> bool:
    return compare_tags(tag_remote, tag_local) > 0


def stage_update_zip(asset_url: str, progress_cb=None) -> tuple[str, str]:
    """Download zip to temp and extract; return (zip_path, extracted_root)."""
    tmp_dir = tempfile.mkdtemp(prefix="rns_update_")
    zip_path = os.path.join(tmp_dir, "update.zip")
    _http_download(asset_url, zip_path, progress_cb=progress_cb)

    extracted_root = os.path.join(tmp_dir, "extracted")
    os.makedirs(extracted_root, exist_ok=True)
    logger.info(f"Extracting archive to: {extracted_root}")
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(extracted_root)

    # Determine top-level folder inside zip
    top_candidates = [name.split(os.sep)[0] for name in zf.namelist() if name and not name.endswith("/")]
    top = top_candidates[0] if top_candidates else "RnSApp"
    top_dir = os.path.join(extracted_root, top)
    if not os.path.isdir(top_dir):
        # Fallback: if contents extracted flat
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
    # Launch detached
    os.spawnl(os.P_NOWAIT, os.environ.get("COMSPEC", "C:\\Windows\\System32\\cmd.exe"), "cmd", "/c", bat)
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
