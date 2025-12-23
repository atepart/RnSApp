from __future__ import annotations

import json
import logging
import os
import platform
import re
import shutil
import subprocess
from dataclasses import dataclass
from typing import Optional

import requests


class UpdateError(RuntimeError):
    def __init__(
        self, message: str, *, url: str | None = None, status: int | None = None, body: str | None = None
    ) -> None:
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
    body: str | None = None


logger = logging.getLogger(__name__)

_TAG_RE = re.compile(r"^(?P<prefix>new)(?P<num>\d+)(?:b(?P<beta>\d+))?$")


def parse_tag(tag: str) -> tuple[int, int | None]:
    """Parse version tag like 'new103b1' into a tuple (103, 1)."""
    m = _TAG_RE.match(tag.strip())
    if not m:
        raise ValueError(f"Unsupported tag format: {tag}")
    num = int(m.group("num"))
    beta = m.group("beta")
    return num, (int(beta) if beta is not None else None)


def compare_tags(a: str, b: str) -> int:
    """Compare two tags by version semantics."""
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

    cmd = ["curl", "-sSL", "--connect-timeout", "5", "--max-time", "10"]
    for k, v in (headers or {}).items():
        cmd.extend(["-H", f"{k}: {v}"])
    cmd.extend(["-w", "\n%{http_code}\n", url])

    proc = subprocess.run(cmd, capture_output=True, text=True)
    stdout = proc.stdout or ""
    stderr = proc.stderr or ""
    if proc.returncode != 0:
        raise UpdateError("GitHub API error (curl)", url=url, body=stderr or stdout)

    body = stdout
    status_code: int | None = None
    if "\n" in stdout:
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


def _normalize_arch() -> str:
    m = platform.machine().lower()
    if m in ("x86_64", "amd64", "x64"):
        return "x64"
    if m in ("i386", "i686", "x86"):
        return "x86"
    if m in ("arm64", "aarch64"):
        return "arm64"
    return m or "unknown"


def _select_asset_for_current_platform(tag: str, assets: list[dict]) -> Optional[ReleaseAsset]:
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
        candidates.append(ReleaseAsset(name=name, download_url=url, size=a.get("size")))

    if not candidates:
        return None

    def score(asset: ReleaseAsset) -> tuple[int, int]:
        n = asset.name.lower()
        return (1 if tag.lower() in n else 0, len(n))

    candidates.sort(key=score, reverse=True)
    return candidates[0]


def find_latest_release(repo_slug: str) -> Optional[ReleaseInfo]:
    url = f"https://api.github.com/repos/{repo_slug}/releases"
    try:
        releases = _http_get_json(url)
    except Exception:
        return None
    if not isinstance(releases, list):
        return None

    prepared: list[tuple[tuple[int, int, int], dict]] = []
    for r in releases:
        tag = r.get("tag_name") or r.get("name") or ""
        if not _TAG_RE.match(tag):
            continue
        try:
            num, beta = parse_tag(tag)
        except ValueError:
            continue
        stable_flag = 1 if beta is None else 0
        beta_val = beta or 0
        prepared.append(((num, stable_flag, beta_val), r))

    if not prepared:
        return None

    prepared.sort(key=lambda it: (it[0][0], it[0][1], it[0][2], (it[1].get("published_at") or "")), reverse=True)
    latest = prepared[0][1]
    tag = latest.get("tag_name") or latest.get("name")
    published_at = latest.get("published_at")
    asset = _select_asset_for_current_platform(tag, latest.get("assets", []) or [])
    return ReleaseInfo(
        tag=tag,
        published_at=published_at,
        asset=asset,
        prerelease=latest.get("prerelease"),
        body=latest.get("body"),
    )


def list_releases(repo_slug: str, limit: int = 10) -> list[ReleaseInfo]:
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
            body=r.get("body"),
        )
        items.append(ri)

    items.sort(key=lambda ri: (ri.published_at or ""), reverse=True)
    logger.info(f"Releases fetched: {len(items)}")
    return items[:limit]


def is_newer(tag_remote: str, tag_local: str) -> bool:
    return compare_tags(tag_remote, tag_local) > 0
