"""
Claude API client for fetching usage data.

Authentication: Uses cookies from the user's browser session on claude.ai.

Three HTTP backends with automatic fallback:
  1. Browser cookies + urllib  — uses ALL cookies (including cf_clearance) extracted
     from Chrome/Edge, so Cloudflare challenge is already solved.
  2. curl_cffi                 — Chrome TLS fingerprint impersonation (fast)
  3. PowerShell                — .NET SChannel, Windows native TLS

API endpoints:
  - GET /api/organizations            → list user's organizations
  - GET /api/organizations/{id}/usage → usage data (5h, 7d limits)
"""

from __future__ import annotations

import json
import logging
import subprocess
import tempfile
import os
import sys
import time
from datetime import datetime, timezone
from urllib import request as urllib_request
from urllib.error import HTTPError, URLError
from http.cookiejar import CookieJar, Cookie

log = logging.getLogger(__name__)

API_BASE = "https://claude.ai/api"


class ClaudeAPIError(Exception):
    """Raised when an API call fails (auth error, network issue, etc.)."""
    pass


# ---------------------------------------------------------------------------
# Browser cookie extraction (requires admin on Chrome/Edge v130+)
# ---------------------------------------------------------------------------

def extract_browser_cookies() -> dict:
    """
    Extract claude.ai cookies from the user's browser.

    First tries without admin. If app-bound encryption blocks it,
    launches a UAC-elevated helper to read cookies.

    Returns:
        Dict of cookie_name -> cookie_value, or empty dict on failure.
    """
    # Step 1: Try directly (works on older browsers or Firefox)
    cookies = _try_rookiepy_direct()
    if cookies and "sessionKey" in cookies:
        log.info(f"Cookies extracted directly from {cookies.get('_browser', 'browser')}")
        return cookies

    # Step 2: Try with admin elevation via UAC
    log.info("Direct extraction failed, trying UAC elevation...")
    cookies = _try_rookiepy_elevated()
    if cookies and "sessionKey" in cookies:
        log.info(f"Cookies extracted with admin from {cookies.get('_browser', 'browser')}")
        return cookies

    return {}


def _try_rookiepy_direct() -> dict:
    """Try to extract cookies without admin. May fail on Chrome/Edge v130+."""
    try:
        import rookiepy
    except ImportError:
        log.warning("rookiepy not installed")
        return {}

    browsers = [
        ("Edge", rookiepy.edge),
        ("Chrome", rookiepy.chrome),
        ("Firefox", rookiepy.firefox),
    ]
    for name, fn in browsers:
        try:
            raw = fn(["claude.ai"])
            cookies = {}
            for c in raw:
                if "claude.ai" in c.get("domain", ""):
                    cookies[c["name"]] = c["value"]
            if cookies and "sessionKey" in cookies:
                cookies["_browser"] = name
                return cookies
        except Exception as e:
            log.debug(f"{name}: {e}")
    return {}


def _try_rookiepy_elevated() -> dict:
    """
    Re-launch the exe/script as admin (UAC prompt) with --extract-cookies flag.

    The elevated process extracts cookies and writes them to a temp file.
    """
    import ctypes

    # Temp file for results
    tmp = os.path.join(tempfile.gettempdir(), "claude_cookies.json")
    # Clean up any stale file
    if os.path.exists(tmp):
        os.remove(tmp)

    # The exe itself handles --extract-cookies
    exe = sys.executable
    params = f'--extract-cookies "{tmp}"'
    log.info(f"Requesting admin elevation: {exe} {params}")

    ret = ctypes.windll.shell32.ShellExecuteW(
        None,           # hwnd
        "runas",        # operation — triggers UAC prompt
        exe,            # the same exe (or python.exe in dev)
        params,         # --extract-cookies <output>
        None,           # working directory
        0,              # SW_HIDE
    )

    # ShellExecuteW returns > 32 on success
    if ret <= 32:
        log.warning(f"UAC elevation failed or was cancelled (return={ret})")
        return {}

    # Wait for the helper to finish and write the file
    for _ in range(30):  # Wait up to 15 seconds
        time.sleep(0.5)
        if os.path.exists(tmp):
            try:
                with open(tmp, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                if content:
                    cookies = json.loads(content)
                    os.remove(tmp)
                    if "error" in cookies:
                        log.warning(f"Helper error: {cookies['error']}")
                        return {}
                    return cookies
            except (json.JSONDecodeError, OSError):
                continue

    log.warning("Timed out waiting for cookie_helper.py")
    return {}


# ---------------------------------------------------------------------------
# HTTP backend 1: urllib with full browser cookies
# ---------------------------------------------------------------------------

def _urllib_request_with_cookies(url: str, cookies: dict, timeout: int = 30) -> dict:
    """
    Make a request using urllib with ALL browser cookies.

    Since we include cf_clearance and other Cloudflare cookies,
    the request passes Cloudflare without needing TLS impersonation.
    """
    # Build cookie jar from extracted cookies
    jar = CookieJar()
    for name, value in cookies.items():
        if name.startswith("_"):
            continue  # Skip metadata keys like _browser
        c = Cookie(
            version=0, name=name, value=value,
            port=None, port_specified=False,
            domain=".claude.ai", domain_specified=True, domain_initial_dot=True,
            path="/", path_specified=True,
            secure=True, expires=None, discard=True,
            comment=None, comment_url=None, rest={}, rfc2109=False,
        )
        jar.set_cookie(c)

    opener = urllib_request.build_opener(urllib_request.HTTPCookieProcessor(jar))
    req = urllib_request.Request(url, headers={
        "Accept": "application/json",
        "Accept-Language": "en-US,en;q=0.9",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
        "Referer": "https://claude.ai/",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
    })

    try:
        with opener.open(req, timeout=timeout) as resp:
            ct = resp.headers.get("Content-Type", "")
            body = resp.read().decode("utf-8")
    except HTTPError as e:
        if e.code in (401, 403):
            raise ClaudeAPIError("Invalid sessionKey or expired cookies — please refresh")
        raise ClaudeAPIError(f"HTTP {e.code}")
    except URLError as e:
        raise ClaudeAPIError(f"Network error: {e.reason}")

    if "application/json" not in ct:
        if "just a moment" in body.lower() or "<html" in body.lower():
            raise ClaudeAPIError("Cloudflare blocked (cookies may be expired — restart app to refresh)")
        raise ClaudeAPIError(f"Unexpected response type: {ct}")

    return json.loads(body)


# ---------------------------------------------------------------------------
# HTTP backend 2: curl_cffi (Chrome TLS fingerprint)
# ---------------------------------------------------------------------------

def _curl_cffi_request(url: str, session_key: str, timeout: int = 30) -> dict:
    """Make request with curl_cffi Chrome impersonation."""
    from curl_cffi import requests

    for fp in ("chrome136", "chrome131", "chrome124", "chrome120", "chrome"):
        try:
            session = requests.Session(impersonate=fp)
            break
        except Exception:
            continue
    else:
        raise ClaudeAPIError("curl_cffi: no Chrome fingerprint available")

    session.headers.update({
        "Accept": "application/json",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://claude.ai/",
        "Origin": "https://claude.ai",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
    })
    session.cookies.set("sessionKey", session_key, domain="claude.ai")

    try:
        session.get("https://claude.ai/", timeout=10)
    except Exception:
        pass

    resp = session.get(url, timeout=timeout)
    if resp.status_code in (401, 403):
        raise ClaudeAPIError(f"HTTP {resp.status_code}")
    if resp.status_code != 200:
        raise ClaudeAPIError(f"HTTP {resp.status_code}")

    ct = resp.headers.get("content-type", "")
    if "application/json" not in ct:
        raise ClaudeAPIError("Cloudflare blocked (curl_cffi)")

    return resp.json()


# ---------------------------------------------------------------------------
# HTTP backend 3: PowerShell (.NET SChannel)
# ---------------------------------------------------------------------------

def _powershell_request(url: str, session_key: str, timeout: int = 30) -> dict:
    """Make request using PowerShell with Windows native TLS."""
    ps_script = f'''
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$cookie = New-Object System.Net.Cookie('sessionKey', '{session_key}', '/', 'claude.ai')
$session.Cookies.Add($cookie)
$session.UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36'
$headers = @{{ 'Accept'='application/json'; 'Referer'='https://claude.ai/'; 'Sec-Fetch-Dest'='empty'; 'Sec-Fetch-Mode'='cors'; 'Sec-Fetch-Site'='same-origin' }}
$resp = Invoke-WebRequest -Uri '{url}' -WebSession $session -Headers $headers -UseBasicParsing -TimeoutSec {timeout}
if ($resp.Headers['Content-Type'] -notlike '*application/json*') {{ Write-Error 'CLOUDFLARE_BLOCKED'; exit 1 }}
Write-Output $resp.Content
'''
    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-NonInteractive', '-Command', ps_script],
            capture_output=True, text=True, timeout=timeout + 10,
            creationflags=0x08000000,
        )
    except subprocess.TimeoutExpired:
        raise ClaudeAPIError("Request timed out")
    except FileNotFoundError:
        raise ClaudeAPIError("PowerShell not found")

    if result.returncode != 0 or not result.stdout.strip():
        stderr = result.stderr.strip()
        if "CLOUDFLARE" in stderr or "just a moment" in stderr.lower():
            raise ClaudeAPIError("Cloudflare blocked (PowerShell)")
        raise ClaudeAPIError(f"HTTP request failed: {stderr[:200]}")

    try:
        return json.loads(result.stdout.strip())
    except json.JSONDecodeError:
        raise ClaudeAPIError("Invalid JSON response")


# ---------------------------------------------------------------------------
# Main API client — tries browser cookies first, then curl_cffi, then PS
# ---------------------------------------------------------------------------

class ClaudeAPI:
    """
    Client for claude.ai API with triple-fallback HTTP backends.

    Priority:
      1. Browser cookies + urllib (most reliable — uses cf_clearance)
      2. curl_cffi (fast, no admin needed, may be blocked by Cloudflare)
      3. PowerShell (Windows native TLS, may be blocked by JS challenge)
    """

    def __init__(self, session_key: str, browser_cookies: dict = None):
        self.session_key = session_key
        self.browser_cookies = browser_cookies or {}
        self._backend = None  # Detected working backend

    def _request(self, url: str) -> dict:
        """Make a GET request with automatic backend selection."""

        # If we already know what works, use it
        if self._backend == "browser_cookies":
            return _urllib_request_with_cookies(url, self.browser_cookies)
        if self._backend == "curl_cffi":
            return _curl_cffi_request(url, self.session_key)
        if self._backend == "powershell":
            return _powershell_request(url, self.session_key)

        # Auto-detect: try each backend
        errors = []

        # Backend 1: browser cookies (if available)
        if self.browser_cookies and "sessionKey" in self.browser_cookies:
            try:
                result = _urllib_request_with_cookies(url, self.browser_cookies)
                self._backend = "browser_cookies"
                log.info("Backend: browser cookies + urllib")
                return result
            except ClaudeAPIError as e:
                errors.append(f"cookies: {e}")
                log.warning(f"Browser cookies backend failed: {e}")

        # Backend 2: curl_cffi
        try:
            result = _curl_cffi_request(url, self.session_key)
            self._backend = "curl_cffi"
            log.info("Backend: curl_cffi")
            return result
        except ClaudeAPIError as e:
            errors.append(f"curl_cffi: {e}")
            log.warning(f"curl_cffi backend failed: {e}")

        # Backend 3: PowerShell
        try:
            result = _powershell_request(url, self.session_key)
            self._backend = "powershell"
            log.info("Backend: PowerShell")
            return result
        except ClaudeAPIError as e:
            errors.append(f"powershell: {e}")

        raise ClaudeAPIError("All backends failed:\n" + "\n".join(errors))

    def get_organizations(self) -> list[dict]:
        return self._request(f"{API_BASE}/organizations")

    def get_usage(self, org_id: str) -> dict:
        return self._request(f"{API_BASE}/organizations/{org_id}/usage")

    def fetch_all(self, org_id: str) -> dict:
        raw = self.get_usage(org_id)
        return _parse_usage(raw)


# ---------------------------------------------------------------------------
# Parsing helpers
# ---------------------------------------------------------------------------

def _parse_utilization(value) -> float:
    if value is None:
        return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def _parse_reset_time(value) -> datetime | None:
    if not value:
        return None
    try:
        if isinstance(value, (int, float)):
            return datetime.fromtimestamp(value, tz=timezone.utc)
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except (ValueError, TypeError):
        return None


def _parse_usage(raw: dict) -> dict:
    result = {}
    for key, label in [
        ("five_hour", "5h Session"),
        ("seven_day", "7d Weekly"),
    ]:
        period = raw.get(key, {})
        if period is None:
            period = {}
        utilization = _parse_utilization(period.get("utilization"))
        reset_time = _parse_reset_time(period.get("resets_at"))
        result[key] = {
            "label": label,
            "percentage": round(utilization * 100, 1) if utilization <= 1.0 else round(utilization, 1),
            "reset_time": reset_time,
        }
    return result
