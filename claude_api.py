"""
Claude API client for fetching usage data.

Uses cookies obtained from the system Edge browser via CDP.
The real browser engine handles Cloudflare challenges automatically.

API endpoints:
  - GET /api/organizations            → list user's organizations
  - GET /api/organizations/{id}/usage → usage data (5h, 7d limits)
"""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from urllib import request as urllib_request
from urllib.error import HTTPError, URLError
from http.cookiejar import CookieJar, Cookie

log = logging.getLogger(__name__)

API_BASE = "https://claude.ai/api"


class ClaudeAPIError(Exception):
    """Raised when an API call fails."""
    pass


# ---------------------------------------------------------------------------
# HTTP request using browser cookies
# ---------------------------------------------------------------------------

def _make_request(url: str, cookies: dict, timeout: int = 30) -> dict:
    """
    Make an HTTP GET request with the given cookies.

    Since cookies come from a real browser session (including cf_clearance),
    Cloudflare will not block the request.
    """
    jar = CookieJar()
    for name, value in cookies.items():
        if name.startswith("_"):
            continue
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
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://claude.ai/",
    })

    try:
        with opener.open(req, timeout=timeout) as resp:
            body = resp.read().decode("utf-8")
            ct = resp.headers.get("Content-Type", "")
    except HTTPError as e:
        if e.code in (401, 403):
            raise ClaudeAPIError("Session expired — please re-login")
        raise ClaudeAPIError(f"HTTP {e.code}")
    except URLError as e:
        raise ClaudeAPIError(f"Network error: {e.reason}")

    if "application/json" not in ct:
        if "just a moment" in body.lower():
            raise ClaudeAPIError("Cloudflare blocked — cookies may be expired, please re-login")
        raise ClaudeAPIError(f"Unexpected response: {ct}")

    return json.loads(body)


class ClaudeAPI:
    """Client for claude.ai API using browser cookies."""

    def __init__(self, cookies: dict):
        self.cookies = cookies

    def get_organizations(self) -> list[dict]:
        return _make_request(f"{API_BASE}/organizations", self.cookies)

    def get_usage(self, org_id: str) -> dict:
        return _make_request(f"{API_BASE}/organizations/{org_id}/usage", self.cookies)

    def fetch_all(self, org_id: str) -> dict:
        raw = self.get_usage(org_id)
        return _parse_usage(raw)


# ---------------------------------------------------------------------------
# Parsing helpers
# ---------------------------------------------------------------------------

def _parse_usage(raw: dict) -> dict:
    result = {}
    for key, label in [("five_hour", "5h Session"), ("seven_day", "7d Weekly")]:
        period = raw.get(key) or {}
        util = period.get("utilization")
        util = float(util) if util is not None else 0.0
        reset = period.get("resets_at")
        reset_time = None
        if reset:
            try:
                if isinstance(reset, (int, float)):
                    reset_time = datetime.fromtimestamp(reset, tz=timezone.utc)
                else:
                    reset_time = datetime.fromisoformat(reset.replace("Z", "+00:00"))
            except (ValueError, TypeError):
                pass
        result[key] = {
            "label": label,
            "percentage": round(util * 100, 1) if util <= 1.0 else round(util, 1),
            "reset_time": reset_time,
        }
    return result
