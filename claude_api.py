"""
Claude.ai API client for fetching usage data.

Authentication: Uses sessionKey cookie from claude.ai browser session.
HTTP client: Uses curl_cffi to impersonate Chrome's TLS fingerprint,
             which is necessary to bypass Cloudflare protection on claude.ai.

API endpoints used:
  - GET /api/organizations          → list user's organizations
  - GET /api/organizations/{id}/usage → usage data (5h, 7d limits)

Usage response format:
  {
    "five_hour":      {"utilization": 0.0-1.0, "resets_at": "ISO8601"},
    "seven_day":      {"utilization": 0.0-1.0, "resets_at": "ISO8601"},
    "seven_day_opus": {"utilization": 0.0-1.0, "resets_at": "ISO8601"},
    "seven_day_sonnet": {"utilization": 0.0-1.0, "resets_at": "ISO8601"},
  }
"""

from __future__ import annotations

from curl_cffi import requests
from datetime import datetime, timezone


API_BASE = "https://claude.ai/api"


class ClaudeAPIError(Exception):
    """Raised when an API call fails (auth error, network issue, etc.)."""
    pass


class ClaudeAPI:
    """
    Client for the claude.ai web API.

    Args:
        session_key: The sessionKey cookie value from an authenticated
                     claude.ai browser session (starts with "sk-ant-sid").
    """

    def __init__(self, session_key: str):
        self.session_key = session_key

        # curl_cffi impersonates Chrome's TLS fingerprint to bypass Cloudflare.
        # Standard Python `requests` gets blocked with a 403 challenge page.
        # Try the latest Chrome fingerprint, fall back to generic "chrome"
        # if the specific version isn't available in the installed curl_cffi.
        for fp in ("chrome130", "chrome124", "chrome120", "chrome"):
            try:
                self.session = requests.Session(impersonate=fp)
                break
            except Exception:
                continue
        self.session.headers.update({
            "Accept": "application/json",
            "Content-Type": "application/json",
        })
        self.session.cookies.set("sessionKey", session_key, domain="claude.ai")

    def _check_response(self, resp, action: str):
        """Validate HTTP response and raise ClaudeAPIError on failure."""
        if resp.status_code in (401, 403):
            try:
                data = resp.json()
                msg = data.get("error", {}).get("message", "")
            except Exception:
                msg = ""
            if "invalid" in msg.lower() or "session" in msg.lower():
                raise ClaudeAPIError("Invalid sessionKey - please update your key")
            raise ClaudeAPIError(f"{action}: HTTP {resp.status_code} - {msg or resp.text[:100]}")
        if resp.status_code != 200:
            raise ClaudeAPIError(f"{action}: HTTP {resp.status_code}")

        # Cloudflare may return 200 with an HTML challenge page instead of JSON.
        # Detect this by checking Content-Type header.
        ct = resp.headers.get("content-type", "")
        if "application/json" not in ct:
            raise ClaudeAPIError(
                f"{action}: Cloudflare blocked the request (got HTML instead of JSON). "
                "Try updating curl_cffi: pip install -U curl_cffi"
            )

    def get_organizations(self) -> list[dict]:
        """
        Fetch the list of organizations for the authenticated user.

        Returns:
            List of org dicts, each containing at least "uuid" and "name".
        """
        resp = self.session.get(f"{API_BASE}/organizations", timeout=30)
        self._check_response(resp, "Fetch organizations")
        return resp.json()

    def get_usage(self, org_id: str) -> dict:
        """
        Fetch raw usage data for a specific organization.

        Args:
            org_id: Organization UUID.

        Returns:
            Raw JSON response dict with five_hour, seven_day, etc.
        """
        resp = self.session.get(f"{API_BASE}/organizations/{org_id}/usage", timeout=30)
        self._check_response(resp, "Fetch usage")
        return resp.json()

    def fetch_all(self, org_id: str) -> dict:
        """
        Fetch and parse usage into a structured result.

        Returns:
            Dict keyed by period name ("five_hour", "seven_day", etc.),
            each containing "label", "percentage" (0-100), and "reset_time".
        """
        raw = self.get_usage(org_id)
        return self._parse_usage(raw)

    # --- Parsing helpers ---

    @staticmethod
    def _parse_utilization(value) -> float:
        """
        Parse utilization value from API response.
        The API may return int, float, or string representation.
        Values are typically 0.0 to 1.0 (representing 0% to 100%).
        """
        if value is None:
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        try:
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    @staticmethod
    def _parse_reset_time(value) -> datetime | None:
        """
        Parse reset timestamp from API response.
        Handles both ISO 8601 strings and Unix timestamps.
        """
        if not value:
            return None
        try:
            if isinstance(value, (int, float)):
                return datetime.fromtimestamp(value, tz=timezone.utc)
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        except (ValueError, TypeError):
            return None

    def _parse_usage(self, raw: dict) -> dict:
        """
        Parse raw API response into structured usage data.

        Converts utilization (0.0-1.0) to percentage (0-100) and
        parses reset timestamps into datetime objects.
        """
        result = {}
        for key, label in [
            ("five_hour", "5h Session"),     # 5-hour rolling window limit
            ("seven_day", "7d Weekly"),       # 7-day overall limit
            ("seven_day_opus", "7d Opus"),    # 7-day Opus model specific limit
            ("seven_day_sonnet", "7d Sonnet"),# 7-day Sonnet model specific limit
        ]:
            period = raw.get(key, {})
            if period is None:
                period = {}
            utilization = self._parse_utilization(period.get("utilization"))
            reset_time = self._parse_reset_time(period.get("resets_at"))
            result[key] = {
                "label": label,
                # API returns 0.0-1.0; convert to percentage.
                # If value > 1.0 it's already a percentage (edge case).
                "percentage": round(utilization * 100, 1) if utilization <= 1.0 else round(utilization, 1),
                "reset_time": reset_time,
            }
        return result
