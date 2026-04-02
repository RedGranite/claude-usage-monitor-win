"""
Cookie extraction helper — runs with admin privileges to read browser cookies.

Chrome/Edge v130+ uses app-bound encryption for cookies, which requires
admin access to decrypt. This script is launched via UAC elevation from
the main app, extracts claude.ai cookies, and writes them to a temp file.

Usage: python cookie_helper.py <output_path>
"""

import json
import sys
import os


def extract_cookies() -> dict:
    """
    Try to extract claude.ai cookies from installed browsers.

    Tries Edge, Chrome, Firefox in order. Returns a dict of
    cookie_name -> cookie_value for all claude.ai cookies found.
    """
    try:
        import rookiepy
    except ImportError:
        return {"error": "rookiepy not installed"}

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
                domain = c.get("domain", "")
                if "claude.ai" in domain:
                    cookies[c["name"]] = c["value"]
            if cookies and "sessionKey" in cookies:
                cookies["_browser"] = name
                return cookies
        except Exception:
            continue

    return {"error": "Could not extract cookies from any browser"}


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: cookie_helper.py <output_path>")
        sys.exit(1)

    output_path = sys.argv[1]
    cookies = extract_cookies()

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(cookies, f)
