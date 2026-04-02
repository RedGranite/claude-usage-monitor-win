"""
Claude Usage Monitor — Windows system tray application.

Displays Claude AI usage limits (5h session, 7d weekly, model-specific)
in the Windows system tray with a color-coded icon and popup dashboard.

Architecture:
  - Main thread: runs pystray tray icon event loop
  - Refresh thread: polls Claude API every 5 minutes
  - Blink thread: animates the tray icon light strip every 1.5s
  - Popup thread: tkinter dashboard window (one at a time)

Files:
  - main.py       — This file. Tray icon, dashboard UI, app lifecycle.
  - claude_api.py  — Claude.ai API client (usage data fetching).
  - config.py      — Configuration management with DPAPI encryption.
"""

import logging
import threading
import time
import tkinter as tk
from tkinter import simpledialog, messagebox
from datetime import datetime, timezone
import sys
import ctypes

from PIL import Image, ImageDraw, ImageFont
import pystray

from claude_api import ClaudeAPI, ClaudeAPIError, extract_browser_cookies
from config import load_config, save_config, CONFIG_DIR

from typing import Optional
import os
import json
import urllib.request
import webbrowser

# ---------------------------------------------------------------------------
# Version info — update this when releasing a new version
# ---------------------------------------------------------------------------

APP_VERSION = "0.1"
GITHUB_REPO = "RedGranite/claude-usage-monitor"

# ---------------------------------------------------------------------------
# Logging — writes to %APPDATA%/ClaudeUsageMonitor/debug.log
# ---------------------------------------------------------------------------

os.makedirs(CONFIG_DIR, exist_ok=True)
LOG_FILE = os.path.join(CONFIG_DIR, "debug.log")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,  # DEBUG for verbose Pillow import logs
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Single instance enforcement
#
# Uses a PID lock file. On startup, checks if another instance is running
# and prompts the user to kill it or abort. The lock is cleaned up on exit.
# ---------------------------------------------------------------------------

LOCK_FILE = os.path.join(CONFIG_DIR, "instance.lock")


def check_for_update():
    """
    Check GitHub Releases for a newer version.
    Compares APP_VERSION with the latest release tag.
    If a newer version exists, prompt the user to open the download page.
    Runs silently on failure (no internet, GitHub down, etc.).
    """
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        req = urllib.request.Request(url, headers={"User-Agent": "ClaudeUsageMonitor"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode())
        latest_tag = data.get("tag_name", "").lstrip("v")
        if not latest_tag:
            return

        # Simple version comparison (works for x.y or x.y.z)
        def ver_tuple(v):
            return tuple(int(x) for x in v.split("."))

        if ver_tuple(latest_tag) > ver_tuple(APP_VERSION):
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            answer = messagebox.askyesno(
                "Claude Usage Monitor — Update Available",
                f"New version v{latest_tag} is available (current: v{APP_VERSION}).\n\n"
                "Open the download page?",
                parent=root,
            )
            root.destroy()
            if answer:
                webbrowser.open(data.get("html_url", f"https://github.com/{GITHUB_REPO}/releases"))
    except Exception:
        pass  # Network error, no internet, etc. — skip silently


def check_single_instance() -> bool:
    """
    Ensure only one instance of the application is running.

    If an existing instance is detected (via PID in lock file):
      - Prompts user to kill the old instance or cancel launch.

    Returns:
        True if this instance should proceed, False to abort.
    """
    if os.path.exists(LOCK_FILE):
        try:
            old_pid = int(open(LOCK_FILE).read().strip())
            # Check if process with that PID is still alive
            kernel32 = ctypes.windll.kernel32
            handle = kernel32.OpenProcess(0x1000, False, old_pid)  # PROCESS_QUERY_LIMITED_INFORMATION
            if handle:
                kernel32.CloseHandle(handle)
                # Old process is alive — ask user what to do
                root = tk.Tk()
                root.withdraw()
                root.attributes("-topmost", True)
                answer = messagebox.askyesno(
                    "Claude Usage Monitor",
                    f"Another instance is already running (PID {old_pid}).\n\n"
                    "Kill it and start a new one?",
                    parent=root,
                )
                root.destroy()
                if answer:
                    try:
                        kernel32.TerminateProcess(
                            kernel32.OpenProcess(0x0001, False, old_pid), 0
                        )
                        time.sleep(0.5)
                    except Exception:
                        pass
                else:
                    return False
        except (ValueError, OSError):
            pass  # Stale or unreadable lock file — safe to overwrite

    # Write our PID to the lock file
    with open(LOCK_FILE, "w") as f:
        f.write(str(os.getpid()))
    return True


def cleanup_lock():
    """Remove lock file if it belongs to the current process."""
    try:
        if os.path.exists(LOCK_FILE):
            pid = int(open(LOCK_FILE).read().strip())
            if pid == os.getpid():
                os.remove(LOCK_FILE)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Main application class
# ---------------------------------------------------------------------------


class UsageMonitor:
    """
    System tray application for monitoring Claude AI usage limits.

    The tray icon shows:
      - A large number: current 5h session usage percentage
      - A blinking colored strip: green (<60%), yellow (60-80%), red (>80%)

    Left-click opens a dark-themed dashboard popup with detailed usage bars.
    Right-click shows a context menu with refresh, key management, and quit.
    """

    def __init__(self):
        self.config = load_config()
        self.api: Optional[ClaudeAPI] = None
        self.usage: Optional[dict] = None        # Parsed usage data from API
        self.last_error: Optional[str] = None    # Last error message for UI display
        self.icon: Optional[pystray.Icon] = None
        self.running = True                      # Controls background thread lifecycle
        self._strip_on = True                    # Blink state for tray icon light strip
        self._popup_open = False                 # Prevents multiple dashboard windows
        self._popup_pinned = False               # Dashboard "always on top" state

        self._browser_cookies = {}
        if self.config.get("session_key"):
            # Restore cookies from config (including cf_clearance if saved)
            cookies = {"sessionKey": self.config["session_key"]}
            if self.config.get("cf_clearance"):
                cookies["cf_clearance"] = self.config["cf_clearance"]
            self._browser_cookies = cookies
            self.api = ClaudeAPI(self.config["session_key"], browser_cookies=cookies)

    # -----------------------------------------------------------------------
    # Tray icon rendering
    # -----------------------------------------------------------------------

    def _get_font(self, size: int):
        """Load Arial Bold (preferred) or fallback font at given size."""
        for name in ("arialbd.ttf", "arial.ttf"):
            try:
                return ImageFont.truetype(name, size)
            except OSError:
                continue
        return ImageFont.load_default()

    def _create_icon(self, color: str = "gray", text: str = "", strip_on: bool = True) -> Image.Image:
        """
        Generate a 64x64 RGBA tray icon.

        Layout:
          - Transparent background
          - Large white number with dark outline (for readability on any taskbar)
          - Colored light strip at bottom (blinks between bright and dim)

        Args:
            color: Status color name ("green", "yellow", "red", "gray").
            text: Text to render (typically the 5h usage percentage).
            strip_on: Whether the light strip is in bright (True) or dim (False) phase.
        """
        colors = {
            "green": (76, 175, 80),    # Material Green — usage < 60%
            "yellow": (255, 193, 7),   # Material Amber — usage 60-80%
            "red": (244, 67, 54),      # Material Red   — usage > 80%
            "gray": (158, 158, 158),   # No data / error
        }
        rgb = colors.get(color, colors["gray"])
        size = 64
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        if text:
            # Find the largest font size that fits within the icon width
            for font_size in (56, 50, 46, 40, 36):
                font = self._get_font(font_size)
                bbox = draw.textbbox((0, 0), text, font=font)
                tw = bbox[2] - bbox[0]
                if tw <= size - 2:
                    break
            bbox = draw.textbbox((0, 0), text, font=font)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
            x = (size - tw) // 2
            y = (size - th - 8) // 2 - bbox[1]  # Shift up to leave room for strip

            # Dark outline for visibility on both light and dark taskbars
            outline_color = (0, 0, 0, 220)
            for dx in (-2, -1, 0, 1, 2):
                for dy in (-2, -1, 0, 1, 2):
                    if dx == 0 and dy == 0:
                        continue
                    draw.text((x + dx, y + dy), text, fill=outline_color, font=font)
            draw.text((x, y), text, fill=(255, 255, 255, 255), font=font)

        # Blinking light strip at the bottom of the icon
        strip_y = size - 8
        strip_alpha = 255 if strip_on else 60  # Bright or dim
        strip_color = (*rgb, strip_alpha)
        draw.rounded_rectangle([4, strip_y, size - 5, size - 2], radius=3, fill=strip_color)

        return img

    def _get_status_color(self) -> str:
        """Determine tray icon color based on the highest usage percentage."""
        if not self.usage:
            return "gray"
        max_pct = 0.0
        for key in ("five_hour", "seven_day"):
            data = self.usage.get(key, {})
            pct = data.get("percentage", 0.0)
            if pct > max_pct:
                max_pct = pct
        if max_pct >= 80:
            return "red"
        elif max_pct >= 60:
            return "yellow"
        return "green"

    # -----------------------------------------------------------------------
    # Dashboard popup (dark-themed tkinter window)
    # -----------------------------------------------------------------------

    @staticmethod
    def _pct_color(pct: float) -> str:
        """Return hex color for a percentage value (bright, for text/bar fill)."""
        if pct >= 80:
            return "#F44336"  # Red
        elif pct >= 60:
            return "#FFC107"  # Amber
        return "#4CAF50"      # Green

    @staticmethod
    def _pct_color_dim(pct: float) -> str:
        """Return hex color for the unfilled portion of the progress bar."""
        if pct >= 80:
            return "#7A1A1A"  # Dark red
        elif pct >= 60:
            return "#7A5A00"  # Dark amber
        return "#1B5E20"      # Dark green

    def _build_menu(self) -> pystray.Menu:
        """Build the right-click context menu for the tray icon."""
        items = []
        # default=True makes this action trigger on left-click
        items.append(pystray.MenuItem("Show Dashboard", self._on_show_popup, default=True))
        items.append(pystray.Menu.SEPARATOR)
        items.append(pystray.MenuItem("Refresh Now", self._on_refresh))
        items.append(pystray.MenuItem("Set Session Key...", self._on_set_key))
        items.append(pystray.Menu.SEPARATOR)
        items.append(pystray.MenuItem(f"v{APP_VERSION}", None, enabled=False))
        items.append(pystray.MenuItem("Quit", self._on_quit))
        return pystray.Menu(*items)

    def _on_show_popup(self, icon=None, item=None):
        """Handle left-click / "Show Dashboard" — open popup if not already open."""
        if self._popup_open:
            return
        threading.Thread(target=self._show_usage_popup, daemon=True).start()

    def _show_usage_popup(self):
        """
        Create and display the usage dashboard popup window.

        Features:
          - No native title bar (overrideredirect); custom title with close/pin buttons
          - Draggable by the title area
          - Color-coded progress bars for each usage period
          - Pin button to keep window always-on-top
          - Only one instance allowed at a time (_popup_open flag)
        """
        if self._popup_open:
            return
        self._popup_open = True

        popup = tk.Tk()
        popup.title("Claude Usage")
        popup.configure(bg="#1E1E1E")
        popup.overrideredirect(True)  # Remove native title bar
        popup.attributes("-topmost", self._popup_pinned)

        # Position near system tray (bottom-right corner of screen)
        sw, sh = popup.winfo_screenwidth(), popup.winfo_screenheight()
        w, h = 320, 240
        popup.geometry(f"{w}x{h}+{sw - w - 20}+{sh - h - 60}")

        # --- Custom title bar ---
        title_row = tk.Frame(popup, bg="#1E1E1E")
        title_row.pack(fill="x", padx=10, pady=(8, 0))
        tk.Label(title_row, text="Claude Usage", font=("Segoe UI", 11, "bold"),
                 fg="#FFFFFF", bg="#1E1E1E").pack(side="left")

        def on_close():
            self._popup_open = False
            popup.destroy()

        # Close button (red ✕)
        tk.Button(title_row, text="\u2715", command=on_close,
                  bg="#1E1E1E", fg="#F44336", font=("Segoe UI", 10, "bold"),
                  relief="flat", bd=0, padx=4, pady=0,
                  activebackground="#F44336", activeforeground="#FFFFFF").pack(side="right")

        # Pin/unpin button (📌)
        def toggle_pin():
            self._popup_pinned = not self._popup_pinned
            popup.attributes("-topmost", self._popup_pinned)
            pin_btn.config(
                bg="#4CAF50" if self._popup_pinned else "#1E1E1E",
                fg="#FFFFFF" if self._popup_pinned else "#888888",
            )

        pin_btn = tk.Button(
            title_row, text="\U0001f4cc", command=toggle_pin,
            bg="#4CAF50" if self._popup_pinned else "#1E1E1E",
            fg="#FFFFFF" if self._popup_pinned else "#888888",
            font=("Segoe UI", 9), relief="flat", bd=0, padx=4, pady=0,
            activebackground="#4CAF50", activeforeground="#FFFFFF",
        )
        pin_btn.pack(side="right", padx=(0, 4))

        # --- Drag support (since there's no native title bar) ---
        drag_data = {"x": 0, "y": 0}

        def start_drag(e):
            drag_data["x"] = e.x
            drag_data["y"] = e.y

        def do_drag(e):
            x = popup.winfo_x() + e.x - drag_data["x"]
            y = popup.winfo_y() + e.y - drag_data["y"]
            popup.geometry(f"+{x}+{y}")

        title_row.bind("<Button-1>", start_drag)
        title_row.bind("<B1-Motion>", do_drag)

        # --- Content area ---
        tk.Frame(popup, bg="#333333", height=1).pack(fill="x", padx=10, pady=(6, 0))

        if self.config.get("org_name"):
            tk.Label(popup, text=f"Org: {self.config['org_name']}", font=("Segoe UI", 9),
                     fg="#888888", bg="#1E1E1E").pack(anchor="w", padx=15, pady=(4, 0))

        if self.usage:
            frame = tk.Frame(popup, bg="#1E1E1E")
            frame.pack(padx=15, pady=6, fill="x")

            # Weekday names for 7d reset display
            _WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

            for key in ("five_hour", "seven_day"):
                data = self.usage.get(key)
                if not data:
                    continue
                label = data["label"]
                pct = data["percentage"]
                reset = data.get("reset_time")
                color = self._pct_color(pct)
                color_dim = self._pct_color_dim(pct)

                # Format reset time based on period type
                reset_str = ""
                if reset:
                    now = datetime.now(timezone.utc)
                    delta = reset - now
                    if delta.total_seconds() > 0:
                        if key == "seven_day":
                            # 7d: show weekday + local time, e.g. "resets Wed 22:00"
                            local_reset = reset.astimezone()
                            day = _WEEKDAYS[local_reset.weekday()]
                            reset_str = f"resets {day} {local_reset.strftime('%H:%M')}"
                        else:
                            # 5h: show countdown, e.g. "resets in 2h11m"
                            hours = int(delta.total_seconds() // 3600)
                            minutes = int((delta.total_seconds() % 3600) // 60)
                            reset_str = f"resets in {hours}h{minutes}m"

                # Label + percentage row
                row = tk.Frame(frame, bg="#1E1E1E")
                row.pack(fill="x", pady=(8, 2))
                tk.Label(row, text=label, font=("Segoe UI", 10),
                         fg="#CCCCCC", bg="#1E1E1E").pack(side="left")
                tk.Label(row, text=f"{pct:.0f}%", font=("Segoe UI", 10, "bold"),
                         fg=color, bg="#1E1E1E").pack(side="right")

                # Color-coded progress bar (fixed 280px width)
                bar_h = 12
                bar_w = 280
                canvas = tk.Canvas(frame, width=bar_w, height=bar_h, bg="#2A2A2A",
                                   highlightthickness=0, bd=0)
                canvas.pack(anchor="w")
                filled_w = max(1, int(pct / 100 * bar_w)) if pct > 0 else 0
                if filled_w > 0:
                    canvas.create_rectangle(0, 0, filled_w, bar_h, fill=color, outline="")
                canvas.create_rectangle(filled_w, 0, bar_w, bar_h, fill=color_dim, outline="")

                # Reset time label (larger font)
                if reset_str:
                    tk.Label(frame, text=reset_str, font=("Segoe UI", 11),
                             fg="#AAAAAA", bg="#1E1E1E").pack(anchor="e", pady=(4, 0))

        elif self.last_error:
            tk.Label(popup, text=f"Error: {self.last_error}", font=("Segoe UI", 9),
                     fg="#F44336", bg="#1E1E1E", wraplength=280).pack(pady=10)
        else:
            tk.Label(popup, text="No data yet...", font=("Segoe UI", 10),
                     fg="#888888", bg="#1E1E1E").pack(pady=20)

        popup.protocol("WM_DELETE_WINDOW", on_close)
        popup.mainloop()

    # -----------------------------------------------------------------------
    # Menu actions
    # -----------------------------------------------------------------------

    def _on_refresh(self, icon=None, item=None):
        """Trigger an immediate usage data refresh."""
        threading.Thread(target=self._refresh_usage, daemon=True).start()

    def _on_set_key(self, icon=None, item=None):
        """Prompt user to enter a new session key (restarts tray icon)."""
        threading.Thread(target=self._prompt_key_from_tray, daemon=True).start()

    def _on_quit(self, icon=None, item=None):
        """Clean shutdown: stop threads, remove lock, exit tray."""
        self.running = False
        cleanup_lock()
        if self.icon:
            self.icon.stop()

    # -----------------------------------------------------------------------
    # Authentication management
    # -----------------------------------------------------------------------

    def _prompt_auth_sync(self) -> bool:
        """
        Show a custom dialog asking for sessionKey and cf_clearance.

        Both cookies are found in the same browser panel:
          F12 → Application → Cookies → https://claude.ai

        cf_clearance is needed to bypass Cloudflare JS challenge on
        machines where curl_cffi TLS impersonation doesn't work.

        Returns:
            True if credentials were entered and saved, False if cancelled.
        """
        result = {"ok": False}

        dialog = tk.Tk()
        dialog.title("Claude Usage Monitor")
        dialog.configure(bg="#1E1E1E")
        dialog.attributes("-topmost", True)
        dialog.resizable(False, False)

        # Center on screen
        dw, dh = 480, 340
        sx = (dialog.winfo_screenwidth() - dw) // 2
        sy = (dialog.winfo_screenheight() - dh) // 2
        dialog.geometry(f"{dw}x{dh}+{sx}+{sy}")

        # Instructions
        tk.Label(dialog, text="Enter cookies from claude.ai",
                 font=("Segoe UI", 12, "bold"), fg="#FFFFFF", bg="#1E1E1E"
                 ).pack(pady=(15, 5))
        tk.Label(dialog,
                 text="F12 → Application → Cookies → https://claude.ai",
                 font=("Segoe UI", 9), fg="#888888", bg="#1E1E1E"
                 ).pack(pady=(0, 10))

        # sessionKey field
        tk.Label(dialog, text="sessionKey (required):",
                 font=("Segoe UI", 10), fg="#CCCCCC", bg="#1E1E1E", anchor="w"
                 ).pack(fill="x", padx=30)
        entry_sk = tk.Entry(dialog, font=("Consolas", 10), width=50,
                            bg="#2A2A2A", fg="#FFFFFF", insertbackground="#FFFFFF",
                            relief="flat", bd=4)
        entry_sk.pack(padx=30, pady=(2, 10))

        # cf_clearance field
        tk.Label(dialog, text="cf_clearance (recommended — for Cloudflare bypass):",
                 font=("Segoe UI", 10), fg="#CCCCCC", bg="#1E1E1E", anchor="w"
                 ).pack(fill="x", padx=30)
        entry_cf = tk.Entry(dialog, font=("Consolas", 10), width=50,
                            bg="#2A2A2A", fg="#FFFFFF", insertbackground="#FFFFFF",
                            relief="flat", bd=4)
        entry_cf.pack(padx=30, pady=(2, 15))

        # Buttons
        btn_frame = tk.Frame(dialog, bg="#1E1E1E")
        btn_frame.pack(pady=(5, 15))

        def on_ok():
            result["ok"] = True
            result["session_key"] = entry_sk.get().strip()
            result["cf_clearance"] = entry_cf.get().strip()
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        tk.Button(btn_frame, text="OK", command=on_ok, width=10,
                  bg="#4CAF50", fg="#FFFFFF", font=("Segoe UI", 10, "bold"),
                  relief="flat", bd=0, activebackground="#388E3C"
                  ).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Cancel", command=on_cancel, width=10,
                  bg="#555555", fg="#FFFFFF", font=("Segoe UI", 10),
                  relief="flat", bd=0, activebackground="#777777"
                  ).pack(side="left", padx=5)

        entry_sk.focus_set()
        dialog.bind("<Return>", lambda e: on_ok())
        dialog.mainloop()

        if result["ok"] and result.get("session_key"):
            key = result["session_key"]
            cf = result.get("cf_clearance", "")
            log.info(f"Session key received (length={len(key)}), cf_clearance={'set' if cf else 'empty'}")

            self.config["session_key"] = key
            if cf:
                self.config["cf_clearance"] = cf

            # Build browser_cookies dict for the urllib backend
            cookies = {"sessionKey": key}
            if cf:
                cookies["cf_clearance"] = cf
            self._browser_cookies = cookies
            self.api = ClaudeAPI(key, browser_cookies=cookies)

            try:
                self._auto_select_org()
                save_config(self.config)
                log.info(f"Config saved. org_id={self.config.get('org_id')}")
            except Exception as e:
                log.error(f"Error after setting key: {e}", exc_info=True)
                self.last_error = str(e)
            return True
        return False

    def _prompt_key_from_tray(self, icon=None, item=None):
        """Re-prompt for session key while the app is running. Restarts the tray icon."""
        if self.icon:
            self.icon.stop()
        self._prompt_auth_sync()
        self._refresh_usage()
        self._start_tray()

    def _auto_select_org(self):
        """Fetch organizations and automatically select the first one."""
        if not self.api:
            return
        try:
            log.info("Fetching organizations...")
            orgs = self.api.get_organizations()
            log.info(f"Got {len(orgs)} org(s)")
            if orgs:
                org = orgs[0]
                self.config["org_id"] = org.get("uuid", "")
                self.config["org_name"] = org.get("name", "Unknown")
            else:
                log.warning("No organizations found")
                self.last_error = "No organizations found"
        except ClaudeAPIError as e:
            log.error(f"API error fetching orgs: {e}")
            self.last_error = str(e)
        except Exception as e:
            log.error(f"Unexpected error fetching orgs: {e}", exc_info=True)
            self.last_error = str(e)

    # -----------------------------------------------------------------------
    # Data refresh and icon update
    # -----------------------------------------------------------------------

    def _refresh_usage(self):
        """Fetch latest usage data from the API and update the tray icon."""
        if not self.api:
            log.debug("Skipping refresh: no API client")
            return
        # Session key mode needs org_id; OAuth mode does not
        if self.api and not self.config.get("org_id"):
            log.debug("Skipping refresh: session_key mode but no org_id")
            return
        try:
            log.info("Refreshing usage data...")
            self.usage = self.api.fetch_all(self.config.get("org_id", ""))
            self.last_error = None
            log.info(f"Usage refreshed: {self.usage}")
        except ClaudeAPIError as e:
            log.error(f"API error: {e}")
            self.last_error = str(e)
        except Exception as e:
            log.error(f"Unexpected error: {e}", exc_info=True)
            self.last_error = f"Unexpected: {e}"
        self._update_icon()

    def _update_icon(self):
        """Redraw the tray icon and tooltip with current usage data."""
        if self.icon:
            color = self._get_status_color()
            pct_text = ""
            if self.usage and "five_hour" in self.usage:
                pct = self.usage["five_hour"]["percentage"]
                pct_text = str(int(pct))
                self.icon.title = f"Claude Usage: {pct}% (5h)"
            elif self.last_error:
                pct_text = "!"
                self.icon.title = "Claude Usage: Error"
            else:
                self.icon.title = "Claude Usage Monitor"
            self.icon.icon = self._create_icon(color, pct_text, self._strip_on)
            self.icon.menu = self._build_menu()

    # -----------------------------------------------------------------------
    # Background threads
    # -----------------------------------------------------------------------

    def _blink_loop(self):
        """Toggle the tray icon light strip every 1.5 seconds for a breathing effect."""
        while self.running:
            time.sleep(1.5)
            self._strip_on = not self._strip_on
            if self.icon and self.usage:
                color = self._get_status_color()
                pct_text = ""
                if "five_hour" in self.usage:
                    pct_text = str(int(self.usage["five_hour"]["percentage"]))
                try:
                    self.icon.icon = self._create_icon(color, pct_text, self._strip_on)
                except Exception:
                    pass

    def _refresh_loop(self):
        """Periodically refresh usage data (default: every 5 minutes)."""
        while self.running:
            if self.api and self.config.get("org_id"):
                self._refresh_usage()
            time.sleep(self.config.get("refresh_interval", 300))

    # -----------------------------------------------------------------------
    # Application lifecycle
    # -----------------------------------------------------------------------

    def _start_tray(self):
        """Create and run the system tray icon. Blocks until icon.stop() is called."""
        color = self._get_status_color()
        pct_text = ""
        title = "Claude Usage Monitor"
        if self.usage and "five_hour" in self.usage:
            pct = self.usage["five_hour"]["percentage"]
            pct_text = str(int(pct))
            title = f"Claude Usage: {pct}% (5h)"
        self.icon = pystray.Icon(
            "claude_usage",
            self._create_icon(color, pct_text, self._strip_on),
            title,
            menu=self._build_menu(),
        )
        log.info("Starting tray icon...")
        self.icon.run()
        log.info("Tray icon stopped.")

    def _ensure_org(self):
        """Fetch organization info if we have a session key but no org_id."""
        # OAuth mode doesn't need org_id
        if self.api and not self.config.get("org_id"):
            log.info("org_id missing, fetching organizations...")
            self._auto_select_org()
            if self.config.get("org_id"):
                save_config(self.config)

    def run(self):
        """
        Main entry point. Sequence:
          1. Prompt for session key if not configured
          2. Ensure organization is selected
          3. Perform initial usage data fetch
          4. Start background refresh + blink threads
          5. Run tray icon (blocks until quit)
        """
        log.info(f"=== Claude Usage Monitor v{APP_VERSION} starting ===")
        log.info(f"Config: session_key={'set' if self.config.get('session_key') else 'empty'}, org_id={self.config.get('org_id')}")
        log.info(f"Log file: {LOG_FILE}")

        # Step 0: Check for updates on GitHub
        check_for_update()

        # Step 0.5: Try to extract cookies from browser (including cf_clearance)
        log.info("Extracting cookies from browser...")
        self._browser_cookies, cookie_err = extract_browser_cookies()
        if self._browser_cookies and "sessionKey" in self._browser_cookies:
            # Got cookies from browser — use them (no manual input needed!)
            key = self._browser_cookies["sessionKey"]
            self.config["session_key"] = key
            self.api = ClaudeAPI(key, browser_cookies=self._browser_cookies)
            self._auto_select_org()
            save_config(self.config)
            log.info(f"Auto-configured from browser cookies ({self._browser_cookies.get('_browser', '?')})")
        else:
            log.warning(f"Cookie extraction failed: {cookie_err}")

        # Step 1: If no API client yet, prompt for manual session key
        if not self.api:
            if not self._prompt_auth_sync():
                log.info("No credentials provided, exiting.")
                return

        # Step 2: Ensure we have an organization selected
        self._ensure_org()

        # Step 3: Initial data fetch
        self._refresh_usage()

        if self.last_error:
            log.warning(f"Initial refresh had error: {self.last_error}")

        # Step 4: Start background threads (daemon=True so they die with main thread)
        threading.Thread(target=self._refresh_loop, daemon=True).start()
        threading.Thread(target=self._blink_loop, daemon=True).start()

        # Step 5: Run tray icon (blocks main thread until quit)
        self._start_tray()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Special mode: extract cookies and exit (used by UAC elevation)
    if len(sys.argv) >= 3 and sys.argv[1] == "--extract-cookies":
        import json as _json
        _result = {}
        try:
            import rookiepy
            for _name, _fn in [("Edge", rookiepy.edge), ("Chrome", rookiepy.chrome), ("Firefox", rookiepy.firefox)]:
                try:
                    _raw = _fn(["claude.ai"])
                    _cookies = {}
                    for _c in _raw:
                        if "claude.ai" in _c.get("domain", ""):
                            _cookies[_c["name"]] = _c["value"]
                    if _cookies and "sessionKey" in _cookies:
                        _cookies["_browser"] = _name
                        _result = _cookies
                        break
                except Exception:
                    continue
        except Exception as _e:
            _result = {"error": str(_e)}
        with open(sys.argv[2], "w", encoding="utf-8") as _f:
            _json.dump(_result, _f)
        sys.exit(0)

    if not check_single_instance():
        sys.exit(0)
    try:
        monitor = UsageMonitor()
        monitor.run()
    finally:
        cleanup_lock()
