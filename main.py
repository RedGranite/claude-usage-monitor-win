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
from tkinter import messagebox
from datetime import datetime, timezone
import sys
import ctypes
import ctypes.wintypes

from PIL import Image, ImageDraw, ImageFont
import pystray

from claude_api import ClaudeAPI, ClaudeAPIError
from webview_login import login_and_get_cookies
from config import load_config, save_config, CONFIG_DIR

from typing import Optional
import os
import json
import urllib.request
import webbrowser
import winreg

# ---------------------------------------------------------------------------
# Version info — update this when releasing a new version
# ---------------------------------------------------------------------------

APP_VERSION = "0.2"
GITHUB_REPO = "RedGranite/claude-usage-monitor-win"

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
# Auto-start (Windows registry)
# ---------------------------------------------------------------------------

_AUTOSTART_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
_AUTOSTART_NAME = "ClaudeUsageMonitor"


def _get_exe_path() -> str:
    """Return the path of the running executable (or script)."""
    if getattr(sys, "frozen", False):
        return sys.executable                        # PyInstaller exe
    return os.path.abspath(sys.argv[0])              # python script


def is_autostart_enabled() -> bool:
    """Check if auto-start registry entry exists and points to current exe."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _AUTOSTART_KEY) as key:
            val, _ = winreg.QueryValueEx(key, _AUTOSTART_NAME)
            return os.path.normcase(val.strip('"')) == os.path.normcase(_get_exe_path())
    except FileNotFoundError:
        return False
    except OSError:
        return False


def set_autostart(enable: bool):
    """Add or remove the auto-start registry entry."""
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, _AUTOSTART_KEY, 0, winreg.KEY_SET_VALUE
        ) as key:
            if enable:
                winreg.SetValueEx(key, _AUTOSTART_NAME, 0, winreg.REG_SZ,
                                  f'"{_get_exe_path()}"')
                log.info(f"Auto-start enabled: {_get_exe_path()}")
            else:
                try:
                    winreg.DeleteValue(key, _AUTOSTART_NAME)
                    log.info("Auto-start disabled")
                except FileNotFoundError:
                    pass
    except OSError as e:
        log.error(f"Failed to set auto-start: {e}")


# ---------------------------------------------------------------------------
# Classic balloon tip (Win32 Shell_NotifyIconW, forces legacy style)
# ---------------------------------------------------------------------------

class _NOTIFYICONDATAW(ctypes.Structure):
    """Win32 NOTIFYICONDATAW — passed to Shell_NotifyIconW."""
    _fields_ = [
        ("cbSize",           ctypes.wintypes.DWORD),
        ("hWnd",             ctypes.wintypes.HWND),
        ("uID",              ctypes.wintypes.UINT),
        ("uFlags",           ctypes.wintypes.UINT),
        ("uCallbackMessage", ctypes.wintypes.UINT),
        ("hIcon",            ctypes.wintypes.HICON),
        ("szTip",            ctypes.c_wchar * 128),
        ("dwState",          ctypes.wintypes.DWORD),
        ("dwStateMask",      ctypes.wintypes.DWORD),
        ("szInfo",           ctypes.c_wchar * 256),
        ("uVersion",         ctypes.wintypes.UINT),   # union with uTimeout
        ("szInfoTitle",      ctypes.c_wchar * 64),
        ("dwInfoFlags",      ctypes.wintypes.DWORD),
        ("guidItem",         ctypes.c_byte * 16),
        ("hBalloonIcon",     ctypes.wintypes.HICON),
    ]

_NIM_MODIFY     = 0x00000001
_NIM_SETVERSION = 0x00000004
_NIF_INFO       = 0x00000010
_NIIF_INFO      = 0x00000001
_NIIF_NONE      = 0x00000000
_NOTIFYICON_VERSION = 3          # version 3 = classic balloon (not toast)

_Shell_NotifyIconW = ctypes.windll.shell32.Shell_NotifyIconW


def _show_classic_balloon(icon: pystray.Icon, title: str, message: str):
    """
    Show a classic Windows balloon tip on the tray icon.

    Forces NOTIFYICON_VERSION=3 so the balloon renders in legacy style
    (small bubble next to tray icon) instead of Win10 toast notification.
    """
    try:
        hwnd = icon._hwnd
    except AttributeError:
        # Fallback: pystray's own notify (may show as toast)
        icon.notify(message, title)
        return

    nid = _NOTIFYICONDATAW()
    nid.cbSize = ctypes.sizeof(_NOTIFYICONDATAW)
    nid.hWnd = hwnd
    nid.uID = 0

    # Force legacy balloon version
    nid.uVersion = _NOTIFYICON_VERSION
    _Shell_NotifyIconW(_NIM_SETVERSION, ctypes.byref(nid))

    # Show the balloon
    nid.uFlags = _NIF_INFO
    nid.szInfoTitle = title[:63]
    nid.szInfo = message[:255]
    nid.dwInfoFlags = _NIIF_INFO
    _Shell_NotifyIconW(_NIM_MODIFY, ctypes.byref(nid))


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
        self._popup_window: Optional[tk.Tk] = None  # Reference to dashboard window
        self._popup_pinned = False               # Dashboard "always on top" state
        self._data_version = 0                   # Incremented on each successful refresh
        self._notified_brackets: dict[str, int] = {}  # Last notified 10% bracket per period
        self._click_time = 0.0                   # Timestamp of last left-click
        self._click_timer: Optional[threading.Timer] = None  # Single-click delay timer

        self._cookies = {}  # Browser cookies for API requests
        if self.config.get("session_key"):
            self._cookies = {"sessionKey": self.config["session_key"]}
            if self.config.get("cf_clearance"):
                self._cookies["cf_clearance"] = self.config["cf_clearance"]
            self.api = ClaudeAPI(self._cookies)

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
        return pystray.Menu(
            # default=True → fires on left-click, we route via _on_tray_click
            pystray.MenuItem("Show Dashboard", self._on_tray_click, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Refresh Now", self._on_refresh),
            pystray.MenuItem(
                "Auto-start",
                self._on_toggle_autostart,
                checked=lambda item: is_autostart_enabled(),
            ),
            pystray.MenuItem("Re-login...", self._on_set_key),
            pystray.MenuItem("Test Animations", self._on_test),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(f"v{APP_VERSION}", None, enabled=False),
            pystray.MenuItem("Quit", self._on_quit),
        )

    # -----------------------------------------------------------------------
    # Single-click / double-click dispatch
    # -----------------------------------------------------------------------

    def _on_tray_click(self, icon=None, item=None):
        """
        Dispatch left-click: single-click → balloon, double-click → dashboard.

        Uses a 400 ms timer to distinguish the two.
        """
        now = time.time()
        if self._click_timer:
            self._click_timer.cancel()
            self._click_timer = None

        if now - self._click_time < 0.4:
            # Double-click detected
            self._click_time = 0.0
            self._open_or_focus_dashboard()
        else:
            # First click — wait 400 ms to see if a second follows
            self._click_time = now
            self._click_timer = threading.Timer(0.4, self._show_usage_balloon)
            self._click_timer.start()

    def _show_usage_balloon(self):
        """Single-click action: show a classic balloon tip with usage summary."""
        if not self.icon:
            return
        if not self.usage:
            _show_classic_balloon(self.icon, "Claude Usage", "No data yet")
            return

        lines = []
        for key in ("five_hour", "seven_day"):
            data = self.usage.get(key)
            if data:
                lines.append(f"{data['label']}: {data['percentage']:.0f}%")
        _show_classic_balloon(self.icon, "Claude Usage", "\n".join(lines) or "No data")

    def _open_or_focus_dashboard(self):
        """Double-click action: open dashboard, or focus it if already open."""
        if self._popup_open and self._popup_window:
            try:
                self._popup_window.attributes("-topmost", True)
                self._popup_window.lift()
                self._popup_window.focus_force()
                # Restore original topmost state
                if not self._popup_pinned:
                    self._popup_window.after(200, lambda:
                        self._popup_window.attributes("-topmost", False)
                        if self._popup_window else None)
                return
            except tk.TclError:
                pass  # Window was destroyed
        if not self._popup_open:
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
          - Live-updating: content refreshes automatically when new data arrives
          - Flash animation on each data refresh
        """
        if self._popup_open:
            return
        self._popup_open = True

        BG = "#1E1E1E"
        FLASH_BG = "#1A2E1A"  # Subtle green tint for refresh flash

        popup = tk.Tk()
        self._popup_window = popup
        popup.title("Claude Usage")
        popup.configure(bg=BG)
        popup.overrideredirect(True)  # Remove native title bar
        popup.attributes("-topmost", self._popup_pinned)

        # Position near system tray (bottom-right corner of screen)
        sw, sh = popup.winfo_screenwidth(), popup.winfo_screenheight()
        w, h = 320, 240
        popup.geometry(f"{w}x{h}+{sw - w - 20}+{sh - h - 60}")

        # --- Custom title bar ---
        title_row = tk.Frame(popup, bg=BG)
        title_row.pack(fill="x", padx=10, pady=(8, 0))
        tk.Label(title_row, text="Claude Usage", font=("Segoe UI", 11, "bold"),
                 fg="#FFFFFF", bg=BG).pack(side="left")

        def on_close():
            self._popup_open = False
            self._popup_window = None
            popup.destroy()

        # Close button (red ✕)
        tk.Button(title_row, text="\u2715", command=on_close,
                  bg=BG, fg="#F44336", font=("Segoe UI", 10, "bold"),
                  relief="flat", bd=0, padx=4, pady=0,
                  activebackground="#F44336", activeforeground="#FFFFFF").pack(side="right")

        # Pin/unpin button (📌)
        def toggle_pin():
            self._popup_pinned = not self._popup_pinned
            popup.attributes("-topmost", self._popup_pinned)
            pin_btn.config(
                bg="#4CAF50" if self._popup_pinned else BG,
                fg="#FFFFFF" if self._popup_pinned else "#888888",
            )

        pin_btn = tk.Button(
            title_row, text="\U0001f4cc", command=toggle_pin,
            bg="#4CAF50" if self._popup_pinned else BG,
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

        # --- Separator ---
        tk.Frame(popup, bg="#333333", height=1).pack(fill="x", padx=10, pady=(6, 0))

        # --- Dynamic content area (rebuilt on each data refresh) ---
        content = tk.Frame(popup, bg=BG)
        content.pack(fill="both", expand=True)

        _WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

        def build_content(bg_color=BG):
            """Clear and rebuild all dynamic content with current data."""
            for child in content.winfo_children():
                child.destroy()
            content.configure(bg=bg_color)

            if self.config.get("org_name"):
                tk.Label(content, text=f"Org: {self.config['org_name']}",
                         font=("Segoe UI", 9), fg="#888888",
                         bg=bg_color).pack(anchor="w", padx=15, pady=(4, 0))

            if self.usage:
                frame = tk.Frame(content, bg=bg_color)
                frame.pack(padx=15, pady=6, fill="x")

                for key in ("five_hour", "seven_day"):
                    data = self.usage.get(key)
                    if not data:
                        continue
                    label = data["label"]
                    pct = data["percentage"]
                    reset = data.get("reset_time")
                    color = self._pct_color(pct)
                    color_dim = self._pct_color_dim(pct)

                    # Format reset time
                    reset_str = ""
                    if reset:
                        now = datetime.now(timezone.utc)
                        delta = reset - now
                        if delta.total_seconds() > 0:
                            if key == "seven_day":
                                local_reset = reset.astimezone()
                                day = _WEEKDAYS[local_reset.weekday()]
                                reset_str = f"resets {day} {local_reset.strftime('%H:%M')}"
                            else:
                                hours = int(delta.total_seconds() // 3600)
                                minutes = int((delta.total_seconds() % 3600) // 60)
                                reset_str = f"resets in {hours}h{minutes}m"

                    # Label + percentage row
                    row = tk.Frame(frame, bg=bg_color)
                    row.pack(fill="x", pady=(8, 2))
                    tk.Label(row, text=label, font=("Segoe UI", 10),
                             fg="#CCCCCC", bg=bg_color).pack(side="left")
                    tk.Label(row, text=f"{pct:.0f}%", font=("Segoe UI", 10, "bold"),
                             fg=color, bg=bg_color).pack(side="right")

                    # Progress bar
                    bar_h = 12
                    bar_w = 280
                    canvas = tk.Canvas(frame, width=bar_w, height=bar_h,
                                       bg="#2A2A2A", highlightthickness=0, bd=0)
                    canvas.pack(anchor="w")
                    filled_w = max(1, int(pct / 100 * bar_w)) if pct > 0 else 0
                    if filled_w > 0:
                        canvas.create_rectangle(0, 0, filled_w, bar_h,
                                                fill=color, outline="")
                    canvas.create_rectangle(filled_w, 0, bar_w, bar_h,
                                            fill=color_dim, outline="")

                    # Reset time label
                    if reset_str:
                        tk.Label(frame, text=reset_str, font=("Segoe UI", 11),
                                 fg="#AAAAAA", bg=bg_color).pack(anchor="e", pady=(4, 0))

            elif self.last_error:
                tk.Label(content, text=f"Error: {self.last_error}",
                         font=("Segoe UI", 9), fg="#F44336", bg=bg_color,
                         wraplength=280).pack(pady=10)
            else:
                tk.Label(content, text="No data yet...",
                         font=("Segoe UI", 10), fg="#888888",
                         bg=bg_color).pack(pady=20)

        # --- Flash animation on data refresh ---
        _flash = {"active": False}

        def do_flash():
            """Play a flash animation when data refreshes.

            Uses place() for the accent line so it overlays without
            pushing content down. Background tints green then restores.
            """
            if _flash["active"]:
                return
            _flash["active"] = True

            # Step 1: green-tinted background + overlay accent line
            build_content(bg_color=FLASH_BG)
            accent = tk.Frame(content, bg="#4CAF50", height=2)
            accent.place(x=0, y=0, relwidth=1.0)

            def step2():
                try:
                    accent.configure(height=3, bg="#66BB6A")
                except tk.TclError:
                    _flash["active"] = False
                    return
                popup.after(150, step3)

            def step3():
                try:
                    accent.destroy()
                    build_content(bg_color=BG)
                except tk.TclError:
                    pass
                _flash["active"] = False

            popup.after(200, step2)

        # --- Periodic check for data updates (every 1 second) ---
        last_ver = [self._data_version]

        def check_for_update():
            if not self._popup_open:
                return
            try:
                if self._data_version != last_ver[0]:
                    last_ver[0] = self._data_version
                    do_flash()
                popup.after(1000, check_for_update)
            except tk.TclError:
                pass  # Window was destroyed

        # --- Initial build + start update loop ---
        build_content()
        popup.after(1000, check_for_update)

        popup.protocol("WM_DELETE_WINDOW", on_close)
        popup.mainloop()

    # -----------------------------------------------------------------------
    # Menu actions
    # -----------------------------------------------------------------------

    def _on_refresh(self, icon=None, item=None):
        """Trigger an immediate usage data refresh."""
        threading.Thread(target=self._refresh_usage, daemon=True).start()

    def _on_toggle_autostart(self, icon=None, item=None):
        """Toggle auto-start on/off in the Windows registry."""
        set_autostart(not is_autostart_enabled())

    def _on_set_key(self, icon=None, item=None):
        """Re-login via webview (restarts tray icon)."""
        threading.Thread(target=self._relogin_from_tray, daemon=True).start()

    def _on_test(self, icon=None, item=None):
        """Run a visual test: simulate usage climbing 0% → 100% with notifications."""
        threading.Thread(target=self._run_test_sequence, daemon=True).start()

    def _run_test_sequence(self):
        """
        Simulate usage data going from 0% to 100% in 10% steps.

        At each step: update data → trigger icon + dashboard refresh → fire
        threshold notification. Pauses 2 s between steps so the user can see
        each animation and bubble.
        """
        log.info("=== TEST SEQUENCE START ===")
        saved_usage = self.usage
        saved_brackets = self._notified_brackets.copy()
        self._notified_brackets = {}

        from datetime import timedelta
        now = datetime.now(timezone.utc)

        for pct in range(0, 101, 10):
            self.usage = {
                "five_hour": {
                    "label": "5h Session",
                    "percentage": float(pct),
                    "reset_time": now + timedelta(hours=4, minutes=30),
                },
                "seven_day": {
                    "label": "7d Weekly",
                    "percentage": float(min(pct // 2, 100)),
                    "reset_time": now + timedelta(days=5),
                },
            }
            self._data_version += 1
            self._update_icon()
            self._check_thresholds()
            log.info(f"TEST step: 5h={pct}%")
            time.sleep(2)

        # Restore real data
        time.sleep(1)
        self.usage = saved_usage
        self._notified_brackets = saved_brackets
        self._data_version += 1
        self._update_icon()
        log.info("=== TEST SEQUENCE END ===")

    def _on_quit(self, icon=None, item=None):
        """Clean shutdown: stop threads, remove lock, exit tray."""
        self.running = False
        cleanup_lock()
        if self.icon:
            self.icon.stop()

    # -----------------------------------------------------------------------
    # Authentication management
    # -----------------------------------------------------------------------

    def _do_webview_login(self) -> bool:
        """
        Open an embedded browser window for the user to log into claude.ai.

        The real browser engine (Edge WebView2) handles Cloudflare automatically.
        Once logged in, cookies are extracted and saved.

        Returns:
            True if login succeeded, False if cancelled.
        """
        log.info("Opening webview login window...")
        cookies = login_and_get_cookies()

        if not cookies:
            log.info("Webview login cancelled or failed")
            return False

        log.info(f"Webview login got cookies: {list(cookies.keys())}")

        # Parse org data if the webview fetched it
        orgs_data = cookies.pop("_orgs_data", None)
        if orgs_data:
            try:
                orgs = json.loads(orgs_data)
                if orgs:
                    self.config["org_id"] = orgs[0].get("uuid", "")
                    self.config["org_name"] = orgs[0].get("name", "Unknown")
                    log.info(f"Org from webview: {self.config['org_name']}")
            except Exception as e:
                log.warning(f"Failed to parse orgs data: {e}")

        # Save sessionKey and cookies
        sk = cookies.get("sessionKey", "")
        cf = cookies.get("cf_clearance", "")
        if sk:
            self.config["session_key"] = sk
        if cf:
            self.config["cf_clearance"] = cf

        self._cookies = cookies
        self.api = ClaudeAPI(cookies)
        save_config(self.config)
        return True

    def _relogin_from_tray(self):
        """Re-login via embedded browser. Restarts the tray icon."""
        if self.icon:
            self.icon.stop()
        if self._do_webview_login():
            self._ensure_org()
            self._refresh_usage()
        # Re-create tray icon in a background thread
        self.icon = pystray.Icon(
            "claude_usage",
            self._create_icon(self._get_status_color(),
                              str(int(self.usage["five_hour"]["percentage"]))
                              if self.usage and "five_hour" in self.usage else "",
                              self._strip_on),
            "Claude Usage Monitor",
            menu=self._build_menu(),
        )
        self._update_icon()
        threading.Thread(target=self.icon.run, daemon=True).start()

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
            self._data_version += 1
            log.info(f"Usage refreshed (v{self._data_version}): {self.usage}")
        except ClaudeAPIError as e:
            log.error(f"API error: {e}")
            self.last_error = str(e)
        except Exception as e:
            log.error(f"Unexpected error: {e}", exc_info=True)
            self.last_error = f"Unexpected: {e}"
        self._update_icon()
        self._check_thresholds()

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

    def _check_thresholds(self):
        """
        Check if any usage crossed a 10% boundary and send a tray notification.

        Tracks per-period brackets so each boundary only fires once.
        A decrease (e.g. after reset) silently updates the tracker.
        """
        if not self.usage or not self.icon:
            return
        for key in ("five_hour", "seven_day"):
            data = self.usage.get(key)
            if not data:
                continue
            pct = data["percentage"]
            bracket = int(pct // 10) * 10          # 0, 10, 20, …, 100
            last = self._notified_brackets.get(key, -1)

            if last == -1:
                # First data — just record, don't notify
                self._notified_brackets[key] = bracket
            elif bracket > last:
                # Crossed upward — notify
                self._notified_brackets[key] = bracket
                label = data["label"]
                try:
                    _show_classic_balloon(
                        self.icon,
                        "Claude Usage Monitor",
                        f"{label} reached {pct:.0f}%",
                    )
                    log.info(f"Threshold notify: {label} → {pct:.0f}% (bracket {bracket})")
                except Exception as e:
                    log.debug(f"Notify failed: {e}")
            elif bracket < last:
                # Decreased (reset) — update silently
                self._notified_brackets[key] = bracket

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

    def _next_sleep(self) -> float:
        """
        Calculate how long to sleep before the next refresh.

        Normally uses the configured interval (default 5 min), but if a
        reset time is approaching sooner, wake up ~10 s after it so the
        dashboard updates promptly.
        """
        interval = self.config.get("refresh_interval", 300)
        if not self.usage:
            return interval

        now = datetime.now(timezone.utc)
        for key in ("five_hour", "seven_day"):
            reset = (self.usage.get(key) or {}).get("reset_time")
            if reset:
                secs = (reset - now).total_seconds()
                if secs <= 0:
                    # Already past reset — refresh soon
                    return 5
                if secs + 10 < interval:
                    # Reset coming before next scheduled refresh
                    interval = secs + 10
        return max(5, interval)

    def _refresh_loop(self):
        """Periodically refresh usage data, waking early when a reset is due."""
        while self.running:
            if self.api and self.config.get("org_id"):
                self._refresh_usage()
            sleep_secs = self._next_sleep()
            log.debug(f"Next refresh in {sleep_secs:.0f}s")
            # Sleep in small increments so we can stop quickly on quit
            elapsed = 0.0
            while elapsed < sleep_secs and self.running:
                time.sleep(min(5, sleep_secs - elapsed))
                elapsed += 5

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

    def _show_splash_tray(self):
        """
        Show a temporary gray tray icon immediately on startup so the user
        knows the app is loading.  Replaced by the real icon once data is ready.
        """
        self.icon = pystray.Icon(
            "claude_usage",
            self._create_icon("gray", "...", strip_on=False),
            "Claude Usage Monitor — loading...",
            menu=self._build_menu(),
        )
        # Run tray in a background thread so we can continue initialising
        threading.Thread(target=self.icon.run, daemon=True).start()
        # Give the tray icon a moment to appear
        time.sleep(0.3)

    def run(self):
        """
        Main entry point. Sequence:
          0. Show a loading tray icon instantly
          1. Prompt for session key if not configured
          2. Ensure organization is selected
          3. Perform initial usage data fetch
          4. Start background refresh + blink threads
          5. Swap to the real tray icon (blocks until quit)
        """
        log.info(f"=== Claude Usage Monitor v{APP_VERSION} starting ===")
        log.info(f"Config: session_key={'set' if self.config.get('session_key') else 'empty'}, org_id={self.config.get('org_id')}")
        log.info(f"Log file: {LOG_FILE}")

        # Step 0: Show loading tray icon right away
        self._show_splash_tray()

        # Step 0.5: Check for updates on GitHub
        check_for_update()

        # Step 1: If no API client yet, open embedded browser for login
        if not self.api:
            if not self._do_webview_login():
                log.info("No credentials provided, exiting.")
                if self.icon:
                    self.icon.stop()
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

        # Step 5: The splash tray is already running — just update it with
        #         real data and let the main thread wait for it to stop.
        log.info("Startup complete, tray icon ready.")
        self._update_icon()

        # Block the main thread until icon.stop() is called (by Quit)
        while self.running:
            time.sleep(1)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if not check_single_instance():
        sys.exit(0)
    try:
        monitor = UsageMonitor()
        monitor.run()
    finally:
        cleanup_lock()
