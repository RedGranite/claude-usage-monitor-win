"""
System browser login for Claude Usage Monitor.

Opens a Chromium-based browser (Edge, Chrome, or Brave) with Chrome DevTools
Protocol (CDP) enabled. The user logs in to claude.ai normally, and cookies
are extracted via CDP Network.getCookies.

Zero external dependencies — uses only Python stdlib + system browser.
Supports: Microsoft Edge, Google Chrome, Brave Browser.
"""

import base64
import json
import logging
import os
import socket
import struct
import subprocess
import threading
import time
import tkinter as tk
import urllib.request
import winreg

log = logging.getLogger(__name__)

CDP_PORT = 9223  # Non-standard port to avoid conflicts


# ---------------------------------------------------------------------------
# Find a Chromium-based browser on the system
# ---------------------------------------------------------------------------


def _get_default_browser_exe() -> str | None:
    """Read the user's default browser exe path from the Windows registry."""
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice",
        ) as key:
            prog_id, _ = winreg.QueryValueEx(key, "ProgId")

        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT,
            rf"{prog_id}\shell\open\command",
        ) as key:
            cmd, _ = winreg.QueryValueEx(key, "")

        # '"C:\...\browser.exe" "%1"'  →  C:\...\browser.exe
        exe = cmd.split('"')[1] if cmd.startswith('"') else cmd.split()[0]
        if os.path.isfile(exe):
            return exe
    except Exception as e:
        log.debug(f"Registry read failed: {e}")
    return None


def _find_edge() -> str | None:
    """Locate Edge (always present on Windows 10/11) as fallback."""
    for env in ("ProgramFiles(x86)", "ProgramFiles", "LOCALAPPDATA"):
        base = os.environ.get(env, "")
        if base:
            p = os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe")
            if os.path.isfile(p):
                return p
    return None


def _try_cdp(proc: subprocess.Popen, timeout: int = 5) -> bool:
    """Check if a launched browser responds to CDP within *timeout* seconds."""
    for _ in range(timeout * 2):
        if proc.poll() is not None:
            return False
        try:
            with urllib.request.urlopen(
                f"http://localhost:{CDP_PORT}/json/version", timeout=1
            ) as r:
                data = json.loads(r.read().decode())
                if "webSocketDebuggerUrl" in data or "Browser" in data:
                    return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


# ---------------------------------------------------------------------------
# Minimal WebSocket client (just enough for CDP JSON-RPC)
# ---------------------------------------------------------------------------

def _ws_connect(url: str) -> socket.socket:
    """Open a WebSocket connection and return the raw socket."""
    rest = url[5:]  # strip "ws://"
    slash = rest.index("/")
    host_port, path = rest[:slash], rest[slash:]
    host, port_s = host_port.rsplit(":", 1)
    port = int(port_s)

    sock = socket.create_connection((host, port), timeout=10)
    key = base64.b64encode(os.urandom(16)).decode()
    handshake = (
        f"GET {path} HTTP/1.1\r\n"
        f"Host: {host}:{port}\r\n"
        f"Upgrade: websocket\r\n"
        f"Connection: Upgrade\r\n"
        f"Sec-WebSocket-Key: {key}\r\n"
        f"Sec-WebSocket-Version: 13\r\n\r\n"
    )
    sock.sendall(handshake.encode())

    buf = b""
    while b"\r\n\r\n" not in buf:
        chunk = sock.recv(4096)
        if not chunk:
            raise ConnectionError("Connection closed during handshake")
        buf += chunk
    if b"101" not in buf.split(b"\r\n")[0]:
        sock.close()
        raise ConnectionError("WebSocket handshake failed")
    return sock


def _ws_send(sock: socket.socket, text: str):
    """Send a masked WebSocket text frame."""
    data = text.encode()
    mask = os.urandom(4)
    hdr = bytearray([0x81])  # FIN + TEXT
    n = len(data)
    if n < 126:
        hdr.append(0x80 | n)
    elif n < 65536:
        hdr.append(0x80 | 126)
        hdr += struct.pack(">H", n)
    else:
        hdr.append(0x80 | 127)
        hdr += struct.pack(">Q", n)
    hdr += mask
    masked = bytearray(n)
    for i in range(n):
        masked[i] = data[i] ^ mask[i & 3]
    sock.sendall(bytes(hdr) + bytes(masked))


def _ws_recv(sock: socket.socket) -> str | None:
    """Receive one WebSocket frame. Returns text content or None on close."""
    def _read(n):
        buf = b""
        while len(buf) < n:
            chunk = sock.recv(n - len(buf))
            if not chunk:
                raise ConnectionError("Connection closed")
            buf += chunk
        return buf

    h = _read(2)
    op = h[0] & 0x0F
    length = h[1] & 0x7F
    if length == 126:
        length = struct.unpack(">H", _read(2))[0]
    elif length == 127:
        length = struct.unpack(">Q", _read(8))[0]
    if h[1] & 0x80:
        mk = _read(4)
        raw = bytearray(_read(length))
        for i in range(length):
            raw[i] ^= mk[i & 3]
        raw = bytes(raw)
    else:
        raw = _read(length)
    if op == 1:  # TEXT
        return raw.decode()
    if op == 8:  # CLOSE
        return None
    return raw.decode("utf-8", errors="replace")


# ---------------------------------------------------------------------------
# CDP cookie extraction
# ---------------------------------------------------------------------------

def _cdp_get_cookies(ws_url: str) -> dict:
    """Fetch claude.ai cookies via CDP Network.getCookies command."""
    sock = _ws_connect(ws_url)
    try:
        _ws_send(sock, json.dumps({
            "id": 1,
            "method": "Network.getCookies",
            "params": {"urls": ["https://claude.ai"]}
        }))
        for _ in range(30):
            msg = _ws_recv(sock)
            if msg is None:
                break
            resp = json.loads(msg)
            if resp.get("id") == 1:
                return {
                    c["name"]: c["value"]
                    for c in resp.get("result", {}).get("cookies", [])
                }
        return {}
    finally:
        sock.close()


def _get_ws_url() -> str | None:
    """Get the first page's WebSocket debugger URL from CDP."""
    try:
        with urllib.request.urlopen(
            f"http://localhost:{CDP_PORT}/json/list", timeout=2
        ) as r:
            pages = json.loads(r.read().decode())
        for p in pages:
            ws = p.get("webSocketDebuggerUrl")
            if ws:
                return ws
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def _launch_browser(exe: str, profile_dir: str) -> subprocess.Popen:
    """Launch a browser with CDP enabled and a dedicated profile."""
    return subprocess.Popen([
        exe,
        f"--remote-debugging-port={CDP_PORT}",
        f"--user-data-dir={profile_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        "https://claude.ai/login",
    ])


def login_and_get_cookies() -> dict:
    """
    Launch a browser for claude.ai login, extract cookies via CDP.

    Strategy:
      1. Try the user's default browser with CDP — works for any Chromium fork.
      2. If CDP doesn't respond (e.g. Firefox), kill it and fall back to Edge.

    Shows a small "waiting" dialog while the user logs in.

    Returns:
        dict of cookie name → value (including sessionKey), or empty dict.
    """
    from config import CONFIG_DIR
    profile_dir = os.path.join(CONFIG_DIR, "browser_profile")
    os.makedirs(profile_dir, exist_ok=True)

    # --- Step 1: try default browser ---
    default_exe = _get_default_browser_exe()
    proc = None
    browser_label = ""

    if default_exe:
        browser_label = os.path.basename(default_exe)
        log.info(f"Trying default browser: {browser_label}")
        proc = _launch_browser(default_exe, profile_dir)

        if _try_cdp(proc):
            log.info(f"Default browser ({browser_label}) supports CDP ✓")
        else:
            # Not Chromium — kill and fall back
            log.info(f"Default browser ({browser_label}) does not support CDP, falling back to Edge")
            proc.terminate()
            proc.wait(timeout=3)
            proc = None

    # --- Step 2: fall back to Edge ---
    if proc is None:
        edge = _find_edge()
        if not edge:
            log.error("No usable browser found (Edge not installed)")
            return {}
        browser_label = "msedge.exe"
        log.info(f"Launching Edge: {edge}")
        proc = _launch_browser(edge, profile_dir)
        if not _try_cdp(proc):
            log.error("Edge CDP also failed")
            proc.terminate()
            return {}

    log.info(f"Browser ready ({browser_label}), waiting for login...")

    # --- Poll for cookies in background ---
    result = {"cookies": {}, "done": False}

    def _poll():
        ws_url = None
        for _ in range(30):
            if proc.poll() is not None:
                result["done"] = True
                return
            ws_url = _get_ws_url()
            if ws_url:
                break
            time.sleep(1)

        if not ws_url:
            log.error("CDP page WebSocket not found")
            result["done"] = True
            return

        for _ in range(150):  # 5 min
            if proc.poll() is not None:
                try:
                    ws = _get_ws_url()
                    if ws:
                        cookies = _cdp_get_cookies(ws)
                        if cookies.get("sessionKey"):
                            result["cookies"] = cookies
                except Exception:
                    pass
                result["done"] = True
                return

            try:
                ws = _get_ws_url() or ws_url
                cookies = _cdp_get_cookies(ws)
                if cookies.get("sessionKey"):
                    log.info(f"Login detected! Cookies: {list(cookies.keys())}")
                    result["cookies"] = cookies
                    result["done"] = True
                    return
            except Exception as e:
                log.debug(f"Poll error: {e}")

            time.sleep(2)

        log.warning("Login poll timed out (5 min)")
        result["done"] = True

    poll_thread = threading.Thread(target=_poll, daemon=True)
    poll_thread.start()

    # --- Waiting dialog ---
    root = tk.Tk()
    root.title("Claude Usage Monitor")
    root.configure(bg="#1E1E1E")
    root.overrideredirect(True)
    root.attributes("-topmost", True)

    w, h = 360, 120
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    tk.Label(
        root, text="请在浏览器中登录 claude.ai",
        font=("Segoe UI", 12), fg="#FFFFFF", bg="#1E1E1E",
    ).pack(pady=(25, 6))
    tk.Label(
        root, text="登录成功后将自动继续...",
        font=("Segoe UI", 9), fg="#888888", bg="#1E1E1E",
    ).pack()
    tk.Button(
        root, text="取消", command=lambda: _cancel(),
        font=("Segoe UI", 9), fg="#CCCCCC", bg="#333333",
        activebackground="#555555", activeforeground="#FFFFFF",
        relief="flat", bd=0, padx=16, pady=4,
    ).pack(pady=(10, 0))

    def _cancel():
        result["done"] = True
        if proc.poll() is None:
            proc.terminate()
        root.destroy()

    def _check():
        if result["done"]:
            root.destroy()
            return
        root.after(500, _check)

    root.protocol("WM_DELETE_WINDOW", _cancel)
    root.after(500, _check)

    # Drag support
    drag = {"x": 0, "y": 0}
    root.bind("<Button-1>", lambda e: drag.update(x=e.x, y=e.y))
    root.bind("<B1-Motion>", lambda e: root.geometry(
        f"+{root.winfo_x() + e.x - drag['x']}+{root.winfo_y() + e.y - drag['y']}"))

    root.mainloop()
    poll_thread.join(timeout=3)

    if result["cookies"] and proc.poll() is None:
        log.info("Login success, closing browser...")
        proc.terminate()

    return result["cookies"]
