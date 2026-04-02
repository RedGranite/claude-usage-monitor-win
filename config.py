"""
Configuration management for Claude Usage Monitor.

Config file location: %APPDATA%/ClaudeUsageMonitor/config.json
Session key is encrypted using Windows DPAPI (Data Protection API),
which ties the encryption to the current Windows user account.
Only the same user on the same machine can decrypt the key.
"""

import json
import os
import base64
import ctypes
import ctypes.wintypes

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

CONFIG_DIR = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "ClaudeUsageMonitor",
)
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")

# ---------------------------------------------------------------------------
# Default configuration values
# ---------------------------------------------------------------------------

DEFAULT_CONFIG = {
    "session_key": "",       # Encrypted with DPAPI before saving to disk
    "cf_clearance": "",      # Cloudflare clearance cookie (encrypted)
    "org_id": "",        # Claude organization UUID
    "org_name": "",      # Claude organization display name
    "refresh_interval": 300,  # Auto-refresh interval in seconds (default 5min)
}

# ---------------------------------------------------------------------------
# Windows DPAPI encryption / decryption
#
# Uses CryptProtectData / CryptUnprotectData from Windows crypt32.dll.
# The encryption key is derived from the current user's Windows login
# credentials, so:
#   - No password needed from the user
#   - Only the same Windows user on the same machine can decrypt
#   - Survives reboots (tied to user account, not session)
#   - Other users or other machines cannot read the data
# ---------------------------------------------------------------------------


class _DATA_BLOB(ctypes.Structure):
    """Windows DATA_BLOB structure used by DPAPI functions."""
    _fields_ = [
        ("cbData", ctypes.wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_char)),
    ]


def _dpapi_encrypt(plaintext: str) -> str:
    """
    Encrypt a string using Windows DPAPI.

    Args:
        plaintext: The string to encrypt (e.g. session key).

    Returns:
        Base64-encoded ciphertext string, safe for JSON storage.
    """
    data = plaintext.encode("utf-8")
    input_blob = _DATA_BLOB(len(data), ctypes.create_string_buffer(data, len(data)))
    output_blob = _DATA_BLOB()

    # CryptProtectData(pDataIn, szDataDescr, pOptionalEntropy,
    #                   pvReserved, pPromptStruct, dwFlags, pDataOut)
    if not ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(input_blob),  # data to encrypt
        "ClaudeUsageMonitor",      # description (visible in DPAPI logs)
        None,                      # no additional entropy
        None,                      # reserved
        None,                      # no prompt
        0,                         # flags
        ctypes.byref(output_blob), # output
    ):
        raise OSError("DPAPI CryptProtectData failed")

    # Copy encrypted bytes and free the Windows-allocated buffer
    encrypted = ctypes.string_at(output_blob.pbData, output_blob.cbData)
    ctypes.windll.kernel32.LocalFree(output_blob.pbData)

    return base64.b64encode(encrypted).decode("ascii")


def _dpapi_decrypt(encrypted_b64: str) -> str:
    """
    Decrypt a DPAPI-encrypted base64 string back to plaintext.

    Args:
        encrypted_b64: Base64-encoded ciphertext from _dpapi_encrypt().

    Returns:
        Original plaintext string.

    Raises:
        OSError: If decryption fails (wrong user, corrupted data, etc.)
    """
    data = base64.b64decode(encrypted_b64)
    input_blob = _DATA_BLOB(len(data), ctypes.create_string_buffer(data, len(data)))
    output_blob = _DATA_BLOB()

    if not ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(input_blob),
        None,   # description out (not needed)
        None,   # no additional entropy
        None,   # reserved
        None,   # no prompt
        0,      # flags
        ctypes.byref(output_blob),
    ):
        raise OSError("DPAPI CryptUnprotectData failed")

    decrypted = ctypes.string_at(output_blob.pbData, output_blob.cbData)
    ctypes.windll.kernel32.LocalFree(output_blob.pbData)

    return decrypted.decode("utf-8")


# ---------------------------------------------------------------------------
# Config load / save
# ---------------------------------------------------------------------------


def load_config() -> dict:
    """
    Load configuration from disk.

    The session_key field is decrypted from DPAPI on load.
    If the file doesn't exist or is corrupted, returns defaults.
    Handles migration from plaintext keys (pre-encryption versions).

    Returns:
        Configuration dict with all keys from DEFAULT_CONFIG.
    """
    if not os.path.exists(CONFIG_FILE):
        return dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            stored = json.load(f)

        config = dict(DEFAULT_CONFIG)
        config.update(stored)

        # Decrypt session key if it's encrypted
        encrypted_key = config.get("session_key_encrypted", "")
        if encrypted_key:
            try:
                config["session_key"] = _dpapi_decrypt(encrypted_key)
            except OSError:
                config["session_key"] = ""

        # Migration: if there's a plaintext session_key but no encrypted one
        elif config.get("session_key", "").startswith("sk-ant-"):
            pass  # Keep plaintext key; it will be encrypted on next save

        # Decrypt cf_clearance if it's encrypted
        encrypted_cf = config.get("cf_clearance_encrypted", "")
        if encrypted_cf:
            try:
                config["cf_clearance"] = _dpapi_decrypt(encrypted_cf)
            except OSError:
                config["cf_clearance"] = ""

        return config
    except (json.JSONDecodeError, OSError):
        return dict(DEFAULT_CONFIG)


def save_config(config: dict):
    """
    Save configuration to disk.

    The session_key is encrypted with DPAPI before writing.
    The plaintext key is NEVER stored on disk.

    Args:
        config: Configuration dict to save.
    """
    os.makedirs(CONFIG_DIR, exist_ok=True)

    # Build a copy for disk storage
    to_save = dict(config)

    # Encrypt secrets before saving — plaintext NEVER written to disk
    for field in ("session_key", "cf_clearance"):
        value = to_save.get(field, "")
        if value:
            to_save[f"{field}_encrypted"] = _dpapi_encrypt(value)
        else:
            to_save[f"{field}_encrypted"] = to_save.get(f"{field}_encrypted", "")
        to_save[field] = ""

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(to_save, f, indent=2, ensure_ascii=False)
