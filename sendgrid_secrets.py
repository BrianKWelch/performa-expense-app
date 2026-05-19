"""Load SendGrid API key from Streamlit secrets."""

from __future__ import annotations

import base64
import re
import urllib.error
import urllib.request
from typing import Any, Dict, List, Optional

import streamlit as st


def secret_get(key: str, default: str = "") -> str:
    try:
        val = st.secrets[key]
        return "" if val is None else str(val)
    except (KeyError, FileNotFoundError):
        return default


def secret_float(key: str, default: float) -> float:
    raw = secret_get(key)
    if not raw.strip():
        return default
    try:
        return float(raw)
    except ValueError:
        return default


def normalize_sendgrid_key(raw: str) -> str:
    return (
        str(raw)
        .strip()
        .strip('"')
        .strip("'")
        .replace("\n", "")
        .replace("\r", "")
        .replace("\t", "")
        .replace(" ", "")
    )


def _collect_sendgrid_parts() -> List[str]:
    """All SENDGRID_API_KEY_PART{n} found in secrets, sorted by n."""
    part_nums: List[tuple[int, str]] = []
    try:
        keys = list(st.secrets.keys())
    except Exception:
        keys = []
    pattern = re.compile(r"^SENDGRID_API_KEY_PART(\d+)$")
    for key in keys:
        match = pattern.match(str(key))
        if match:
            part_nums.append((int(match.group(1)), secret_get(key)))
    part_nums.sort(key=lambda x: x[0])
    return [value for _num, value in part_nums if value.strip()]


def raw_sendgrid_key_material() -> tuple[str, str]:
    # Best option for Streamlit Cloud: one line, no wrapping issues
    b64 = secret_get("SENDGRID_API_KEY_B64")
    if b64.strip():
        try:
            decoded = base64.b64decode(b64.strip(), validate=True).decode("utf-8")
            return decoded, "SENDGRID_API_KEY_B64"
        except Exception as ex:
            raise ValueError(f"SENDGRID_API_KEY_B64 is not valid base64: {ex}") from ex

    parts = _collect_sendgrid_parts()
    if parts:
        label = f"PART×{len(parts)}" if len(parts) > 1 else "PART1"
        return "".join(parts), label

    single = secret_get("SENDGRID_API_KEY")
    if single.strip():
        return single, "SENDGRID_API_KEY"

    raise KeyError(
        "No SendGrid key in secrets. Easiest fix: add SENDGRID_API_KEY_B64 (one line). "
        "See the app's Email setup help or secrets.toml.example."
    )


def sendgrid_api_key() -> str:
    raw, _source = raw_sendgrid_key_material()
    key = normalize_sendgrid_key(raw)
    if not key.startswith("SG."):
        raise ValueError("SendGrid API key must start with SG.")
    if len(key) < 50:
        raise ValueError(
            f"SendGrid API key looks too short ({len(key)} characters). "
            "If using PART1, PART2, … each value must be on one line (no Enter inside quotes). "
            "Recommended: use SENDGRID_API_KEY_B64 instead (single line)."
        )
    return key


def has_sendgrid_key() -> bool:
    try:
        return bool(sendgrid_api_key())
    except (KeyError, FileNotFoundError, ValueError):
        return False


def sendgrid_key_diagnostics() -> Dict[str, Any]:
    parts_loaded = []
    try:
        keys = list(st.secrets.keys())
    except Exception:
        keys = []
    for key in keys:
        if str(key).startswith("SENDGRID_API_KEY_PART"):
            parts_loaded.append(str(key))
    has_b64 = bool(secret_get("SENDGRID_API_KEY_B64").strip())
    try:
        raw, source = raw_sendgrid_key_material()
        key = normalize_sendgrid_key(raw)
        return {
            "ok": key.startswith("SG.") and len(key) >= 50,
            "source": source,
            "length": len(key),
            "prefix": key[:8] + "…" if len(key) >= 8 else key,
            "has_b64_secret": has_b64,
            "part_keys_found": parts_loaded,
        }
    except Exception as ex:
        return {
            "ok": False,
            "error": str(ex),
            "has_b64_secret": has_b64,
            "part_keys_found": parts_loaded,
        }


def verify_sendgrid_key_with_api() -> Optional[str]:
    try:
        key = sendgrid_api_key()
    except (KeyError, ValueError) as ex:
        return str(ex)

    req = urllib.request.Request(
        "https://api.sendgrid.com/v3/scopes",
        headers={"Authorization": f"Bearer {key}"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            if 200 <= resp.status < 300:
                return None
            return f"SendGrid returned status {resp.status} when validating the API key."
    except urllib.error.HTTPError as ex:
        if ex.code == 401:
            diag = sendgrid_key_diagnostics()
            return (
                "SendGrid rejected the API key (401). "
                f"Loaded from {diag.get('source', '?')}, length {diag.get('length', 0)}. "
                "Create a new API key in SendGrid (Mail Send), update secrets, Reboot app."
            )
        return f"SendGrid key check failed: HTTP {ex.code}"
    except Exception as ex:
        return f"Could not reach SendGrid to validate the key: {ex}"


def format_sendgrid_error(ex: Exception) -> str:
    msg = str(ex)
    if 'has no key "SENDGRID_API_KEY"' in msg or (
        isinstance(ex, KeyError) and "SENDGRID_API_KEY" in str(ex)
    ):
        return (
            "SendGrid key not loaded. Use SENDGRID_API_KEY_B64 in Cloud secrets (one line) — "
            "see **Email setup help** on this page. Push latest code from GitHub, then Reboot."
        )
    if "401" in msg or "Unauthorized" in msg:
        diag = sendgrid_key_diagnostics()
        return (
            "SendGrid rejected the API key (401). "
            f"Source: {diag.get('source', '?')}, length: {diag.get('length', '?')}. "
            "Try SENDGRID_API_KEY_B64 or a new SendGrid API key."
        )
    if "403" in msg or "Forbidden" in msg:
        return (
            f"SendGrid forbidden (403). Verify `{secret_get('SENDER_EMAIL', 'SENDER_EMAIL')}` "
            "is a verified sender in SendGrid → Settings → Sender Authentication."
        )
    return f"Email failed: {msg}"


def make_sendgrid_b64(api_key: str) -> str:
    """Helper for local setup: base64-encode a SendGrid key for secrets."""
    return base64.b64encode(api_key.strip().encode("utf-8")).decode("ascii")
