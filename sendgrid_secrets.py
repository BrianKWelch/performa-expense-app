"""Load SendGrid API key from Streamlit secrets."""

from __future__ import annotations

import base64
import re
from typing import List, Optional

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


def _try_sendgrid_parts() -> Optional[tuple[str, str]]:
    parts = _collect_sendgrid_parts()
    if not parts:
        return None
    label = f"PART×{len(parts)}" if len(parts) > 1 else "PART1"
    return "".join(parts), label


def _try_sendgrid_single() -> Optional[tuple[str, str]]:
    single = secret_get("SENDGRID_API_KEY")
    if single.strip():
        return single, "SENDGRID_API_KEY"
    return None


def _try_sendgrid_b64() -> Optional[tuple[str, str]]:
    b64 = secret_get("SENDGRID_API_KEY_B64")
    if not b64.strip():
        return None
    try:
        decoded = base64.b64decode(b64.strip(), validate=True).decode("utf-8")
        return decoded, "SENDGRID_API_KEY_B64"
    except Exception:
        return None


def raw_sendgrid_key_material() -> tuple[str, str]:
    for loader in (_try_sendgrid_parts, _try_sendgrid_single, _try_sendgrid_b64):
        result = loader()
        if result:
            return result

    raise KeyError(
        "No SendGrid key in secrets. Add SENDGRID_API_KEY on one line, "
        "or SENDGRID_API_KEY_PART1 + PART2 (PART1 must start with SG.)."
    )


def sendgrid_api_key() -> str:
    raw, _source = raw_sendgrid_key_material()
    key = normalize_sendgrid_key(raw)
    if not key.startswith("SG."):
        raise ValueError("SendGrid API key must start with SG.")
    if len(key) < 50:
        raise ValueError(
            f"SendGrid API key looks too short ({len(key)} characters). "
            "Use SENDGRID_API_KEY on one line in `.streamlit/secrets.toml` or Cloud Secrets."
        )
    return key


def has_sendgrid_key() -> bool:
    try:
        return bool(sendgrid_api_key())
    except (KeyError, FileNotFoundError, ValueError):
        return False


def _sendgrid_response_detail(ex: Exception) -> str:
    body = getattr(ex, "body", None)
    if body is None:
        return ""
    text = body.decode("utf-8", errors="replace") if isinstance(body, bytes) else str(body)
    if "Maximum credits exceeded" in text:
        return (
            "SendGrid blocked sending: **Maximum credits exceeded**. "
            "Your API key is valid, but this account has no email credits left. "
            "Open [SendGrid Billing](https://app.sendgrid.com/settings/billing) to add a plan "
            "or credits, or switch to an account with sending available."
        )
    if "does not match a verified Sender Identity" in text:
        return (
            "SendGrid blocked sending: the From address is not a verified sender. "
            "Verify your sender email in SendGrid → Settings → Sender Authentication."
        )
    return ""


def format_sendgrid_error(ex: Exception) -> str:
    msg = str(ex)
    detail = _sendgrid_response_detail(ex)
    if detail:
        return detail

    if 'has no key "SENDGRID_API_KEY"' in msg or (
        isinstance(ex, KeyError) and "SENDGRID_API_KEY" in str(ex)
    ):
        return (
            "SendGrid key not loaded. Add SENDGRID_API_KEY to `.streamlit/secrets.toml` "
            "(local) or Streamlit Cloud Secrets."
        )
    if "401" in msg or "Unauthorized" in msg:
        return (
            "SendGrid returned 401 Unauthorized. Check billing and sender verification, "
            "or update SENDGRID_API_KEY in secrets and restart the app."
        )
    if "403" in msg or "Forbidden" in msg:
        return (
            f"SendGrid forbidden (403). Verify `{secret_get('SENDER_EMAIL', 'SENDER_EMAIL')}` "
            "is a verified sender in SendGrid → Settings → Sender Authentication."
        )
    return f"Email failed: {msg}"
