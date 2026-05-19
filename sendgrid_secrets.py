"""Load SendGrid API key from Streamlit secrets."""

from __future__ import annotations

import base64
import re
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional

import streamlit as st

LOCAL_SECRETS_PATH = Path(".streamlit") / "secrets.toml"

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


def _part_length_diagnostics() -> List[Dict[str, Any]]:
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
    return [
        {
            "part": num,
            "chars": len(normalize_sendgrid_key(value)),
        }
        for num, value in part_nums
        if value.strip()
    ]


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
        # Ignore invalid B64 so PART1+PART2 or SENDGRID_API_KEY still work.
        return None


def raw_sendgrid_key_material() -> tuple[str, str]:
    for loader in (_try_sendgrid_parts, _try_sendgrid_single, _try_sendgrid_b64):
        result = loader()
        if result:
            return result

    raise KeyError(
        "No SendGrid key in secrets. Add SENDGRID_API_KEY_PART1 + PART2 "
        "(PART1 must start with SG.) or SENDGRID_API_KEY on one line."
    )


def sendgrid_api_key() -> str:
    raw, _source = raw_sendgrid_key_material()
    key = normalize_sendgrid_key(raw)
    if not key.startswith("SG."):
        raise ValueError("SendGrid API key must start with SG.")
    if len(key) < 50:
        raise ValueError(
            f"SendGrid API key looks too short ({len(key)} characters). "
            "PART1 must include SG. (e.g. SG.xxxx…). Each PART value on one line only."
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
        length = len(key)
        return {
            "ok": key.startswith("SG.") and length >= 50,
            "source": source,
            "length": length,
            "prefix": key[:8] + "…" if length >= 8 else key,
            "has_b64_secret": has_b64,
            "part_keys_found": parts_loaded,
            "part_lengths": _part_length_diagnostics(),
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
                "Check PART1 starts with SG., or use SENDGRID_API_KEY on one line. "
                "If the key was rotated, create a new Mail Send key in SendGrid and update secrets."
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
            "SendGrid key not loaded. Add SENDGRID_API_KEY_PART1 + PART2 in secrets "
            "(PART1 must start with SG.). Remove any broken SENDGRID_API_KEY_B64 line."
        )
    if "401" in msg or "Unauthorized" in msg:
        hint = loaded_sendgrid_key_hint()
        return (
            "SendGrid rejected the API key (401). "
            f"Key currently loaded from secrets: {hint}. "
            "If you created a new key, it must be saved in `.streamlit/secrets.toml` "
            "(sidebar paste alone does not update secrets). "
            "Use **Save to local secrets** in the sidebar, or paste "
            "`SENDGRID_API_KEY = \"SG.…\"` on one line, then **stop and restart** Streamlit."
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


def split_sendgrid_key_for_secrets(raw: str) -> tuple[str, str]:
    """Split a full SG.… key into PART1 + PART2 for secrets.toml (first line ~25 chars)."""
    key = normalize_sendgrid_key(raw)
    if not key.startswith("SG."):
        raise ValueError("SendGrid API key must start with SG.")
    if len(key) < 50:
        raise ValueError(f"Key too short ({len(key)} chars). Copy the full key from SendGrid.")
    return key[:25], key[25:]


def format_secrets_toml_lines(raw_key: str) -> str:
    key = normalize_sendgrid_key(raw_key)
    if not key.startswith("SG."):
        raise ValueError("SendGrid API key must start with SG.")
    if len(key) < 50:
        raise ValueError(f"Key too short ({len(key)} chars). Copy the full key from SendGrid.")
    return f'SENDGRID_API_KEY = "{key}"'


def loaded_sendgrid_key_hint() -> str:
    try:
        key = sendgrid_api_key()
        return f"{key[:12]}… (length {len(key)})"
    except Exception as ex:
        return f"not loaded ({ex})"


def pasted_key_matches_loaded(pasted: str) -> bool:
    try:
        loaded = sendgrid_api_key()
        pasted_norm = normalize_sendgrid_key(pasted)
        return pasted_norm == loaded
    except Exception:
        return False


def _strip_sendgrid_lines(text: str) -> str:
    drop_prefixes = (
        "SENDGRID_API_KEY",
        "SENDGRID_API_KEY_PART",
        "SENDGRID_API_KEY_B64",
    )
    kept: List[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if any(stripped.startswith(p) for p in drop_prefixes):
            continue
        kept.append(line)
    return "\n".join(kept).rstrip()


def write_local_secrets_sendgrid_key(raw_key: str) -> None:
    """Update .streamlit/secrets.toml with a single-line SENDGRID_API_KEY (local dev)."""
    key = normalize_sendgrid_key(raw_key)
    if not key.startswith("SG.") or len(key) < 50:
        raise ValueError("Paste the full SendGrid API key from the dashboard.")

    path = LOCAL_SECRETS_PATH
    if not path.parent.exists():
        path.parent.mkdir(parents=True, exist_ok=True)

    existing = path.read_text(encoding="utf-8") if path.exists() else ""
    body = _strip_sendgrid_lines(existing)
    block = f'SENDGRID_API_KEY = "{key}"'
    new_text = f"{body}\n\n{block}\n" if body.strip() else f"{block}\n"
    path.write_text(new_text, encoding="utf-8")
