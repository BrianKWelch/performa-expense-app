#!/usr/bin/env python3
"""Print base64-encoded SendGrid API key for Streamlit secrets (SENDGRID_API_KEY_B64)."""

import base64
import sys


def main() -> None:
    if len(sys.argv) > 1:
        raw = " ".join(sys.argv[1:])
    else:
        raw = input("Paste SendGrid API key: ").strip()

    if not raw:
        print("No key provided.", file=sys.stderr)
        sys.exit(1)

    b64 = base64.b64encode(raw.strip().encode("utf-8")).decode("ascii")
    print(f'SENDGRID_API_KEY_B64 = "{b64}"')


if __name__ == "__main__":
    main()
