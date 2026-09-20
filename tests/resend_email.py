"""
Standalone test: resend the email via Brevo, using an existing screenshots
folder, without re-running the full browser capture.

Usage: python resend_email.py screenshots\20260920_0423
"""

import os
import sys
import base64
from datetime import datetime
from pathlib import Path

import requests

BREVO_API_KEY = os.environ.get("BREVO_API_KEY", "")
EMAIL_FROM = os.environ.get("EMAIL_FROM", "")
EMAIL_TO = os.environ.get("EMAIL_TO", "")


def main():
    if len(sys.argv) < 2:
        print("Usage: python resend_email.py <folder>")
        sys.exit(1)

    out_dir = Path(sys.argv[1])
    images = sorted(out_dir.glob("*.png"))
    print(f"Found {len(images)} PNG files in {out_dir}")

    if not BREVO_API_KEY or not EMAIL_FROM or not EMAIL_TO:
        print("Missing one of BREVO_API_KEY / EMAIL_FROM / EMAIL_TO.")
        sys.exit(1)

    to_list = [{"email": addr.strip()} for addr in EMAIL_TO.split(",") if addr.strip()]

    attachments = []
    for path in images:
        content = base64.b64encode(path.read_bytes()).decode()
        attachments.append({"content": content, "name": path.name})

    payload = {
        "sender": {"name": "CNOM Automation", "email": EMAIL_FROM},
        "to": to_list,
        "subject": f"CNOM Status Report (resend test) - {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "htmlContent": f"<p>Test resend. Files: {', '.join(p.name for p in images)}</p>",
        "attachment": attachments,
    }

    response = requests.post(
        "https://api.brevo.com/v3/smtp/email",
        headers={
            "api-key": BREVO_API_KEY,
            "Content-Type": "application/json",
            "accept": "application/json",
        },
        json=payload,
        timeout=30,
    )
    print(f"\nStatus: {response.status_code}")
    print(response.text)


if __name__ == "__main__":
    main()