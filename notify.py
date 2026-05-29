#!/usr/bin/env python3
"""Send a task-completion email via Gmail SMTP.

Usage:
    python notify.py "subject line" "email body"
"""

import os
import smtplib
import sys
from email.mime.text import MIMEText

GMAIL_ADDR = "yzsun0123@gmail.com"


def send(subject: str, body: str) -> None:
    app_pw = os.environ.get("GMAIL_APP_PASSWORD")
    if not app_pw:
        print("ERROR: GMAIL_APP_PASSWORD not set", file=sys.stderr)
        sys.exit(1)

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = GMAIL_ADDR
    msg["To"] = GMAIL_ADDR

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(GMAIL_ADDR, app_pw)
        smtp.send_message(msg)
    print("Email sent.")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} SUBJECT BODY", file=sys.stderr)
        sys.exit(1)
    send(sys.argv[1], sys.argv[2])
