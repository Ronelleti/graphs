"""Quick standalone test: does Exchange Online SMTP sending work?
Set these first (PowerShell):
    $env:SMTP_USERNAME = "@we-com.co.il"
    $env:SMTP_PASSWORD = "your_password_or_app_password"
Then run: python test_email.py
"""
import os
import smtplib
from email.message import EmailMessage

SMTP_USERNAME = os.environ.get("SMTP_USERNAME", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")

msg = EmailMessage()
msg["Subject"] = "Test email from script"
msg["From"] = SMTP_USERNAME
msg["To"] = "ronel.l@we-com.co.il"
msg.set_content("If you got this, SMTP sending works.")

try:
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        server.send_message(msg)
    print("Success — check your inbox / Sent Items.")
except Exception as e:
    print(f"Failed: {e}")