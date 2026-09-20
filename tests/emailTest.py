"""Quick standalone test: does classic Outlook COM automation work?
Run with Outlook (classic) already open: python emailTestOutlook.py
"""
import win32com.client

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = "Test email from script"
    mail.To = "ronel.l@we-com.co.il"
    mail.Body = "If you got this, Outlook automation works."
    mail.Send()
    print("Success — check your Sent Items folder.")
except Exception as e:
    print(f"Failed: {e}")