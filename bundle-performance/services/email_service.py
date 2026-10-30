import os
import sys
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()


def is_outlook_running() -> bool:
    """Checks if Outlook desktop application is active."""
    if sys.platform != "win32":
        return False
    try:
        import win32com.client
        win32com.client.GetActiveObject("Outlook.Application")
        return True
    except Exception:
        return False


def send_outlook_email(
    subject: str,
    body: str,
    attachments: list,
    to: str = None,
    cc: str = None,
    html_body: str = None,
    display_only: bool = False
) -> bool:
    """
    Sends an automated email via local Outlook Desktop client.
    Supports To, CC, plain text or rich HTML body, and file attachments.
    """
    if sys.platform != "win32":
        raise RuntimeError(
            "Outlook desktop integration requires Windows. "
            "Use Microsoft Graph or SMTP for Linux/macOS."
        )

    import win32com.client

    to_recipient = to or os.getenv("REPORT_EMAIL_TO") or os.getenv("REPORT_EMAIL")
    cc_recipient = cc or os.getenv("REPORT_EMAIL_CC", "")

    if not to_recipient:
        raise ValueError("Recipient email (To) is missing in configuration and .env")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = to_recipient
        if cc_recipient:
            mail.CC = cc_recipient

        mail.Subject = subject

        if html_body:
            mail.HTMLBody = html_body
        else:
            mail.Body = body

        for attachment in attachments:
            att_path = Path(attachment)
            if att_path.exists():
                mail.Attachments.Add(str(att_path.resolve()))

        if display_only:
            mail.Display(True)
        else:
            mail.Send()

        return True
    except Exception as e:
        raise RuntimeError(f"Outlook Dispatch failed: {e}. Please ensure Microsoft Outlook is open and signed in.")
