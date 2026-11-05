import mimetypes
import os
import re
import sys
from pathlib import Path
from collections.abc import Mapping, Sequence

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
    attachments: Sequence[str | Path],
    to: str = None,
    cc: str = None,
    html_body: str = None,
    display_only: bool = False,
    inline_images: Mapping[str, str | Path] | None = None,
) -> bool:
    """
    Sends an automated email via local Outlook Desktop client.
    Supports To, CC, plain text or rich HTML body, file attachments, and inline
    images referenced from HTML as ``<img src="cid:content_id">``.
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
    if inline_images and not html_body:
        raise ValueError("Inline images require an HTML email body.")

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = to_recipient
        if cc_recipient:
            mail.CC = cc_recipient

        mail.Subject = subject

        for attachment in attachments:
            att_path = Path(attachment)
            if not att_path.is_file():
                raise FileNotFoundError(f"Email attachment not found: {att_path}")
            mail.Attachments.Add(str(att_path.resolve()))

        for content_id, image_path in (inline_images or {}).items():
            if not re.fullmatch(r"[A-Za-z0-9._-]+", content_id):
                raise ValueError(f"Invalid inline image content ID: {content_id!r}")

            image = Path(image_path)
            if not image.is_file():
                raise FileNotFoundError(f"Inline email image not found: {image}")
            if f"cid:{content_id}" not in html_body:
                raise ValueError(
                    f"HTML email body does not reference inline image {content_id!r}."
                )
            media_type, _ = mimetypes.guess_type(image.name)
            if not media_type or not media_type.startswith("image/"):
                raise ValueError(f"Unsupported inline image type: {image.suffix}")

            inline_attachment = mail.Attachments.Add(
                str(image.resolve()), 1, 0, image.name
            )
            accessor = inline_attachment.PropertyAccessor
            accessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                content_id,
            )
            accessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x370E001F",
                media_type,
            )
            accessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
                True,
            )

        if html_body:
            mail.HTMLBody = html_body
        else:
            mail.Body = body

        if display_only:
            mail.Display(True)
        else:
            mail.Send()

        return True
    except Exception as e:
        raise RuntimeError(f"Outlook Dispatch failed: {e}. Please ensure Microsoft Outlook is open and signed in.")
