"""
Email Service for KG Recharge Tracker.
=====================================
Sends emails through internal SMTP relay using IP-based authorization and modern EmailMessage API.
Provides standalone HTML preview generation.
"""

from collections.abc import Mapping, Sequence
from email.message import EmailMessage
from email.utils import getaddresses
import mimetypes
import os
from pathlib import Path
import re
import smtplib
import ssl
from dotenv import load_dotenv

from config.settings import PREVIEWS_DIR

load_dotenv()


def _env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    raise ValueError(f"{name} must be a boolean value.")


def _addresses(value: str, label: str) -> list[str]:
    addresses = [
        address
        for _, address in getaddresses([value.replace(";", ",")])
        if address
    ]
    if not addresses or any("@" not in address for address in addresses):
        raise ValueError(f"Valid {label} email address(es) are required.")
    return addresses


def save_html_previews(html_content: str, filename_prefix: str = "kg_performance_email") -> dict[str, Path]:
    """
    Save standalone HTML preview files for browser inspection.
    """
    PREVIEWS_DIR.mkdir(parents=True, exist_ok=True)
    desktop_path = PREVIEWS_DIR / f"{filename_prefix}_desktop.html"
    mobile_path = PREVIEWS_DIR / f"{filename_prefix}_mobile.html"

    desktop_path.write_text(html_content, encoding="utf-8")

    # Mobile wrapped preview
    mobile_html = f"""<!DOCTYPE html>
<html>
<head><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Mobile Preview</title></head>
<body style="margin:0;padding:20px;background:#E5E7EB;display:flex;justify-content:center;">
  <div style="width:390px;background:#FFF;border-radius:24px;box-shadow:0 10px 25px rgba(0,0,0,0.15);overflow:hidden;border:8px solid #1F2937;">
    {html_content}
  </div>
</body>
</html>"""
    mobile_path.write_text(mobile_html, encoding="utf-8")

    print(f"[PREVIEW] Desktop preview saved: {desktop_path}")
    print(f"[PREVIEW] Mobile preview saved: {mobile_path}")
    return {"desktop": desktop_path, "mobile": mobile_path}


def send_smtp_email(
    subject: str,
    body: str,
    attachments: Sequence[str | Path],
    to: str = None,
    cc: str = None,
    from_email: str = None,
    html_body: str = None,
    inline_images: Mapping[str, str | Path] | None = None,
) -> bool:
    """
    Send an email through the internal SMTP relay using IP-based authorization.

    No SMTP username, password, SSL, or STARTTLS authentication is required.
    The source machine IP must be whitelisted by the SMTP administrator.
    """

    # SMTP configuration
    host = os.getenv("SMTP_HOST", "10.3.124.25")
    port = int(os.getenv("SMTP_PORT", "25"))

    sender = (
        from_email
        or os.getenv("EMAIL_FROM")
        or os.getenv("EMAIL_SENDER")
        or os.getenv("SMTP_FROM")
    )

    to_recipients = (
        to
        or os.getenv("EMAIL_TO")
        or os.getenv("EMAIL_RECIPIENTS")
        or os.getenv("REPORT_EMAIL_TO")
        or os.getenv("REPORT_EMAIL")
    )

    cc_recipients_value = cc or os.getenv("REPORT_EMAIL_CC")

    # Validate email configuration
    if not to_recipients:
        raise ValueError("No email recipients configured.")

    if not sender:
        raise ValueError("EMAIL_FROM or SMTP_FROM must be configured.")

    if len(_addresses(sender, "From")) != 1:
        raise ValueError("Exactly one sender email address is required.")

    if inline_images and not html_body:
        raise ValueError("Inline images require an HTML email body.")

    # Validate recipients
    to_addresses = _addresses(to_recipients, "To")

    cc_recipients = (
        _addresses(cc_recipients_value, "Cc")
        if cc_recipients_value
        else []
    )

    recipients = list(dict.fromkeys(to_addresses + cc_recipients))

    # Create email message
    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(to_addresses)

    if cc_recipients:
        message["Cc"] = ", ".join(cc_recipients)

    message["Subject"] = subject

    message.set_content(body)

    html_part = None

    if html_body:
        message.add_alternative(html_body, subtype="html")
        html_part = message.get_payload()[-1]

    # Attach files
    for attachment in attachments:
        path = Path(attachment)

        if not path.is_file():
            raise FileNotFoundError(
                f"Email attachment not found: {path}"
            )

        media_type, _ = mimetypes.guess_type(path.name)

        maintype, subtype = (
            (media_type or "application/octet-stream").split("/", 1)
        )

        message.add_attachment(
            path.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            filename=path.name,
        )

    # Attach inline images
    for content_id, image_path in (inline_images or {}).items():

        if not re.fullmatch(r"[A-Za-z0-9._-]+", content_id):
            raise ValueError(
                f"Invalid inline image content ID: {content_id!r}"
            )

        if f"cid:{content_id}" not in html_body:
            raise ValueError(
                f"HTML body does not reference image: {content_id}"
            )

        path = Path(image_path)

        if not path.is_file():
            raise FileNotFoundError(
                f"Inline image not found: {path}"
            )

        media_type, _ = mimetypes.guess_type(path.name)

        if not media_type or not media_type.startswith("image/"):
            raise ValueError(
                f"Unsupported inline image type: {path.suffix}"
            )

        subtype = media_type.split("/", 1)[1]

        html_part.add_related(
            path.read_bytes(),
            maintype="image",
            subtype=subtype,
            cid=f"<{content_id}>",
            disposition="inline",
            filename=path.name,
        )

    # Connect to internal SMTP relay
    print(f"[SMTP] Connecting to {host}:{port} ...")

    try:
        with smtplib.SMTP(host, port, timeout=60) as server:

            server.ehlo()

            # No SMTP authentication or STARTTLS.
            # Relay access is controlled through source IP whitelisting.

            refused = server.send_message(
                message,
                from_addr=sender,
                to_addrs=recipients,
            )

        if refused:
            raise RuntimeError(
                "SMTP server refused recipient(s): "
                + ", ".join(sorted(refused))
            )

        print(
            f"[SUCCESS] Email successfully accepted by SMTP relay "
            f"for: {', '.join(recipients)}"
        )

        return True

    except (smtplib.SMTPException, OSError) as exc:
        print(f"[SMTP ERROR] Email delivery failed: {exc}")
        raise


# Backward-compatible alias
send_kg_email = send_smtp_email
