import mimetypes
import os
import re
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import getaddresses
from pathlib import Path
from collections.abc import Mapping, Sequence

from dotenv import load_dotenv

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
    """Send an email through the configured SMTP server."""
    host = os.getenv("SMTP_SERVER") or os.getenv("SMTP_HOST")
    username = os.getenv("SMTP_USERNAME")
    password = os.getenv("SMTP_PASSWORD")
    sender = (
        from_email
        or os.getenv("EMAIL_SENDER")
        or os.getenv("SENDER_EMAIL")
        or os.getenv("SMTP_FROM")
        or username
    )
    to_recipients = (
        to
        or os.getenv("EMAIL_RECIPIENTS")
        or os.getenv("REPORT_EMAIL_TO")
        or os.getenv("REPORT_EMAIL")
    )
    cc_recipients_value = cc or os.getenv("REPORT_EMAIL_CC")

    if not to_recipients:
        return False
    if not host:
        raise ValueError("SMTP_SERVER or SMTP_HOST is missing from configuration and .env.")
    if not sender:
        raise ValueError(
            "SENDER_EMAIL, EMAIL_SENDER, SMTP_FROM, or SMTP_USERNAME must be configured."
        )
    if len(_addresses(sender, "From")) != 1:
        raise ValueError("Exactly one SMTP sender address must be configured.")
    if bool(username) != bool(password):
        raise ValueError("Configure both SMTP_USERNAME and SMTP_PASSWORD, or neither.")
    if inline_images and not html_body:
        raise ValueError("Inline images require an HTML email body.")

    port = int(os.getenv("SMTP_PORT", "25"))
    use_ssl = _env_bool("SMTP_USE_SSL", False)
    use_starttls = _env_bool("SMTP_USE_STARTTLS", False)
    if use_ssl and use_starttls:
        raise ValueError("Enable only one of SMTP_USE_SSL and SMTP_USE_STARTTLS.")

    cc_recipients = (
        _addresses(cc_recipients_value, "Cc") if cc_recipients_value else []
    )
    to_addresses = _addresses(to_recipients, "To")
    message = EmailMessage()
    message["From"] = sender
    message["To"] = to_recipients.replace(";", ",")
    if cc_recipients:
        message["Cc"] = cc_recipients_value.replace(";", ",")
    message["Subject"] = subject
    message.set_content(body)
    html_part = None
    if html_body:
        message.add_alternative(html_body, subtype="html")
        html_part = message.get_payload()[-1]

    for attachment in attachments:
        path = Path(attachment)
        if not path.is_file():
            raise FileNotFoundError(f"Email attachment not found: {path}")
        media_type, _ = mimetypes.guess_type(path.name)
        maintype, subtype = (media_type or "application/octet-stream").split("/", 1)
        message.add_attachment(
            path.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            filename=path.name,
        )

    for content_id, image_path in (inline_images or {}).items():
        if not re.fullmatch(r"[A-Za-z0-9._-]+", content_id):
            raise ValueError(f"Invalid inline image content ID: {content_id!r}")
        if f"cid:{content_id}" not in html_body:
            raise ValueError(
                f"HTML email body does not reference inline image {content_id!r}."
            )
        path = Path(image_path)
        if not path.is_file():
            raise FileNotFoundError(f"Inline email image not found: {path}")
        media_type, _ = mimetypes.guess_type(path.name)
        if not media_type or not media_type.startswith("image/"):
            raise ValueError(f"Unsupported inline image type: {path.suffix}")
        if html_part is None:
            raise ValueError("Inline images require an HTML email body.")
        subtype = media_type.split("/", 1)[1]
        html_part.add_related(
            path.read_bytes(),
            maintype="image",
            subtype=subtype,
            cid=f"<{content_id}>",
            disposition="inline",
            filename=path.name,
        )

    recipients = list(dict.fromkeys(to_addresses + cc_recipients))
    context = ssl.create_default_context() if use_ssl or use_starttls else None
    if use_ssl:
        with smtplib.SMTP_SSL(
            host, port, timeout=60, context=context
        ) as server:
            if username:
                server.login(username, password)
            refused = server.send_message(message, to_addrs=recipients)
    else:
        with smtplib.SMTP(host, port, timeout=60) as server:
            server.ehlo()
            if use_starttls:
                server.starttls(context=context)
                server.ehlo()
            if username:
                server.login(username, password)
            refused = server.send_message(message, to_addrs=recipients)

    if refused:
        raise RuntimeError(
            "SMTP server refused recipient(s): " + ", ".join(sorted(refused))
        )
    return True
