"""
Email Dispatcher and Preview Service for KG Recharge Tracker.
=============================================================
Manages SMTP transport, MIME message construction, Outlook CID attachments,
and standalone HTML preview generation.
"""

from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
from pathlib import Path
import smtplib

from config.settings import (
    PREVIEWS_DIR,
    REPORT_EMAIL_CC,
    REPORT_EMAIL_TO,
    SENDER_EMAIL,
    SMTP_PASSWORD,
    SMTP_PORT,
    SMTP_SERVER,
    SMTP_USERNAME,
    SMTP_USE_SSL,
    SMTP_USE_STARTTLS,
)
from kg.services.email_template import (
    CHART_IMAGE_CID,
    FOOTER_IMAGE,
    FOOTER_IMAGE_CID,
)


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


def send_kg_email(
    subject: str,
    html_body: str,
    attachment_paths: list[Path] | None = None,
    chart_path: Path | None = None,
    recipient_override: str | None = None,
) -> bool:
    """
    Send the executive performance email with attachments and inline CID images via SMTP.
    """
    recipient_to = recipient_override or REPORT_EMAIL_TO
    recipients = [r.strip() for r in recipient_to.split(",") if r.strip()]
    cc_recipients = [r.strip() for r in REPORT_EMAIL_CC.split(",") if r.strip()]
    all_recipients = recipients + cc_recipients

    if not recipients:
        print("[EMAIL-ERR] No recipient email specified; skipping send.")
        return False

    msg = MIMEMultipart("related")
    msg["Subject"] = subject
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(recipients)
    if cc_recipients:
        msg["Cc"] = ", ".join(cc_recipients)

    # Alternative container for HTML body
    msg_alt = MIMEMultipart("alternative")
    msg.attach(msg_alt)

    # Attach HTML
    msg_alt.attach(MIMEText(html_body, "html", "utf-8"))

    # Attach Inline Images (CID)
    if FOOTER_IMAGE.exists():
        with open(FOOTER_IMAGE, "rb") as f:
            img = MIMEImage(f.read())
            img.add_header("Content-ID", f"<{FOOTER_IMAGE_CID}>")
            img.add_header("Content-Disposition", "inline", filename=FOOTER_IMAGE.name)
            msg.attach(img)

    if chart_path and chart_path.exists():
        with open(chart_path, "rb") as f:
            img = MIMEImage(f.read())
            img.add_header("Content-ID", f"<{CHART_IMAGE_CID}>")
            img.add_header("Content-Disposition", "inline", filename=chart_path.name)
            msg.attach(img)

    # Attach Documents (.xlsx)
    if attachment_paths:
        for p in attachment_paths:
            if p.exists():
                with open(p, "rb") as f:
                    part = MIMEApplication(f.read(), Name=p.name)
                    part["Content-Disposition"] = f'attachment; filename="{p.name}"'
                    msg.attach(part)

    # SMTP Dispatch
    print(f"[EMAIL] Connecting to SMTP server {SMTP_SERVER}:{SMTP_PORT}...")
    try:
        if SMTP_USE_SSL:
            server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=30)
        else:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30)
            if SMTP_USE_STARTTLS:
                server.starttls()

        if SMTP_USERNAME and SMTP_PASSWORD:
            server.login(SMTP_USERNAME, SMTP_PASSWORD)

        server.send_message(msg, from_addr=SENDER_EMAIL, to_addrs=all_recipients)
        server.quit()
        print(f"[EMAIL] Successfully sent report email to: {', '.join(all_recipients)}")
        return True
    except Exception as e:
        print(f"[EMAIL-ERR] Failed to send email via SMTP: {e}")
        return False

