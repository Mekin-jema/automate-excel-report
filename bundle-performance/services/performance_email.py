"""Responsive, data-driven Data and Voice performance email content."""

from __future__ import annotations

import base64
import mimetypes
from collections.abc import Mapping, Sequence
from datetime import datetime
from html import escape
from pathlib import Path

KPI_LABELS = ("Revenue", "Bundle Subscribers", "Purchase Quantity", "Usage")

REPORTING_NOTE = (
    "*Bundle subscriber totals represent the sum of daily distinct subscribers "
    "across individual bundles. Therefore, subscribers using multiple bundles "
    "may be counted more than once."
)


def format_report_period(start_date: str, end_date: str) -> str:
    start = datetime.strptime(start_date, "%Y%m%d")
    end = datetime.strptime(end_date, "%Y%m%d")
    if start == end:
        return f"{start.day} {start.strftime('%B %Y')}"
    if start.year == end.year and start.month == end.month:
        return f"{start.day}–{end.day} {start.strftime('%B %Y')}"
    return f"{start.day} {start.strftime('%b %Y')} – {end.day} {end.strftime('%b %Y')}"


def _image_source(image: str | Path, *, embed: bool) -> str:
    if not embed:
        return f"cid:{image}"

    image_path = Path(image)
    if not image_path.is_file():
        raise FileNotFoundError(f"Dashboard image not found: {image_path}")
    media_type, _ = mimetypes.guess_type(image_path.name)
    if not media_type or not media_type.startswith("image/"):
        raise ValueError(f"Unsupported dashboard image type: {image_path.suffix}")
    encoded = base64.b64encode(image_path.read_bytes()).decode("ascii")
    return f"data:{media_type};base64,{encoded}"


def _validate_metrics(metrics: Mapping[str, Sequence[str]]) -> None:
    missing = set(KPI_LABELS) - set(metrics)
    if missing:
        raise ValueError(f"Missing report KPI values: {', '.join(sorted(missing))}")
    invalid = [
        label
        for label in KPI_LABELS
        if len(metrics[label]) != 2 or not all(metrics[label])
    ]
    if invalid:
        raise ValueError(f"Each KPI requires Data and Voice values: {', '.join(invalid)}")


def build_html_email(
    reporting_period: str,
    metrics: Mapping[str, Sequence[str]],
    data_dashboard: str | Path = "data_dashboard",
    voice_dashboard: str | Path = "voice_dashboard",
    *,
    embed_images: bool = False,
) -> str:
    """Build email-safe responsive HTML from the current report metrics."""
    _validate_metrics(metrics)
    data_src = _image_source(data_dashboard, embed=embed_images)
    voice_src = _image_source(voice_dashboard, embed=embed_images)

    kpi_cards = []
    for label in KPI_LABELS:
        data_value, voice_value = metrics[label]
        kpi_cards.append(
            '<td class="kpi-cell" width="25%" valign="top" '
            'style="width:25%;padding:5px">'
            '<table role="presentation" width="100%" cellpadding="0" '
            'cellspacing="0" style="width:100%;border:1px solid #DDE7E1;'
            'background-color:#FFFFFF">'
            '<tr><td style="padding:12px 12px 6px;color:#52645A;'
            'font-family:Arial,Helvetica,sans-serif;font-size:12px;'
            f'font-weight:bold">{escape(label)}</td></tr>'
            '<tr><td style="padding:0 12px 12px">'
            '<table role="presentation" width="100%" cellpadding="0" '
            'cellspacing="0"><tr>'
            '<td width="50%" valign="top" style="width:50%;padding-right:5px">'
            '<div style="color:#00843D;font-family:Arial,Helvetica,sans-serif;'
            'font-size:10px;font-weight:bold">DATA</div>'
            '<div class="kpi-value" style="padding-top:3px;color:#26332B;'
            'font-family:Arial,Helvetica,sans-serif;font-size:16px;'
            f'font-weight:bold;line-height:1.25">{escape(data_value)}</div></td>'
            '<td width="50%" valign="top" style="width:50%;padding-left:5px">'
            '<div style="color:#006341;font-family:Arial,Helvetica,sans-serif;'
            'font-size:10px;font-weight:bold">VOICE</div>'
            '<div class="kpi-value" style="padding-top:3px;color:#26332B;'
            'font-family:Arial,Helvetica,sans-serif;font-size:16px;'
            f'font-weight:bold;line-height:1.25">{escape(voice_value)}</div></td>'
            '</tr></table></td></tr></table></td>'
        )

    cards_html = "".join(kpi_cards)
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<title>Data &amp; Voice Business Performance Report</title>
<style type="text/css">
html,body{{margin:0!important;padding:0!important;width:100%!important;}}
table{{border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt;}}
img{{border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic;}}
@media only screen and (max-width:620px){{
  .email-container{{width:100%!important;}}
  .outer-pad{{padding:8px 0!important;}}
  .content-pad{{padding-left:16px!important;padding-right:16px!important;}}
  .header-title{{font-size:24px!important;line-height:1.2!important;}}
  .kpi-cell{{display:block!important;width:100%!important;padding:4px 0!important;}}
  .dashboard-image{{width:100%!important;height:auto!important;}}
  .footer-copy{{font-size:13px!important;line-height:1.55!important;}}
}}
</style>
<!--[if mso]>
<style type="text/css">
.email-container{{width:800px!important;}}
.kpi-cell{{display:table-cell!important;width:25%!important;}}
</style>
<![endif]-->
</head>
<body style="margin:0;padding:0;background-color:#F2F5F3">
<div style="display:none;font-size:1px;color:#F2F5F3;line-height:1px;
max-height:0;max-width:0;opacity:0;overflow:hidden">
Data &amp; Voice Business Performance | {escape(reporting_period)}
</div>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
       style="width:100%;background-color:#F2F5F3">
<tr><td class="outer-pad" align="center" style="padding:24px 12px">
<!--[if mso]>
<table role="presentation" width="800" align="center" cellpadding="0" cellspacing="0"><tr><td>
<![endif]-->
<table role="presentation" class="email-container" width="800" align="center"
       cellpadding="0" cellspacing="0"
       style="width:100%;max-width:800px;background-color:#FFFFFF;border:1px solid #DDE7E1">
<tr><td class="content-pad" style="padding:25px 32px;background-color:#006341">
<div style="font-family:'Segoe UI',Arial,Helvetica,sans-serif;font-size:11px;
letter-spacing:1.2px;font-weight:bold;color:#D9EEE2">SAFARICOM ETHIOPIA</div>
<h1 class="header-title" style="margin:7px 0 4px;font-family:'Segoe UI',Arial,
Helvetica,sans-serif;font-size:28px;line-height:1.2;font-weight:600;color:#FFFFFF">
Data &amp; Voice Business Performance Report</h1>
<div style="padding-top:8px;font-family:'Segoe UI',Arial,Helvetica,sans-serif;
font-size:13px;color:#D9EEE2">Reporting period: {escape(reporting_period)}</div>
</td></tr>
<tr><td class="content-pad" style="padding:22px 32px 8px">
<p style="margin:0;font-family:'Segoe UI',Arial,Helvetica,sans-serif;
font-size:15px;line-height:1.55;color:#303B35">Dear colleagues,</p>
<p style="margin:8px 0 0;font-family:'Segoe UI',Arial,Helvetica,sans-serif;
font-size:14px;line-height:1.6;color:#4D5A53">
Please find the executive summary and Data and Voice dashboards for the
reporting period below. The detailed Excel reports are attached for review.</p>
</td></tr>
<tr><td class="content-pad" style="padding:13px 27px 19px">
<h2 style="margin:0 0 7px;font-family:'Segoe UI',Arial,Helvetica,sans-serif;
font-size:18px;font-weight:600;color:#26332B">Executive KPI Summary</h2>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
       style="width:100%"><tr>{cards_html}</tr></table>
</td></tr>
<tr><td class="content-pad" style="padding:4px 32px 20px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
       style="width:100%;border-top:1px solid #E3EAE6">
<tr><td style="padding:19px 0 10px">
<h2 style="margin:0;font-family:'Segoe UI',Arial,Helvetica,sans-serif;
font-size:20px;font-weight:600;color:#006341">Data Performance Dashboard</h2>
</td></tr>
<tr><td style="padding:0">
<img class="dashboard-image" src="{data_src}" width="736"
     alt="Data Performance Dashboard"
     style="display:block;width:100%;max-width:736px;height:auto">
</td></tr></table>
</td></tr>
<tr><td class="content-pad" style="padding:0 32px 21px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
       style="width:100%;border-top:1px solid #E3EAE6">
<tr><td style="padding:19px 0 10px">
<h2 style="margin:0;font-family:'Segoe UI',Arial,Helvetica,sans-serif;
font-size:20px;font-weight:600;color:#006341">Voice Performance Dashboard</h2>
</td></tr>
<tr><td style="padding:0">
<img class="dashboard-image" src="{voice_src}" width="736"
     alt="Voice Performance Dashboard"
     style="display:block;width:100%;max-width:736px;height:auto">
</td></tr></table>
</td></tr>
<tr><td class="content-pad" style="padding:0 32px 18px">
<p style="margin:0;padding-top:13px;border-top:1px solid #E3EAE6;
font-family:'Segoe UI',Arial,Helvetica,sans-serif;font-size:12px;
line-height:1.55;color:#536159">{escape(REPORTING_NOTE)}</p>
</td></tr>
<tr><td class="content-pad footer-copy" style="padding:16px 32px 23px;
border-top:1px solid #DDE7E1;background-color:#F7F9F8;
font-family:'Segoe UI',Arial,Helvetica,sans-serif;font-size:13px;
line-height:1.55;color:#46534B">
Best regards,<br><strong style="color:#26332B">Mekin Jemal</strong><br>
Data Analyst | Pricing, Reporting &amp; CVM<br>
Safaricom Ethiopia Telecommunications PLC
</td></tr>
</table>
<!--[if mso]>
</td></tr></table>
<![endif]-->
</td></tr></table>
</body>
</html>"""


def build_plain_text_email(reporting_period: str, metrics: Mapping[str, Sequence[str]]) -> str:
    _validate_metrics(metrics)
    lines = [
        "Data & Voice Business Performance Report",
        f"Reporting period: {reporting_period}",
        "",
    ]
    for label in KPI_LABELS:
        data_value, voice_value = metrics[label]
        lines.extend(
            (label, f"Data: {data_value}", f"Voice: {voice_value}", "")
        )
    lines.extend(
        (
            REPORTING_NOTE,
            "",
            "Best regards,",
            "Mekin Jemal",
            "Data Analyst | Pricing, Reporting & CVM",
            "Safaricom Ethiopia Telecommunications PLC",
        )
    )
    return "\n".join(lines)
