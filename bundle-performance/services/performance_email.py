"""
Responsive Commercial Email Rendering Engine.
==============================================
Constructs responsive, brand-aligned HTML and plain-text performance emails for
Safaricom Ethiopia commercial leadership.

Design & Structural Standards:
1. Safaricom Visual Identity:
   - Primary Green (#00B050), Deep Forest Green (#087A3B), Accent Red (#EF373E),
     Safaricom Yellow (#FFC20E), Pale Background (#EEF3EE).
   - Official Safaricom Email Footer brand banner.
2. Email Client Compatibility:
   - MSO conditional tables (`<!--[if mso]>`) ensuring rigid 800px layout in Microsoft Outlook.
   - Fluid responsive CSS media queries (`@media only screen and (max-width:620px)`)
     collapsing two-column cards into single-column vertical flows on mobile devices.
   - Dual image rendering modes:
       a) CID Inline (`cid:xxx`) for production SMTP / Outlook dispatch.
       b) Base64 Data-URI (`data:image/png;base64,...`) for standalone browser previews.
3. Commercial Intelligence Features:
   - Day-over-Day (DoD) metric delta pills with dynamic color badges (growth/drop/neutral).
   - Natural language executive commentary synthesizing trends across Revenue, Subs, Qty, and Usage.
   - Side-by-side Data vs Voice KPI breakdown cards.
"""

from __future__ import annotations

import base64
import mimetypes
from collections.abc import Mapping, Sequence
from datetime import datetime
from html import escape
from pathlib import Path
from typing import TypedDict

# =====================================================================
# SAFARICOM ETHIOPIA BRAND CONSTANTS & STYLING PALETTE
# =====================================================================
BRAND_GREEN = "#00B050"          # Safaricom Primary Brand Green
BRAND_DEEP_GREEN = "#087A3B"     # Deep Executive Green for Voice & accents
BRAND_RED = "#EF373E"            # Safaricom Alert Red
BRAND_YELLOW = "#FFC20E"         # Safaricom Highlight Yellow
BRAND_PALE_GREEN = "#EAF6ED"     # Soft background for callout containers
BRAND_BACKGROUND = "#EEF3EE"     # Outer email canvas neutral tone

# Brand footer banner asset path and MIME Content-ID
FOOTER_IMAGE = Path(__file__).with_name("assets") / "safaricom_email_footer.png"
FOOTER_IMAGE_CID = "safaricom_email_footer"

# Universal font stack prioritizing Microsoft corporate desktop rendering
FONT_STACK = "'Segoe UI',Arial,Helvetica,sans-serif"

# Executive reporting metrics tracked in period KPI cards
KPI_LABELS = ("Revenue", "Purchase Quantity", "Usage")
KPI_ACCENTS = (BRAND_GREEN, BRAND_DEEP_GREEN, BRAND_YELLOW)
DAILY_METRIC_LABELS = ("Revenue", "Bundle Subscribers", "Purchase Quantity", "Usage")

# CSS badge styling for Day-over-Day movement indicators (background_color, text_color)
MOVEMENT_PILLS = {
    "positive": ("#E3F5E8", "#087A3B"),   # Soft green pill for positive commercial growth
    "negative": ("#FDE8E9", "#C62830"),   # Soft red pill for negative commercial drop
    "neutral": ("#EEF1EF", "#5A6A61"),    # Soft slate pill for unchanged / baseline days
}


class DailySummary(TypedDict):
    """
    Type-annotated container for Day-over-Day commercial performance summaries.
    
    Fields:
    -------
    reporting_date : str
        Target report date (e.g., '31 October 2026').
    comparison_date : str
        Baseline comparison date (typically T-1, e.g., '30 October 2026').
    data : Mapping[str, tuple[float | None, float | None]]
        Data portfolio indicators mapping metric name to (today_value, yesterday_value).
    voice : Mapping[str, tuple[float | None, float | None]]
        Voice portfolio indicators mapping metric name to (today_value, yesterday_value).
    """
    reporting_date: str
    comparison_date: str
    data: Mapping[str, tuple[float | None, float | None]]
    voice: Mapping[str, tuple[float | None, float | None]]


def format_report_period(start_date: str, end_date: str) -> str:
    """
    Format a start and end YYYYMMDD date pair into an executive-ready period string.

    Examples:
    - '20261019', '20261019' -> '19 October 2026'
    - '20261002', '20261031' -> '2–31 October 2026'
    - '20260928', '20261004' -> '28 Sep 2026 – 4 Oct 2026'
    """
    start = datetime.strptime(start_date, "%Y%m%d")
    end = datetime.strptime(end_date, "%Y%m%d")
    if start == end:
        return f"{start.day} {start.strftime('%B %Y')}"
    if start.year == end.year and start.month == end.month:
        return f"{start.day}–{end.day} {start.strftime('%B %Y')}"
    return f"{start.day} {start.strftime('%b %Y')} – {end.day} {end.strftime('%b %Y')}"


def _image_source(
    image: str | Path,
    *,
    embed: bool,
    content_id: str | None = None,
) -> str:
    """
    Resolve an image source URL based on the email delivery context.

    Parameters:
    -----------
    image : str | Path
        Filesystem path to the PNG image or default CID identifier.
    embed : bool
        If True, base64-encodes the image into a data-uri (used for offline HTML preview).
        If False, returns a 'cid:<content_id>' reference for MIME email transmission.
    content_id : str, optional
        Explicit Content-ID string to override default.

    Returns:
    --------
    str
        Either 'cid:...' or 'data:image/png;base64,...'.
    """
    if not embed:
        return f"cid:{content_id or image}"

    image_path = Path(image)
    if not image_path.is_file():
        raise FileNotFoundError(f"Dashboard image not found: {image_path}")
    media_type, _ = mimetypes.guess_type(image_path.name)
    if not media_type or not media_type.startswith("image/"):
        raise ValueError(f"Unsupported dashboard image type: {image_path.suffix}")
    encoded = base64.b64encode(image_path.read_bytes()).decode("ascii")
    return f"data:{media_type};base64,{encoded}"


def _validate_metrics(metrics: Mapping[str, Sequence[str]]) -> None:
    """Ensure all required executive KPI keys are present and contain both Data and Voice values."""
    missing = set(KPI_LABELS) - set(metrics)
    if missing:
        raise ValueError(f"Missing report KPI values: {', '.join(sorted(missing))}")
    invalid = [
        label
        for label in KPI_LABELS
        if len(metrics[label]) != 2 or not all(metrics[label])
    ]
    if invalid:
        raise ValueError(
            f"Each KPI requires Data and Voice values: {', '.join(invalid)}"
        )


def _daily_metric_value(label: str, value: float | None, *, voice: bool) -> str:
    """
    Format numeric values into standard Safaricom commercial presentation strings.

    Units:
    - Revenue: ETB (rounded integer)
    - Subscribers / Quantity: comma-separated count
    - Usage: GB (1 decimal place) for Data; Minutes (integer) for Voice.
    """
    if value is None:
        return "Not available"
    if label == "Revenue":
        return f"{value:,.0f} ETB"
    if label in ("Bundle Subscribers", "Purchase Quantity"):
        return f"{value:,.0f}"
    if label == "Usage":
        return f"{value:,.0f} minutes" if voice else f"{value:,.1f} GB"
    raise ValueError(f"Unknown daily KPI label: {label}")


def _daily_movement(
    current: float | None, previous: float | None
) -> tuple[str, str, str | None]:
    """
    Calculate Day-over-Day percentage change, direction, and badge color category.

    Returns:
    --------
    tuple[str, str, str | None]
        (movement_description, css_class_key, direction_name)
        e.g. ('Increased 4.2% vs previous day', 'positive', 'increased')
    """
    if current is None:
        return "Daily value unavailable", "neutral", None
    if previous is None:
        return "Previous-day comparison unavailable", "neutral", None
    if current == previous:
        return "No change vs previous day", "neutral", None
    direction = "Increased" if current > previous else "Decreased"
    css_class = "positive" if current > previous else "negative"
    if previous == 0:
        return (
            f"{direction} from 0 (percentage change unavailable)",
            css_class,
            direction.lower(),
        )
    change_pct = abs((current - previous) / abs(previous) * 100)
    return (
        f"{direction} {change_pct:,.1f}% vs previous day",
        css_class,
        direction.lower(),
    )


def _daily_insight(values: Mapping[str, tuple[float | None, float | None]]) -> str:
    """
    Synthesize natural language executive takeaways from daily movements.

    Example output:
    'Daily insight: revenue, purchase quantity increased; usage decreased.'
    """
    movements: dict[str, list[str]] = {"increased": [], "decreased": []}
    unavailable = []
    for label in DAILY_METRIC_LABELS:
        current, previous = values[label]
        _, _, direction = _daily_movement(current, previous)
        if direction:
            movements[direction].append(label.lower())
        elif current is None or previous is None:
            unavailable.append(label.lower())

    parts = []
    for direction in ("increased", "decreased"):
        if movements[direction]:
            labels = ", ".join(movements[direction])
            parts.append(f"{labels} {direction}")
    if parts:
        insight = "Daily insight: " + "; ".join(parts) + "."
    elif not unavailable:
        insight = "Daily insight: all measured indicators were unchanged."
    else:
        insight = "Daily insight: prior-day comparisons are unavailable for " + (
            ", ".join(unavailable) + "."
        )
    return insight


def _daily_service_html(
    service: str, values: Mapping[str, tuple[float | None, float | None]]
) -> str:
    """Render a responsive HTML table card for either Data or Voice daily metrics."""
    is_voice = service == "Voice"
    accent = BRAND_DEEP_GREEN if is_voice else BRAND_GREEN
    metric_rows = []
    for label in DAILY_METRIC_LABELS:
        current, previous = values[label]
        movement, css_class, _ = _daily_movement(current, previous)
        pill_bg, pill_fg = MOVEMENT_PILLS[css_class]
        metric_rows.append(
            f'<tr><td valign="middle" style="padding:11px 0;border-bottom:1px solid #E8F0EA">'
            f'<div style="padding-bottom:5px;font-family:{FONT_STACK};font-size:12px;'
            f'color:#52645A">{escape(label)}</div>'
            f'<span style="display:inline-block;padding:3px 10px;border-radius:12px;'
            f"background-color:{pill_bg};color:{pill_fg};font-family:{FONT_STACK};"
            f'font-size:11px;font-weight:bold;line-height:1.45">{escape(movement)}</span>'
            "</td>"
            '<td class="metric-value" align="right" valign="middle" '
            'style="padding:11px 0 11px 8px;border-bottom:1px solid #E8F0EA;'
            f"font-family:{FONT_STACK};font-size:17px;font-weight:bold;"
            f'color:#1B2B21;line-height:1.3">'
            f"{escape(_daily_metric_value(label, current, voice=is_voice))}</td></tr>"
        )

    return (
        '<td class="daily-service" width="50%" valign="top" '
        'style="width:50%;padding:6px">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
        'style="width:100%;border-collapse:separate;border:1px solid #DCE9DF;'
        f'border-top:4px solid {accent};border-radius:10px;background-color:#FFFFFF">'
        f'<tr><td style="padding:15px 16px 4px;font-family:{FONT_STACK};'
        f'font-size:17px;font-weight:bold;color:{BRAND_DEEP_GREEN}">'
        f"{escape(service)}</td></tr>"
        '<tr><td style="padding:0 16px 8px"><table role="presentation" '
        'width="100%" cellpadding="0" cellspacing="0">'
        + "".join(metric_rows)
        + "</table></td></tr></table></td>"
    )


def _daily_insight_callout(daily_summary: DailySummary) -> str:
    """Highlighted green callout container displaying synthesized insights for Data & Voice."""
    return (
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
        'style="width:100%;margin-top:8px"><tr><td style="padding:0 6px">'
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
        'style="width:100%;border-collapse:separate;background-color:#EAF6ED;'
        f'border-left:4px solid {BRAND_GREEN};border-radius:8px">'
        f'<tr><td style="padding:11px 14px 3px;font-family:{FONT_STACK};'
        'font-size:12px;line-height:1.55;color:#33453A">'
        f'{escape(_daily_insight(daily_summary["data"]))}</td></tr>'
        f'<tr><td style="padding:3px 14px 11px;font-family:{FONT_STACK};'
        'font-size:12px;line-height:1.55;color:#33453A">'
        f'{escape(_daily_insight(daily_summary["voice"]))}</td></tr>'
        "</table></td></tr></table>"
    )


def build_html_email(
    reporting_period: str,
    metrics: Mapping[str, Sequence[str]],
    data_dashboard: str | Path = "data_dashboard",
    voice_dashboard: str | Path = "voice_dashboard",
    *,
    embed_images: bool = False,
    daily_summary: DailySummary | None = None,
) -> str:
    """
    Assemble the complete responsive HTML commercial email document.

    Features included in markup:
    - Hidden preview preheader text for Outlook/mobile inbox snippet.
    - Tricolor brand stripe (Safaricom Green, Yellow, Red).
    - Executive title banner with linear gradient background.
    - Daily performance cards with movement badges.
    - Period KPI summary cards.
    - Inline dashboard chart visual panels for Data and Voice.
    - Official corporate email signature with footer banner image.

    Parameters:
    -----------
    reporting_period : str
        Human-readable date range label (e.g. '2–31 October 2026').
    metrics : Mapping[str, Sequence[str]]
        Dictionary containing ('Revenue', 'Purchase Quantity', 'Usage') paired with [data_str, voice_str].
    data_dashboard : str | Path
        Path or CID identifier for Data dashboard image.
    voice_dashboard : str | Path
        Path or CID identifier for Voice dashboard image.
    embed_images : bool, default False
        True to embed images as base64 (for browser previews); False to use cid: URLs for SMTP.
    daily_summary : DailySummary, optional
        Pre-aggregated DoD commercial metrics.

    Returns:
    --------
    str
        Full HTML5 email document string.
    """
    _validate_metrics(metrics)
    data_src = _image_source(data_dashboard, embed=embed_images)
    voice_src = _image_source(voice_dashboard, embed=embed_images)
    footer_src = _image_source(
        FOOTER_IMAGE,
        embed=embed_images,
        content_id=FOOTER_IMAGE_CID,
    )

    # Build the 3 period summary KPI cards (Revenue, Purchase Quantity, Usage)
    kpi_cards = []
    for label, accent in zip(KPI_LABELS, KPI_ACCENTS):
        data_value, voice_value = metrics[label]
        kpi_cards.append(
            '<td class="kpi-cell" width="33%" valign="top" '
            'style="width:33.333%;padding:5px">'
            '<table role="presentation" width="100%" cellpadding="0" '
            'cellspacing="0" style="width:100%;border-collapse:separate;'
            f"border:1px solid #DCE9DF;border-top:4px solid {accent};"
            'border-radius:10px;background-color:#FFFFFF">'
            '<tr><td style="padding:13px 14px 9px;color:#52645A;'
            f"font-family:{FONT_STACK};font-size:13px;"
            f'font-weight:bold">{escape(label)}</td></tr>'
            '<tr><td style="padding:0 14px 15px">'
            '<table role="presentation" width="100%" cellpadding="0" '
            'cellspacing="0"><tr>'
            '<td width="50%" valign="top" style="width:50%;padding-right:8px;'
            'border-right:1px solid #E5EEE7">'
            f'<div style="color:{BRAND_GREEN};font-family:{FONT_STACK};'
            'font-size:10px;font-weight:bold;letter-spacing:0.6px">DATA</div>'
            '<div class="kpi-value" style="padding-top:4px;color:#1B2B21;'
            f"font-family:{FONT_STACK};font-size:17px;"
            f'font-weight:bold;line-height:1.25">{escape(data_value)}</div></td>'
            '<td width="50%" valign="top" style="width:50%;padding-left:10px">'
            f'<div style="color:{BRAND_DEEP_GREEN};font-family:{FONT_STACK};'
            'font-size:10px;font-weight:bold;letter-spacing:0.6px">VOICE</div>'
            '<div class="kpi-value" style="padding-top:4px;color:#1B2B21;'
            f"font-family:{FONT_STACK};font-size:17px;"
            f'font-weight:bold;line-height:1.25">{escape(voice_value)}</div></td>'
            "</tr></table></td></tr></table></td>"
        )

    cards_html = "".join(kpi_cards)
    daily_summary_html = ""
    if daily_summary is not None:
        daily_summary_html = (
            '<tr><td class="content-pad" style="padding:17px 27px 12px">'
            f'<h2 style="margin:0 0 6px 6px;padding-left:10px;border-left:4px solid {BRAND_GREEN};'
            f"font-family:{FONT_STACK};"
            'font-size:20px;line-height:1.2;font-weight:600;color:#1B2B21">'
            "Daily Performance Summary</h2>"
            f'<p style="margin:0 0 10px 6px;font-family:{FONT_STACK};'
            'font-size:12px;line-height:1.5;color:#59685F">'
            f'Reporting date: <strong>{escape(daily_summary["reporting_date"])}</strong>'
            f' &nbsp;|&nbsp; Compared with: {escape(daily_summary["comparison_date"])}'
            '</p><table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
            'style="width:100%"><tr>'
            + _daily_service_html("Data", daily_summary["data"])
            + _daily_service_html("Voice", daily_summary["voice"])
            + "</tr></table>"
            + _daily_insight_callout(daily_summary)
            + "</td></tr>"
        )

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<title>Unlimited Data &amp; Voice Business Performance Report</title>
<style type="text/css">
html,body{{margin:0!important;padding:0!important;width:100%!important;}}
table{{border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt;}}
img{{border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic;}}
@media only screen and (max-width:620px){{
  .email-container{{width:100%!important;}}
  .outer-pad{{padding:8px 0!important;}}
  .content-pad{{padding-left:16px!important;padding-right:16px!important;}}
  .header-title{{font-size:24px!important;line-height:1.2!important;}}
  .kpi-cell,.daily-service{{display:block!important;width:100%!important;padding:4px 0!important;}}
  .dashboard-image{{width:100%!important;height:auto!important;}}
  .footer-copy{{font-size:13px!important;line-height:1.55!important;}}
}}
</style>
<!--[if mso]>
<style type="text/css">
.email-container{{width:800px!important;}}
.kpi-cell{{display:table-cell!important;width:33.333%!important;}}
</style>
<![endif]-->
</head>
<body style="margin:0;padding:0;background-color:{BRAND_BACKGROUND}">
<div style="display:none;font-size:1px;color:{BRAND_BACKGROUND};line-height:1px;
max-height:0;max-width:0;opacity:0;overflow:hidden">
Unlimited Data &amp; Voice Business Performance | {escape(reporting_period)}
</div>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
       style="width:100%;background-color:{BRAND_BACKGROUND}">
<tr><td class="outer-pad" align="center" style="padding:28px 12px">
<!--[if mso]>
<table role="presentation" width="800" align="center" cellpadding="0" cellspacing="0"><tr><td>
<![endif]-->
<table role="presentation" class="email-container" width="800" align="center"
       cellpadding="0" cellspacing="0"
       style="width:100%;max-width:800px;border-collapse:separate;background-color:#FFFFFF;
border:2px solid {BRAND_GREEN};border-radius:16px;overflow:hidden">
<tr><td style="padding:0;font-size:0;line-height:0">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%">
<tr>
<td width="60%" height="6" style="width:60%;height:6px;background-color:{BRAND_GREEN};font-size:0;line-height:0">&nbsp;</td>
<td width="20%" height="6" style="width:20%;height:6px;background-color:{BRAND_YELLOW};font-size:0;line-height:0">&nbsp;</td>
<td width="20%" height="6" style="width:20%;height:6px;background-color:{BRAND_RED};font-size:0;line-height:0">&nbsp;</td>
</tr></table>
</td></tr>
<tr><td class="content-pad" style="padding:30px 32px 28px;background-color:{BRAND_DEEP_GREEN};
background-image:linear-gradient(135deg,#065F2E 0%,#087A3B 55%,#16A044 100%)">
<div style="font-family:{FONT_STACK};font-size:11px;
letter-spacing:1.6px;font-weight:bold;color:#BFE8CB">SAFARICOM ETHIOPIA</div>
<h1 class="header-title" style="margin:9px 0 14px;font-family:{FONT_STACK};
font-size:30px;line-height:1.2;font-weight:600;color:#FFFFFF">
Unlimited Data &amp; Voice Business Performance Report</h1>
<span style="display:inline-block;padding:6px 14px;border-radius:16px;
background-color:#065F2E;border:1px solid #3CC15A;font-family:{FONT_STACK};
font-size:13px;color:#EAF6ED">Reporting period: {escape(reporting_period)}</span>
</td></tr>
<tr><td class="content-pad" style="padding:26px 32px 8px">
<p style="margin:0;font-family:{FONT_STACK};
font-size:16px;line-height:1.55;font-weight:600;color:{BRAND_GREEN}">Dear Team,</p>
<p style="margin:8px 0 0;font-family:{FONT_STACK};
font-size:14px;line-height:1.65;color:#52645A">
Please find an overview of Unlimited Data and Voice business performance,
including the latest daily summary, reporting-period KPIs, and dashboards.
The detailed Excel reports are attached for review.</p>
</td></tr>
{daily_summary_html}
<tr><td class="content-pad" style="padding:17px 27px 22px">
<h2 style="margin:0 0 10px 6px;padding-left:10px;border-left:4px solid {BRAND_GREEN};
font-family:{FONT_STACK};
font-size:20px;line-height:1.2;font-weight:600;color:#1B2B21">Reporting-Period KPI Summary</h2>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
       style="width:100%"><tr>{cards_html}</tr></table>
</td></tr>
<tr><td class="content-pad" style="padding:4px 32px 22px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
       style="width:100%;border-top:1px solid #DCE9DF">
<tr><td style="padding:22px 0 10px">
<h2 style="margin:0;padding-left:10px;border-left:4px solid {BRAND_GREEN};
font-family:{FONT_STACK};
font-size:20px;line-height:1.2;font-weight:600;color:{BRAND_DEEP_GREEN}">Data Performance Dashboard</h2>
<p style="margin:8px 0 4px;font-family:{FONT_STACK};
font-size:13px;line-height:1.55;color:#59685F">Provides an overview of Unlimited Data
performance, highlighting revenue, usage, subscriber activity, purchase trends,
and other relevant business indicators.</p>
</td></tr>
<tr><td style="padding:0">
<img class="dashboard-image" src="{data_src}" width="736"
     alt="Data Performance Dashboard"
     style="display:block;width:100%;max-width:736px;height:auto;
border:1px solid #DCE9DF;border-radius:10px">
</td></tr></table>
</td></tr>
<tr><td class="content-pad" style="padding:0 32px 26px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
       style="width:100%;border-top:1px solid #DCE9DF">
<tr><td style="padding:22px 0 10px">
<h2 style="margin:0;padding-left:10px;border-left:4px solid {BRAND_DEEP_GREEN};
font-family:{FONT_STACK};
font-size:20px;line-height:1.2;font-weight:600;color:{BRAND_DEEP_GREEN}">Voice Performance Dashboard</h2>
<p style="margin:8px 0 4px;font-family:{FONT_STACK};
font-size:13px;line-height:1.55;color:#59685F">Provides an overview of Unlimited Voice
performance, highlighting revenue, usage, subscriber activity, purchase trends,
and other relevant business indicators.</p>
</td></tr>
<tr><td style="padding:0">
<img class="dashboard-image" src="{voice_src}" width="736"
     alt="Voice Performance Dashboard"
     style="display:block;width:100%;max-width:736px;height:auto;
border:1px solid #DCE9DF;border-radius:10px">
</td></tr></table>
</td></tr>
<tr><td class="content-pad footer-copy" style="padding:22px 32px 26px;
border-top:2px solid {BRAND_GREEN};background-color:#FFFFFF;
font-family:{FONT_STACK};font-size:13px;
line-height:1.6;color:{BRAND_GREEN}">
Best regards,<br><strong style="color:{BRAND_GREEN};font-size:15px">Mekin Jemal</strong><br>
Graduate Management Trainee<br>Pricing, Reporting &amp; CVM<br>
Safaricom Ethiopia Telecommunications PLC
<img src="{footer_src}" width="736" alt="Safaricom Ethiopia"
     style="display:block;width:100%;max-width:736px;height:auto;padding-top:16px">
</td></tr>
</table>
<!--[if mso]>
</td></tr></table>
<![endif]-->
</td></tr></table>
</body>
</html>"""


def build_plain_text_email(
    reporting_period: str,
    metrics: Mapping[str, Sequence[str]],
    *,
    daily_summary: DailySummary | None = None,
) -> str:
    """
    Construct plain-text fallback content for email clients blocking HTML or terminal logs.

    Parameters:
    -----------
    reporting_period : str
        Human-readable date range label.
    metrics : Mapping[str, Sequence[str]]
        Dictionary containing ('Revenue', 'Purchase Quantity', 'Usage') paired with [data_str, voice_str].
    daily_summary : DailySummary, optional
        Pre-aggregated DoD commercial metrics.

    Returns:
    --------
    str
        Plain-text formatted summary email.
    """
    _validate_metrics(metrics)
    lines = [
        "Unlimited Data & Voice Business Performance Report",
        f"Reporting period: {reporting_period}",
        "",
        "Dear Team,",
        "",
        "Overview of Unlimited Data and Voice business performance follows. "
        "Detailed Excel reports are attached for review.",
        "",
    ]
    if daily_summary is not None:
        lines.extend(
            (
                "Daily Performance Summary",
                f"Reporting date: {daily_summary['reporting_date']}",
                f"Compared with: {daily_summary['comparison_date']}",
            )
        )
        for service, key, is_voice in (
            ("Data", "data", False),
            ("Voice", "voice", True),
        ):
            lines.append(f"{service}:")
            for label in DAILY_METRIC_LABELS:
                current, previous = daily_summary[key][label]
                movement, _, _ = _daily_movement(current, previous)
                lines.append(
                    f"  {label}: {_daily_metric_value(label, current, voice=is_voice)} "
                    f"({movement})"
                )
            lines.append(f"  {_daily_insight(daily_summary[key])}")
            lines.append("")
    lines.append(" Monthly Reporting-Period KPI Summary")
    for label in KPI_LABELS:
        data_value, voice_value = metrics[label]
        lines.extend((label, f"Data: {data_value}", f"Voice: {voice_value}", ""))
    lines.extend(
        (
            "Best regards,",
            "Mekin Jemal",
            "Graduate Management Trainee",
            "Pricing, Reporting & CVM",
            "Safaricom Ethiopia Telecommunications PLC",
        )
    )
    return "\n".join(lines)
