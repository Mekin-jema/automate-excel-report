"""
Responsive Commercial Email Rendering Engine for KG Recharge Tracker.
=====================================================================
Constructs responsive, brand-aligned HTML and plain-text performance emails for
Safaricom Ethiopia commercial leadership.

Design & Structural Standards:
1. Safaricom Visual Identity:
   - Primary Green (#00B050), Deep Forest Green (#087A3B), Accent Red (#EF373E),
     Safaricom Yellow (#FFC20E), Pale Background (#EEF3EE).
   - Official Safaricom Email Footer brand banner.
2. Email Client Compatibility:
   - MSO conditional tables (<!--[if mso]>) ensuring rigid 800px layout in Microsoft Outlook.
   - Fluid responsive CSS media queries (@media only screen and (max-width:620px)).
   - Dual image rendering modes:
       a) CID Inline (cid:xxx) for production SMTP / Outlook dispatch.
       b) Base64 Data-URI (data:image/png;base64,...) for standalone browser previews.
3. Commercial Intelligence Features:
   - Milestone delta pills with dynamic color badges (positive/negative/neutral).
   - National Weekly Cohort progression table.
   - Regional Division Leaderboard table.
   - Executive commentary synthesizing trends and conversion milestones.
"""

from __future__ import annotations

import base64
import mimetypes
from datetime import datetime
from html import escape
from pathlib import Path
import pandas as pd

from config.settings import (
    ASSETS_DIR,
    BRAND_BACKGROUND,
    BRAND_DARK_SLATE,
    BRAND_DEEP_GREEN,
    BRAND_GREEN,
    BRAND_PALE_GREEN,
    BRAND_RED,
    BRAND_YELLOW,
    CHARTS_DIR,
    TARGETS_1RC,
    TARGETS_2RC,
)

FOOTER_IMAGE = ASSETS_DIR / "safaricom_email_footer.png"
FOOTER_IMAGE_CID = "safaricom_email_footer"
CHART_IMAGE_CID = "kg_performance_dashboard"
FONT_STACK = "'Segoe UI',Arial,Helvetica,sans-serif"

MOVEMENT_PILLS = {
    "positive": ("#E3F5E8", "#087A3B"),   # Soft green pill
    "negative": ("#FDE8E9", "#C62830"),   # Soft red pill
    "neutral": ("#EEF1EF", "#5A6A61"),    # Soft slate pill
}


def _image_source(
    image: str | Path,
    *,
    embed: bool,
    content_id: str | None = None,
) -> str:
    """Resolve image source URL for email delivery or browser preview."""
    if not embed:
        cid = content_id or Path(image).stem
        return f"cid:{cid}"
    
    path = Path(image)
    if not path.is_file():
        return ""
    mime_type, _ = mimetypes.guess_type(path.name)
    mime_type = mime_type or "image/png"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def render_delta_pill(val: float, target: float | None = None, is_pp_diff: bool = False) -> str:
    """Render a colored badge pill for milestone rates or target variances."""
    if is_pp_diff:
        # Percentage points difference
        if val > 0.001:
            bg, text = MOVEMENT_PILLS["positive"]
            label = f"+{val * 100:.1f} pp"
        elif val < -0.001:
            bg, text = MOVEMENT_PILLS["negative"]
            label = f"{val * 100:.1f} pp"
        else:
            bg, text = MOVEMENT_PILLS["neutral"]
            label = "0.0 pp"
    else:
        # Rate vs target
        rate_pct = val * 100
        if target is not None:
            diff = val - target
            if diff >= 0:
                bg, text = MOVEMENT_PILLS["positive"]
                label = f"{rate_pct:.1f}% (+{diff * 100:.1f}%)"
            else:
                bg, text = MOVEMENT_PILLS["negative"]
                label = f"{rate_pct:.1f}% ({diff * 100:.1f}%)"
        else:
            bg, text = MOVEMENT_PILLS["neutral"]
            label = f"{rate_pct:.1f}%"

    return (
        f'<span style="display:inline-block;padding:2px 8px;font-size:11px;'
        f'font-weight:600;border-radius:12px;background-color:{bg};color:{text};">'
        f'{label}</span>'
    )


def build_kg_performance_email(
    df_national: pd.DataFrame,
    df_regional: pd.DataFrame,
    chart_path: Path | None = None,
    as_of_date: datetime | None = None,
    embed_images: bool = False,
) -> str:
    """
    Construct the executive responsive HTML performance email.
    """

    if as_of_date is None:
        as_of_date = df_national["ACTIVATION_DATE"].max()
        if isinstance(as_of_date, str):
            as_of_date = datetime.strptime(as_of_date[:10], "%Y-%m-%d")

    date_str = as_of_date.strftime("%d %B %Y")

    # ---------------------------------------------------------
    # 1. FILTER FIRST AND SECOND RECHARGE
    # ---------------------------------------------------------
    is_1rc = df_national["REPORT_TYPE"].astype(str).str.contains(
        "1|first", case=False, regex=True
    )
    is_2rc = df_national["REPORT_TYPE"].astype(str).str.contains(
        "2|second", case=False, regex=True
    )

    df_1rc = df_national[is_1rc]
    df_2rc = df_national[is_2rc]

    if df_2rc.empty:
        df_2rc = df_national

    if df_1rc.empty:
        df_1rc = df_2rc

    # ---------------------------------------------------------
    # 2. OVERALL KPIs
    # ---------------------------------------------------------
    total_gross_adds = int(df_2rc["KG_BASE"].sum())

    avg_1rc_20d = df_1rc["PCT_20D"].dropna().mean()
    avg_2rc_20d = df_2rc["PCT_20D"].dropna().mean()
    avg_2rc_7d = df_2rc["PCT_7D"].dropna().mean()

    # ---------------------------------------------------------
    # 3. REGIONAL LEADERBOARD
    # ---------------------------------------------------------
    reg_summary = (
        df_regional.groupby("DS_EXEC_REGION", as_index=False)
        .agg({
            "KG_BASE": "sum",
            "PCT_3D": "mean",
            "PCT_7D": "mean",
            "PCT_16D": "mean",
            "PCT_20D": "mean",
        })
        .sort_values("PCT_20D", ascending=False)
        .reset_index(drop=True)
    )

    top_region = (
        reg_summary.iloc[0]["DS_EXEC_REGION"]
        if not reg_summary.empty else "N/A"
    )

    top_region_rate = (
        reg_summary.iloc[0]["PCT_20D"]
        if not reg_summary.empty else 0.0
    )

    lagging_region = (
        reg_summary.iloc[-1]["DS_EXEC_REGION"]
        if not reg_summary.empty else "N/A"
    )

    lagging_region_rate = (
        reg_summary.iloc[-1]["PCT_20D"]
        if not reg_summary.empty else 0.0
    )

    # ---------------------------------------------------------
    # 4. IMAGE SOURCES
    # ---------------------------------------------------------
    footer_src = _image_source(
        FOOTER_IMAGE,
        embed=embed_images,
        content_id=FOOTER_IMAGE_CID,
    )

    chart_src = ""

    if chart_path and chart_path.exists():
        chart_src = _image_source(
            chart_path,
            embed=embed_images,
            content_id=CHART_IMAGE_CID,
        )

    # ---------------------------------------------------------
    # 5. NATIONAL WEEKLY COHORT SUMMARY
    # ---------------------------------------------------------
    weekly_2rc = (
        df_2rc.groupby("WEEK", as_index=False)
        .agg({
            "ACTIVATION_DATE": "min",
            "KG_BASE": "sum",
            "PCT_3D": "mean",
            "PCT_7D": "mean",
            "PCT_16D": "mean",
            "PCT_20D": "mean",
        })
        .sort_values("ACTIVATION_DATE", ascending=False)
        .reset_index(drop=True)
    )

    weekly_rows_html = []

    for _, w in weekly_2rc.head(6).iterrows():

        wk_name = escape(str(w["WEEK"]))
        kg_b = f"{int(w['KG_BASE']):,}"

        p3 = f"{w['PCT_3D'] * 100:.1f}%"
        p7 = f"{w['PCT_7D'] * 100:.1f}%"
        p16 = f"{w['PCT_16D'] * 100:.1f}%"

        p20 = render_delta_pill(
            w["PCT_20D"],
            target=0.70,
        )

        weekly_rows_html.append(
            f"""
            <tr style="border-bottom:1px solid #E5E7EB;text-align:right;">
                <td style="padding:10px 12px;text-align:left;
                    font-weight:600;color:#111827;">
                    {wk_name}
                </td>

                <td style="padding:10px 12px;color:#374151;">
                    {kg_b}
                </td>

                <td style="padding:10px 12px;color:#374151;">
                    {p3}
                </td>

                <td style="padding:10px 12px;color:#374151;">
                    {p7}
                </td>

                <td style="padding:10px 12px;color:#374151;">
                    {p16}
                </td>

                <td style="padding:10px 12px;">
                    {p20}
                </td>
            </tr>
            """
        )

    weekly_rows_str = "\n".join(weekly_rows_html)

    # ---------------------------------------------------------
    # 6. REGIONAL LEADERBOARD TABLE
    # ---------------------------------------------------------
    reg_rows_html = []

    for rank, (_, r) in enumerate(reg_summary.iterrows(), start=1):

        reg_name = escape(str(r["DS_EXEC_REGION"]))
        kg_b = f"{int(r['KG_BASE']):,}"

        p3 = f"{r['PCT_3D'] * 100:.1f}%"
        p7 = f"{r['PCT_7D'] * 100:.1f}%"
        p16 = f"{r['PCT_16D'] * 100:.1f}%"

        p20 = render_delta_pill(
            r["PCT_20D"],
            target=0.70,
        )

        badge_bg = "#FEF3C7" if rank <= 3 else "#F3F4F6"
        badge_text = "#92400E" if rank <= 3 else "#4B5563"

        reg_rows_html.append(
            f"""
            <tr style="border-bottom:1px solid #E5E7EB;text-align:right;">

                <td style="padding:9px 12px;text-align:left;">

                    <span style="display:inline-block;
                        width:20px;height:20px;line-height:20px;
                        text-align:center;border-radius:50%;
                        background-color:{badge_bg};
                        color:{badge_text};font-size:11px;
                        font-weight:700;margin-right:8px;">
                        {rank}
                    </span>

                    <span style="font-weight:600;color:#111827;">
                        {reg_name}
                    </span>

                </td>

                <td style="padding:9px 12px;color:#374151;">
                    {kg_b}
                </td>

                <td style="padding:9px 12px;color:#374151;">
                    {p3}
                </td>

                <td style="padding:9px 12px;color:#374151;">
                    {p7}
                </td>

                <td style="padding:9px 12px;color:#374151;">
                    {p16}
                </td>

                <td style="padding:9px 12px;">
                    {p20}
                </td>

            </tr>
            """
        )

    reg_rows_str = "\n".join(reg_rows_html)

    # ---------------------------------------------------------
    # 7. OPTIONAL DASHBOARD CHART
    # ---------------------------------------------------------
    chart_html = ""

    if chart_src:
        chart_html = f"""
        <tr>
            <td style="padding:16px 32px 10px 32px;">

                <table role="presentation" width="100%"
                    cellpadding="0" cellspacing="0">

                    <tr>
                        <td style="padding:0;">

                            <img
                                class="dashboard-image"
                                src="{chart_src}"
                                width="736"
                                alt="KG Recharge Performance Dashboard"
                                style="display:block;width:100%;
                                max-width:736px;height:auto;
                                border:1px solid #DCE9DF;
                                border-radius:10px;margin:0 auto;"
                            >

                        </td>
                    </tr>

                </table>

            </td>
        </tr>
        """

    # ---------------------------------------------------------
    # 8. HTML EMAIL TEMPLATE
    # ---------------------------------------------------------
    html = f"""<!DOCTYPE html>
<html lang="en">

<head>

    <meta charset="utf-8">

    <meta name="viewport"
        content="width=device-width, initial-scale=1.0">

    <meta http-equiv="X-UA-Compatible" content="IE=edge">

    <title>Daily Gross Add (KG) — Cohort Performance</title>

    <style type="text/css">

        html, body {{
            margin: 0 !important;
            padding: 0 !important;
            width: 100% !important;
        }}

        body {{
            background-color: {BRAND_BACKGROUND};
            font-family: {FONT_STACK};
            -webkit-font-smoothing: antialiased;
        }}

        table {{
            border-collapse: collapse;
            mso-table-lspace: 0pt;
            mso-table-rspace: 0pt;
        }}

        img {{
            border: 0;
            outline: none;
            text-decoration: none;
            -ms-interpolation-mode: bicubic;
        }}

        @media only screen and (max-width:620px) {{

            .wrapper {{
                width: 100% !important;
            }}

            .content-pad {{
                padding-left: 16px !important;
                padding-right: 16px !important;
            }}

            .kpi-column {{
                display: block !important;
                width: 100% !important;
                margin-bottom: 12px !important;
            }}

            .dashboard-image {{
                width: 100% !important;
                height: auto !important;
            }}

            .footer-copy {{
                font-size: 13px !important;
                line-height: 1.55 !important;
            }}

        }}

    </style>


</head>

<body style="margin:0;padding:0;
    background-color:{BRAND_BACKGROUND};">

    <center>

        <!-- OUTER EMAIL BACKGROUND -->
        <table role="presentation"
            width="100%"
            cellpadding="0"
            cellspacing="0"
            style="width:100%;
            background-color:{BRAND_BACKGROUND};">

            <tr>

                <td align="center"
                    style="padding:28px 12px;">

                    <!--[if mso]>
                    <table role="presentation"
                        width="800"
                        align="center"
                        cellpadding="0"
                        cellspacing="0">
                    <tr>
                    <td>
                    <![endif]-->

                    <!-- MAIN EMAIL CONTAINER -->
                    <table
                        class="wrapper"
                        role="presentation"
                        width="800"
                        align="center"
                        cellpadding="0"
                        cellspacing="0"
                        style="width:100%;
                        max-width:800px;
                        border-collapse:separate;
                        background-color:#FFFFFF;
                        border:2px solid {BRAND_GREEN};
                        border-radius:16px;
                        overflow:hidden;">

                        <!-- SAFARICOM THREE-COLOR TOP ACCENT -->
                        <tr>

                            <td style="padding:0;
                                font-size:0;
                                line-height:0;">

                                <table role="presentation"
                                    width="100%"
                                    cellpadding="0"
                                    cellspacing="0"
                                    style="width:100%;">

                                    <tr>

                                        <td width="60%"
                                            height="6"
                                            style="width:60%;
                                            height:6px;
                                            background-color:{BRAND_GREEN};
                                            font-size:0;
                                            line-height:0;">
                                            &nbsp;
                                        </td>

                                        <td width="20%"
                                            height="6"
                                            style="width:20%;
                                            height:6px;
                                            background-color:{BRAND_YELLOW};
                                            font-size:0;
                                            line-height:0;">
                                            &nbsp;
                                        </td>

                                        <td width="20%"
                                            height="6"
                                            style="width:20%;
                                            height:6px;
                                            background-color:{BRAND_RED};
                                            font-size:0;
                                            line-height:0;">
                                            &nbsp;
                                        </td>

                                    </tr>

                                </table>

                            </td>

                        </tr>

                        <!-- HEADER BANNER -->
                        <tr>

                            <td style="background-color:#006341;
                                padding:28px 32px;
                                color:#FFFFFF;">

                                <table role="presentation" width="100%">

                                    <tr>

                                        <td>

                                            <span style="display:inline-block;
                                                padding:4px 10px;
                                                background-color:{BRAND_GREEN};
                                                color:#FFFFFF;
                                                font-size:11px;
                                                font-weight:700;
                                                border-radius:12px;
                                                letter-spacing:0.5px;
                                                text-transform:uppercase;">

                                                CVM &amp; Customer Onboarding

                                            </span>

                                            <h1 style="margin:10px 0 4px 0;
                                                font-size:22px;
                                                font-weight:700;
                                                letter-spacing:-0.3px;">

                                                Daily Gross Add (KG)
                                                Recharge Tracker

                                            </h1>

                                            <p style="margin:0;
                                                font-size:13px;
                                                color:#D5F5E3;">

                                                First and Second Recharge
                                                Cohort Conversion Intelligence

                                            </p>

                                        </td>

                                        <td align="right" valign="top">

                                            <div style="background-color:rgba(255,255,255,0.15);
                                                padding:8px 14px;
                                                border-radius:8px;
                                                font-size:12px;
                                                font-weight:600;
                                                color:#FFFFFF;
                                                white-space:nowrap;">

                                                As of: {date_str}

                                            </div>

                                        </td>

                                    </tr>

                                </table>

                            </td>

                        </tr>

                        <!-- KPI METRIC TILES -->
                        <tr>

                            <td class="content-pad"
                                style="padding:24px 32px 12px 32px;">

                                <table role="presentation"
                                    width="100%"
                                    style="table-layout:fixed;">

                                    <tr>

                                        <!-- TILE 1 -->
                                        <td class="kpi-column"
                                            width="25%"
                                            style="padding:0 6px;">

                                            <div style="background-color:#F9FAFB;
                                                border:1px solid #E5E7EB;
                                                border-top:3px solid {BRAND_GREEN};
                                                padding:14px;
                                                border-radius:8px;">

                                                <div style="font-size:11px;
                                                    color:#6B7280;
                                                    font-weight:600;
                                                    text-transform:uppercase;">

                                                    Gross Add SIMs

                                                </div>

                                                <div style="font-size:20px;
                                                    font-weight:700;
                                                    color:#111827;
                                                    margin-top:4px;">

                                                    {total_gross_adds:,}

                                                </div>

                                                <div style="font-size:11px;
                                                    color:#059669;
                                                    margin-top:4px;">

                                                    Active Onboarding Cohorts

                                                </div>

                                            </div>

                                        </td>

                                        <!-- TILE 2 -->
                                        <td class="kpi-column"
                                            width="25%"
                                            style="padding:0 6px;">

                                            <div style="background-color:#F9FAFB;
                                                border:1px solid #E5E7EB;
                                                border-top:3px solid {BRAND_DEEP_GREEN};
                                                padding:14px;
                                                border-radius:8px;">

                                                <div style="font-size:11px;
                                                    color:#6B7280;
                                                    font-weight:600;
                                                    text-transform:uppercase;">

                                                    1st RC 20D Avg

                                                </div>

                                                <div style="font-size:20px;
                                                    font-weight:700;
                                                    color:#111827;
                                                    margin-top:4px;">

                                                    {avg_1rc_20d * 100:.1f}%

                                                </div>

                                                <div style="font-size:11px;
                                                    color:#DC2626;
                                                    margin-top:4px;">

                                                    Target: 90.0%
                                                    ({(avg_1rc_20d - 0.90) * 100:.1f} pp)

                                                </div>

                                            </div>

                                        </td>

                                        <!-- TILE 3 -->
                                        <td class="kpi-column"
                                            width="25%"
                                            style="padding:0 6px;">

                                            <div style="background-color:#F9FAFB;
                                                border:1px solid #E5E7EB;
                                                border-top:3px solid {BRAND_RED};
                                                padding:14px;
                                                border-radius:8px;">

                                                <div style="font-size:11px;
                                                    color:#6B7280;
                                                    font-weight:600;
                                                    text-transform:uppercase;">

                                                    2nd RC 20D Avg

                                                </div>

                                                <div style="font-size:20px;
                                                    font-weight:700;
                                                    color:#111827;
                                                    margin-top:4px;">

                                                    {avg_2rc_20d * 100:.1f}%

                                                </div>

                                                <div style="font-size:11px;
                                                    color:#DC2626;
                                                    margin-top:4px;">

                                                    Target: 70.0%
                                                    ({(avg_2rc_20d - 0.70) * 100:.1f} pp)

                                                </div>

                                            </div>

                                        </td>

                                        <!-- TILE 4 -->
                                        <td class="kpi-column"
                                            width="25%"
                                            style="padding:0 6px;">

                                            <div style="background-color:#F9FAFB;
                                                border:1px solid #E5E7EB;
                                                border-top:3px solid {BRAND_YELLOW};
                                                padding:14px;
                                                border-radius:8px;">

                                                <div style="font-size:11px;
                                                    color:#6B7280;
                                                    font-weight:600;
                                                    text-transform:uppercase;">

                                                    Top Region

                                                </div>

                                                <div style="font-size:18px;
                                                    font-weight:700;
                                                    color:#111827;
                                                    margin-top:4px;
                                                    white-space:nowrap;
                                                    overflow:hidden;
                                                    text-overflow:ellipsis;">

                                                    {escape(str(top_region))}

                                                </div>

                                                <div style="font-size:11px;
                                                    color:#059669;
                                                    margin-top:4px;">

                                                    {top_region_rate * 100:.1f}%
                                                    20-Day Rate

                                                </div>

                                            </div>

                                        </td>

                                    </tr>

                                </table>

                            </td>

                        </tr>

                        <!-- DASHBOARD CHART -->
                        {chart_html}

                        <!-- NATIONAL COHORT TABLE -->
                        <tr>

                            <td class="content-pad"
                                style="padding:16px 32px;">

                                <h2 style="font-size:15px;
                                    color:#111827;
                                    margin:0 0 10px 0;
                                    font-weight:700;">

                                    <span style="display:inline-block;
                                        width:4px;
                                        height:16px;
                                        background-color:{BRAND_GREEN};
                                        border-radius:2px;
                                        margin-right:8px;">
                                    </span>

                                    National 2nd Recharge Milestone
                                    Progression (Recent Cohort Weeks)

                                </h2>

                                <table role="presentation"
                                    width="100%"
                                    style="font-size:12px;
                                    border:1px solid #E5E7EB;
                                    border-radius:8px;
                                    overflow:hidden;">

                                    <thead>

                                        <tr style="background-color:#F9FAFB;
                                            color:#4B5563;
                                            font-weight:700;
                                            border-bottom:2px solid #E5E7EB;
                                            text-align:right;">

                                            <th style="padding:10px 12px;
                                                text-align:left;">
                                                Cohort Week
                                            </th>

                                            <th style="padding:10px 12px;">
                                                KG Base
                                            </th>

                                            <th style="padding:10px 12px;">
                                                3 Days<br>
                                                <span style="font-size:10px;
                                                    color:#9CA3AF;
                                                    font-weight:normal;">
                                                    Target 15%
                                                </span>
                                            </th>

                                            <th style="padding:10px 12px;">
                                                7 Days<br>
                                                <span style="font-size:10px;
                                                    color:#9CA3AF;
                                                    font-weight:normal;">
                                                    Target 35%
                                                </span>
                                            </th>

                                            <th style="padding:10px 12px;">
                                                16 Days<br>
                                                <span style="font-size:10px;
                                                    color:#9CA3AF;
                                                    font-weight:normal;">
                                                    Target 50%
                                                </span>
                                            </th>

                                            <th style="padding:10px 12px;">
                                                20 Days<br>
                                                <span style="font-size:10px;
                                                    color:#9CA3AF;
                                                    font-weight:normal;">
                                                    Target 70%
                                                </span>
                                            </th>

                                        </tr>

                                    </thead>

                                    <tbody>
                                        {weekly_rows_str}
                                    </tbody>

                                </table>

                            </td>

                        </tr>

                        <!-- REGIONAL LEADERBOARD -->
                        <tr>

                            <td class="content-pad"
                                style="padding:16px 32px;">

                                <h2 style="font-size:15px;
                                    color:#111827;
                                    margin:0 0 10px 0;
                                    font-weight:700;">

                                    <span style="display:inline-block;
                                        width:4px;
                                        height:16px;
                                        background-color:{BRAND_DEEP_GREEN};
                                        border-radius:2px;
                                        margin-right:8px;">
                                    </span>

                                    Regional Second Recharge
                                    Performance Leaderboard

                                </h2>

                                <table role="presentation"
                                    width="100%"
                                    style="font-size:12px;
                                    border:1px solid #E5E7EB;
                                    border-radius:8px;
                                    overflow:hidden;">

                                    <thead>

                                        <tr style="background-color:#F9FAFB;
                                            color:#4B5563;
                                            font-weight:700;
                                            border-bottom:2px solid #E5E7EB;
                                            text-align:right;">

                                            <th style="padding:9px 12px;
                                                text-align:left;">
                                                Region
                                            </th>

                                            <th style="padding:9px 12px;">
                                                Total KG Base
                                            </th>

                                            <th style="padding:9px 12px;">
                                                3 Days Rate
                                            </th>

                                            <th style="padding:9px 12px;">
                                                7 Days Rate
                                            </th>

                                            <th style="padding:9px 12px;">
                                                16 Days Rate
                                            </th>

                                            <th style="padding:9px 12px;">
                                                20 Days Rate vs Target
                                            </th>

                                        </tr>

                                    </thead>

                                    <tbody>
                                        {reg_rows_str}
                                    </tbody>

                                </table>

                            </td>

                        </tr>

                        <!-- EXECUTIVE COMMENTARY -->
                        <tr>

                            <td class="content-pad"
                                style="padding:16px 32px 24px 32px;">

                                <div style="background-color:{BRAND_PALE_GREEN};
                                    border-left:4px solid {BRAND_GREEN};
                                    border-radius:4px;
                                    padding:16px 20px;">

                                    <h3 style="margin:0 0 6px 0;
                                        font-size:13px;
                                        font-weight:700;
                                        color:#006341;">

                                        Executive Commercial Summary
                                        &amp; Key Takeaways

                                    </h3>

                                    <ul style="margin:0;
                                        padding-left:18px;
                                        font-size:12px;
                                        color:#1F2937;
                                        line-height:1.6;">

                                        <li>
                                            <strong>
                                                Second Recharge Conversion:
                                            </strong>

                                            National 20-day second recharge
                                            conversion is currently averaging

                                            <strong>
                                                {avg_2rc_20d * 100:.1f}%
                                            </strong>

                                            across tracked cohorts against
                                            the target of 70.0%.
                                        </li>

                                        <li>
                                            <strong>
                                                Regional Standouts:
                                            </strong>

                                            <strong>{escape(str(top_region))}</strong>
                                            leads regional performance with

                                            <strong>
                                                {top_region_rate * 100:.1f}%
                                            </strong>

                                            20-day second recharge rate,
                                            while

                                            <strong>
                                                {escape(str(lagging_region))}
                                            </strong>

                                            ({lagging_region_rate * 100:.1f}%)
                                            requires targeted CVM nursery
                                            interventions.
                                        </li>

                                        <li>
                                            <strong>
                                                Workbook Attachment:
                                            </strong>

                                            Detailed cohort breakdown with
                                            Day 0 through Day 60 distributions
                                            and formula-linked regional tabs
                                            are available in the attached

                                            <em>KG_Recharge_Tracker.xlsx</em>.
                                        </li>

                                    </ul>

                                </div>

                            </td>

                        </tr>

                        <!-- UPDATED FOOTER SIGNATURE & BANNER -->
                        <tr>

                            <td class="content-pad footer-copy"
                                style="padding:22px 32px 26px;
                                font-family:{FONT_STACK};
                                font-size:13px;
                                line-height:1.6;
                                color:{BRAND_GREEN};">

                                <strong>

                                    Best regards,<br>

                                    <span style="font-size:15px;">
                                        Mekin Jemal
                                    </span><br>

                                    Graduate Management Trainee<br>

                                    Pricing, Reporting &amp; CVM<br>

                                    Safaricom Ethiopia Telecommunications PLC

                                </strong>

                                {
                                    "<img src='" + footer_src + "' "
                                    "width='736' "
                                    "alt='Safaricom Ethiopia' "
                                    "style='display:block;width:100%;"
                                    "max-width:736px;height:auto;"
                                    "padding-top:16px;'>"
                                    if footer_src else ""
                                }

                            </td>

                        </tr>

                    </table>
                    <!-- END MAIN EMAIL CONTAINER -->

                    <!--[if mso]>
                    </td>
                    </tr>
                    </table>
                    <![endif]-->

                </td>

            </tr>

        </table>

    </center>

</body>

</html>
"""

    return html
