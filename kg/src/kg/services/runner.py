"""
End-to-End Orchestrator for KG Recharge Tracker Pipeline.
=========================================================
Coordinates:
1. Data extraction from Oracle or local DuckDB
2. High-fidelity 14-tab Excel report synthesis
3. Branded graphical chart rendering
4. Executive responsive HTML email construction
5. Desktop/mobile preview generation and optional SMTP dispatch
"""

from datetime import datetime
from pathlib import Path
from typing import Literal

from config.settings import DEFAULT_SOURCE, EXCEL_DIR
from kg.services.database import get_national_cohort_data, get_regional_cohort_data
from kg.services.email_service import save_html_previews, send_kg_email
from kg.services.email_template import build_kg_performance_email
from kg.services.excel_report import generate_kg_report
from kg.services.visual_charts import render_kg_dashboard_charts


def run_kg_pipeline(
    source: Literal["duckdb", "oracle"] = DEFAULT_SOURCE,
    send: bool = False,
    preview_only: bool = True,
    recipient: str | None = None,
    output_filename: str | None = None,
) -> dict:
    """
    Execute full KG performance reporting and email workflow.

    Parameters:
    -----------
    source : str
        Data source: 'duckdb' or 'oracle'.
    send : bool
        Whether to dispatch email via SMTP.
    preview_only : bool
        If True, skips SMTP delivery even if send is requested.
    recipient : str | None
        Override target recipient email address.
    output_filename : str | None
        Custom Excel report filename.

    Returns:
    --------
    dict
        Execution outcome with report paths and status.
    """
    print(f"\n{'='*70}")
    print(f"STARTING KG RECHARGE TRACKER PIPELINE [SOURCE: {source.upper()}]")
    print(f"{'='*70}\n")

    # 1. Fetch Datasets
    print("[1/5] Extracting cohort datasets...")
    df_national = get_national_cohort_data(source=source)
    df_regional = get_regional_cohort_data(source=source)

    if df_national.empty:
        raise ValueError("National cohort query returned no rows; aborting.")
    if df_regional.empty:
        raise ValueError("Regional cohort query returned no rows; aborting.")

    as_of_date = df_national["ACTIVATION_DATE"].max()
    if isinstance(as_of_date, str):
        as_of_date = datetime.strptime(as_of_date[:10], "%Y-%m-%d")

    date_str = as_of_date.strftime("%Y%m%d")
    report_name = output_filename or f"KG_Recharge_Tracker_{date_str}.xlsx"

    # 2. Generate 14-sheet Excel Report
    print("\n[2/5] Generating 14-sheet Excel workbook with formulas...")
    excel_path = generate_kg_report(
        df_national=df_national,
        df_regional=df_regional,
        output_filename=report_name,
    )

    # 3. Render Executive Dashboard Charts
    print("\n[3/5] Rendering executive graphical charts...")
    chart_path = render_kg_dashboard_charts(
        df_national=df_national,
        df_regional=df_regional,
        output_filename=f"kg_dashboard_{date_str}.png",
    )

    # 4. Generate Standalone HTML Previews
    print("\n[4/5] Building responsive executive HTML email previews...")
    preview_html = build_kg_performance_email(
        df_national=df_national,
        df_regional=df_regional,
        chart_path=chart_path,
        as_of_date=as_of_date,
        embed_images=True,
    )
    previews = save_html_previews(preview_html, filename_prefix=f"kg_performance_email_{date_str}")

    # 5. Email Dispatch
    email_sent = False
    if send and not preview_only:
        print("\n[5/5] Dispatching email via SMTP...")
        prod_html = build_kg_performance_email(
            df_national=df_national,
            df_regional=df_regional,
            chart_path=chart_path,
            as_of_date=as_of_date,
            embed_images=False,
        )
        subject = f"Daily Gross Add (KG) — Cohort Performance Report [{as_of_date.strftime('%d %b %Y')}]"
        email_sent = send_kg_email(
            subject=subject,
            html_body=prod_html,
            attachment_paths=[excel_path],
            chart_path=chart_path,
            recipient_override=recipient,
        )
    else:
        print("\n[5/5] Preview mode active; SMTP send skipped.")

    print(f"\n{'='*70}")
    print("KG RECHARGE TRACKER PIPELINE COMPLETED SUCCESSFULLY")
    print(f"Excel Report: {excel_path}")
    print(f"Desktop Preview: {previews['desktop']}")
    print(f"Mobile Preview:  {previews['mobile']}")
    print(f"{'='*70}\n")

    return {
        "status": "success",
        "excel_path": excel_path,
        "chart_path": chart_path,
        "previews": previews,
        "email_sent": email_sent,
        "as_of_date": as_of_date,
    }

