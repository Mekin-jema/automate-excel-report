from datetime import date, datetime, timedelta
from pathlib import Path
import json
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from services.database import load_bundle_data
from services.report import create_excel, create_pdf
from services.email_service import send_outlook_email, is_outlook_running
from services.scheduler import (
    load_schedule_config,
    save_schedule_config,
    load_schedule_history,
    run_daily_report_job,
    start_scheduler_thread,
    stop_scheduler_thread,
    is_scheduler_running,
    generate_windows_batch_script,
)

BASE_DIR = Path(__file__).resolve().parent

# ============================================================
# PAGE CONFIGURATION & STYLING
# ============================================================
st.set_page_config(
    page_title="Safaricom | Bundle Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS for modern telecom executive dashboard aesthetic
st.markdown("""
<style>
    /* Metric Cards */
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 18px 22px;
        border-radius: 12px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        border-left: 5px solid #00A859;
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
    }
    div[data-testid="stMetricLabel"] {
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        color: #4a5568 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem !important;
        font-weight: 700 !important;
        color: #1a202c !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 48px;
        white-space: pre-wrap;
        border-radius: 8px 8px 0px 0px;
        padding: 10px 20px;
        font-weight: 600;
    }

    /* Badges */
    .badge-success {
        background-color: #e6f4ea;
        color: #137333;
        padding: 4px 12px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.85rem;
        display: inline-block;
    }
    .badge-info {
        background-color: #e8f0fe;
        color: #1a73e8;
        padding: 4px 12px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.85rem;
        display: inline-block;
    }

    /* Cards */
    .dashboard-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03);
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# SIDEBAR CONTROLS
# ============================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/cd/Safaricom_logo.svg/320px-Safaricom_logo.svg.png", width=180) if False else st.markdown("### 🟢 **Safaricom Ethiopia**")
    st.subheader("Data & Query Controls")

    # 1. Data Source Selection
    data_source_mode = st.radio(
        "Database Source",
        options=["DuckDB (Sample Mode)", "Oracle (Production)"],
        index=0,
        help="Use DuckDB for offline development and testing with sample data. Use Oracle for live production connection.",
    )
    selected_source = "duckdb" if "DuckDB" in data_source_mode else "oracle"

    # 2. Date Presets & Inputs
    st.markdown("#### Reporting Date Range")
    preset = st.selectbox(
        "Quick Date Preset",
        options=["Sample Date (2026-10-19)", "Yesterday", "Last 7 Days", "Today", "Custom Range"],
        index=0,
    )

    today = date.today()
    if preset == "Sample Date (2026-10-19)":
        start_date = date(2026, 10, 19)
        end_date = date(2026, 10, 19)
    elif preset == "Yesterday":
        start_date = today - timedelta(days=1)
        end_date = today - timedelta(days=1)
    elif preset == "Last 7 Days":
        start_date = today - timedelta(days=7)
        end_date = today - timedelta(days=1)
    elif preset == "Today":
        start_date = today
        end_date = today
    else:
        start_date = st.date_input("Start Date", value=today - timedelta(days=1))
        end_date = st.date_input("End Date", value=today)

    run_queries = st.button("🚀 Run SQL Queries", type="primary", use_container_width=True)

    st.divider()

    # Outlook Desktop Status
    outlook_active = is_outlook_running()
    if outlook_active:
        st.markdown("🟢 **Outlook:** Desktop Client Active")
    else:
        st.markdown("⚪ **Outlook:** Client Not Running")

    st.caption("Bundle Performance Reporting Engine v2.0")


# ============================================================
# STATE INITIALIZATION & DATA LOADING
# ============================================================
if "usage" not in st.session_state and not run_queries:
    # Auto-load initial sample data for instant preview
    try:
        start_k = start_date.strftime("%Y%m%d")
        end_k = end_date.strftime("%Y%m%d")
        usage, purchases, actual_src = load_bundle_data(start_k, end_k, source=selected_source)
        st.session_state.usage = usage
        st.session_state.purchases = purchases
        st.session_state.actual_source = actual_src
        st.session_state.start_key = start_k
        st.session_state.end_key = end_k
        st.session_state.excel = str(create_excel(usage, purchases, start_k, end_k))
        st.session_state.pdf = str(create_pdf(usage, purchases, start_k, end_k))
    except Exception:
        pass

if run_queries:
    if start_date > end_date:
        st.sidebar.error("Error: Start date cannot be greater than End date.")
    else:
        start_k = start_date.strftime("%Y%m%d")
        end_k = end_date.strftime("%Y%m%d")
        try:
            with st.spinner("Executing SQL queries and preparing reports..."):
                usage, purchases, actual_src = load_bundle_data(start_k, end_k, source=selected_source)
                st.session_state.usage = usage
                st.session_state.purchases = purchases
                st.session_state.actual_source = actual_src
                st.session_state.start_key = start_k
                st.session_state.end_key = end_k
                st.session_state.excel = str(create_excel(usage, purchases, start_k, end_k))
                st.session_state.pdf = str(create_pdf(usage, purchases, start_k, end_k))
            st.sidebar.success(f"Data loaded from {actual_src}")
        except Exception as error:
            st.sidebar.error(f"Execution failed: {error}")


# ============================================================
# HEADER
# ============================================================
head_col1, head_col2 = st.columns([3, 1])
with head_col1:
    st.title("📊 Bundle Performance Dashboard")
    st.caption("Daily Tracking, Subscriber Metrics, Usage Analytics & Automated Scheduling")
with head_col2:
    current_src = st.session_state.get("actual_source", "DuckDB (Sample Mode)")
    badge_class = "badge-success" if "DuckDB" in current_src else "badge-info"
    st.markdown(f"<div style='text-align: right; padding-top: 15px;'><span class='{badge_class}'>● {current_src}</span></div>", unsafe_allow_html=True)

# Tabs
tab_report, tab_schedule, tab_data = st.tabs([
    "📈 Daily Performance Report",
    "⏰ Schedule Dashboard",
    "🔍 Data & Query Inspector",
])


# ============================================================
# TAB 1: DAILY PERFORMANCE REPORT
# ============================================================
with tab_report:
    if "purchases" in st.session_state and "usage" in st.session_state:
        purchases = st.session_state.purchases
        usage = st.session_state.usage
        start_k = st.session_state.start_key
        end_k = st.session_state.end_key

        # Calculate KPIs
        def get_metric_sum(metric_code):
            vals = pd.to_numeric(
                purchases.loc[purchases["METRIC"].astype(str) == metric_code, "VALUE"],
                errors="coerce"
            )
            return vals.sum()

        total_subs = get_metric_sum("A.SUBS")
        total_quantity = get_metric_sum("B.QUANTITY")
        total_revenue = get_metric_sum("C.REVENUE")
        total_mb = pd.to_numeric(usage["VALUE"], errors="coerce").sum()
        total_gb = total_mb / 1024.0
        arpu = (total_revenue / total_subs) if total_subs > 0 else 0.0

        st.markdown(f"##### Reporting Window: `{start_k}` to `{end_k}`")

        # 1. KPI CARDS
        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
        kpi1.metric("Total Revenue", f"{total_revenue:,.2f} ETB")
        kpi2.metric("Subscribers (Unique)", f"{total_subs:,.0f}")
        kpi3.metric("Purchase Transactions", f"{total_quantity:,.0f}")
        kpi4.metric("Data Traffic", f"{total_gb:,.1f} GB")
        kpi5.metric("ARPU (Per Sub)", f"{arpu:,.2f} ETB")

        st.divider()

        # 2. CHARTS SECTION
        ch_col1, ch_col2 = st.columns(2)

        with ch_col1:
            st.subheader("Revenue by Category")
            rev_df = purchases[purchases["METRIC"] == "C.REVENUE"].copy()
            if not rev_df.empty:
                rev_df["VALUE"] = pd.to_numeric(rev_df["VALUE"], errors="coerce")
                cat_summary = rev_df.groupby("BUNDLE_CATEGORY", as_index=False)["VALUE"].sum()
                cat_summary = cat_summary.sort_values(by="VALUE", ascending=False)

                fig_cat = px.pie(
                    cat_summary,
                    names="BUNDLE_CATEGORY",
                    values="VALUE",
                    hole=0.45,
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                    title="Revenue Contribution by Bundle Category",
                )
                fig_cat.update_traces(textposition="inside", textinfo="percent+label")
                fig_cat.update_layout(showlegend=False, margin=dict(t=40, b=20, l=20, r=20))
                st.plotly_chart(fig_cat, use_container_width=True)

        with ch_col2:
            st.subheader("Top Revenue-Generating Bundles")
            if not rev_df.empty:
                bundle_rev = rev_df.groupby("BUNDLE_NAME", as_index=False)["VALUE"].sum()
                bundle_rev = bundle_rev.sort_values(by="VALUE", ascending=True).tail(8)

                fig_top_rev = px.bar(
                    bundle_rev,
                    x="VALUE",
                    y="BUNDLE_NAME",
                    orientation="h",
                    title="Top Bundles by Revenue (ETB)",
                    labels={"VALUE": "Revenue (ETB)", "BUNDLE_NAME": ""},
                    color="VALUE",
                    color_continuous_scale="Greens",
                )
                fig_top_rev.update_layout(margin=dict(t=40, b=20, l=20, r=20), coloraxis_showscale=False)
                st.plotly_chart(fig_top_rev, use_container_width=True)

        st.subheader("Data Usage Volume by Bundle (GB)")
        usage_chart = usage.copy()
        usage_chart["VALUE"] = pd.to_numeric(usage_chart["VALUE"], errors="coerce")
        usage_chart["GB"] = usage_chart["VALUE"] / 1024.0
        grouped_usage = usage_chart.groupby("BUNDLE_NAME", as_index=False)["GB"].sum().sort_values(by="GB", ascending=False)

        fig_usage = px.bar(
            grouped_usage,
            x="BUNDLE_NAME",
            y="GB",
            title="Data Consumed by Bundle (GB)",
            labels={"GB": "Usage (GB)", "BUNDLE_NAME": "Bundle"},
            color="GB",
            color_continuous_scale="Viridis",
        )
        fig_usage.update_layout(xaxis_tickangle=-40, margin=dict(t=40, b=60, l=20, r=20))
        st.plotly_chart(fig_usage, use_container_width=True)

        st.divider()

        # 3. EXPORT & ACTIONS BAR
        st.subheader("Export & Report Dispatch")
        exp_col1, exp_col2, exp_col3 = st.columns(3)

        with exp_col1:
            with open(st.session_state.excel, "rb") as f:
                excel_bytes = f.read()
            st.download_button(
                "📥 Download Excel Report (.xlsx)",
                excel_bytes,
                file_name=f"Bundle_Performance_{start_k}_{end_k}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with exp_col2:
            with open(st.session_state.pdf, "rb") as f:
                pdf_bytes = f.read()
            st.download_button(
                "📄 Download PDF Summary (.pdf)",
                pdf_bytes,
                file_name=f"Bundle_Performance_{start_k}_{end_k}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        with exp_col3:
            if st.button("✉️ Send Report via Outlook Now", use_container_width=True):
                cfg = load_schedule_config()
                if not is_outlook_running():
                    st.warning("⚠️ Microsoft Outlook is not currently running. Please open Outlook Desktop to send.")
                else:
                    try:
                        res = send_outlook_email(
                            subject=f"Bundle Performance Report [{start_k} - {end_k}]",
                            body="Hello,\n\nPlease find attached the daily Bundle Performance Report.\n\nBest regards,\nMekin Jemal",
                            attachments=[st.session_state.excel, st.session_state.pdf],
                            to=cfg.get("to_emails"),
                            cc=cfg.get("cc_emails"),
                        )
                        st.success("✅ Email successfully sent to recipients!")
                    except Exception as err:
                        st.error(f"Failed to send email: {err}")

    else:
        st.info("Select a date range and click 'Run SQL Queries' in the sidebar to view the report.")


# ============================================================
# TAB 2: SCHEDULE DASHBOARD
# ============================================================
with tab_schedule:
    st.subheader("⏰ Automated Daily Report Scheduler")
    st.caption("Configure automated report generation, Outlook distribution lists, and background daemon execution.")

    cfg = load_schedule_config()
    history = load_schedule_history(limit=15)
    running_status = is_scheduler_running()

    # Scheduler Status Banner
    stat_col1, stat_col2, stat_col3, stat_col4 = st.columns(4)
    with stat_col1:
        st.metric("Daemon Status", "RUNNING 🟢" if running_status else "STOPPED ⚪")
    with stat_col2:
        st.metric("Configured Daily Time", cfg.get("run_time", "08:00"))
    with stat_col3:
        st.metric("Last Run Status", cfg.get("last_run_status", "N/A"))
    with stat_col4:
        st.metric("Last Run Time", cfg.get("last_run_timestamp", "Never") or "Never")

    # Start / Stop Buttons
    btn_c1, btn_c2, btn_c3 = st.columns([1, 1, 2])
    with btn_c1:
        if st.button("▶️ Start Scheduler", disabled=running_status, use_container_width=True):
            start_scheduler_thread()
            st.rerun()
    with btn_c2:
        if st.button("⏹️ Stop Scheduler", disabled=not running_status, use_container_width=True):
            stop_scheduler_thread()
            st.rerun()
    with btn_c3:
        if st.button("⚡ Run Scheduled Job Now (Manual Trigger)", type="primary", use_container_width=True):
            with st.spinner("Executing scheduled job pipeline..."):
                job_res = run_daily_report_job(trigger_type="Dashboard Manual")
                if job_res["status"] == "SUCCESS":
                    st.success(f"Job completed successfully in {job_res['duration_sec']}s! {job_res['message']}")
                else:
                    st.error(f"Job encountered an issue: {job_res['message']}")
            st.rerun()

    st.divider()

    # Configuration Form
    st.markdown("#### ⚙️ Schedule Configuration")
    with st.form("schedule_config_form"):
        f_col1, f_col2 = st.columns(2)

        with f_col1:
            new_enabled = st.checkbox("Enable Daily Automation", value=cfg.get("enabled", True))
            new_time = st.text_input("Daily Run Time (HH:MM 24h format)", value=cfg.get("run_time", "08:00"))
            new_date_mode = st.selectbox(
                "Report Date Scope",
                options=["yesterday", "today", "last_7_days", "sample_date"],
                index=["yesterday", "today", "last_7_days", "sample_date"].index(cfg.get("date_mode", "yesterday")),
                help="'yesterday' reports previous day data. 'sample_date' uses 20261019.",
            )
            new_source = st.selectbox(
                "Automation Data Source",
                options=["duckdb", "oracle"],
                index=0 if cfg.get("data_source") == "duckdb" else 1,
            )

        with f_col2:
            new_sender = st.text_input("Sender Email", value=cfg.get("sender_email", "Mekin.Jemal@safaricom.et"))
            new_to = st.text_area("Direct Recipients (To)", value=cfg.get("to_emails", ""), height=70)
            new_cc = st.text_area("CC Recipients", value=cfg.get("cc_emails", ""), height=70)

        opt_c1, opt_c2, opt_c3 = st.columns(3)
        with opt_c1:
            new_send_email = st.checkbox("Automatically Send via Outlook", value=cfg.get("send_email", False))
        with opt_c2:
            new_excel = st.checkbox("Attach Excel Report", value=cfg.get("attach_excel", True))
        with opt_c3:
            new_pdf = st.checkbox("Attach PDF Summary", value=cfg.get("attach_pdf", True))

        save_btn = st.form_submit_button("💾 Save Schedule Settings", type="primary")

        if save_btn:
            updated_cfg = {
                "enabled": new_enabled,
                "run_time": new_time.strip(),
                "date_mode": new_date_mode,
                "data_source": new_source,
                "sender_email": new_sender.strip(),
                "to_emails": new_to.strip(),
                "cc_emails": new_cc.strip(),
                "send_email": new_send_email,
                "attach_excel": new_excel,
                "attach_pdf": new_pdf,
                "last_run_timestamp": cfg.get("last_run_timestamp"),
                "last_run_status": cfg.get("last_run_status"),
            }
            save_schedule_config(updated_cfg)
            generate_windows_batch_script()
            st.success("✅ Schedule configuration saved!")
            st.rerun()

    st.divider()

    # Schedule History Table
    st.markdown("#### 📜 Execution Run History & Logs")
    if history:
        history_df = pd.DataFrame(history)
        cols_to_show = ["timestamp", "trigger", "reporting_period", "data_source", "status", "email_sent", "duration_sec", "message"]
        existing_cols = [c for c in cols_to_show if c in history_df.columns]
        st.dataframe(
            history_df[existing_cols],
            use_container_width=True,
            column_config={
                "timestamp": "Execution Time",
                "trigger": "Trigger Type",
                "reporting_period": "Reporting Scope",
                "data_source": "Data Source",
                "status": "Status",
                "email_sent": "Email Sent",
                "duration_sec": "Duration (s)",
                "message": "Message / Details",
            },
            hide_index=True,
        )
    else:
        st.info("No execution history recorded yet. Click 'Run Scheduled Job Now' above to test.")

    st.divider()

    # Windows Task Scheduler Integration
    st.markdown("#### 🖥️ Windows Task Scheduler Integration (24/7 Unattended Run)")
    st.write(
        "To run reports automatically every day even when your browser or IDE is closed, "
        "you can set up a Windows Task Scheduler task using the pre-generated batch script:"
    )
    batch_script_path = generate_windows_batch_script()
    st.code(f'"{batch_script_path}"', language="bat")
    with st.expander("Show Windows Task Scheduler Setup Steps"):
        st.markdown("""
        1. Press `Win + R`, type `taskschd.msc`, and press Enter.
        2. Click **Create Basic Task** in the Actions panel.
        3. Name it `Safaricom Daily Bundle Report`.
        4. Trigger: Select **Daily**, then specify the time (e.g. `08:00 AM`).
        5. Action: Select **Start a Program**.
        6. Program/Script: Browse and select:
           `{}`
        7. Finish the wizard. Your report will now run automatically every morning!
        """.format(batch_script_path))


# ============================================================
# TAB 3: DATA & QUERY INSPECTOR
# ============================================================
with tab_data:
    st.subheader("🔍 Data & Query Inspector")

    sub_t1, sub_t2, sub_t3 = st.tabs(["Bundle Purchases Table", "Bundle Usage Table", "SQL Query Source"])

    with sub_t1:
        if "purchases" in st.session_state:
            st.dataframe(st.session_state.purchases, use_container_width=True)
            csv_p = st.session_state.purchases.to_csv(index=False).encode("utf-8")
            st.download_button("Download Purchases CSV", csv_p, "bundle_purchases.csv", "text/csv")
        else:
            st.info("Run queries to inspect purchase data.")

    with sub_t2:
        if "usage" in st.session_state:
            st.dataframe(st.session_state.usage, use_container_width=True)
            csv_u = st.session_state.usage.to_csv(index=False).encode("utf-8")
            st.download_button("Download Usage CSV", csv_u, "bundle_usage.csv", "text/csv")
        else:
            st.info("Run queries to inspect usage data.")

    with sub_t3:
        st.markdown("##### 📄 `sql/bundle_purchase.sql`")
        purchase_sql_code = (BASE_DIR / "sql" / "bundle_purchase.sql").read_text(encoding="utf-8")
        st.code(purchase_sql_code, language="sql")

        st.markdown("##### 📄 `sql/bundle_usage.sql`")
        usage_sql_code = (BASE_DIR / "sql" / "bundle_usage.sql").read_text(encoding="utf-8")
        st.code(usage_sql_code, language="sql")
