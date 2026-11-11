"""
Executive Visual Chart Rendering Service for KG Recharge Tracker.
=================================================================
Renders high-resolution, Safaricom brand-aligned graphical charts:
1. National Weekly Milestone Cohort Trends (1RC vs 2RC vs Targets)
2. Regional Second Recharge 20-Day Conversion Leaderboard vs 70% Target
3. Daily Cumulative Conversion Curve (Day 0 to Day 20)
"""

from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patheffects as path_effects
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
from PIL import Image

from config.settings import (
    BRAND_DEEP_GREEN,
    BRAND_GREEN,
    BRAND_RED,
    BRAND_YELLOW,
    CHARTS_DIR,
    TARGETS_1RC,
    TARGETS_2RC,
)


def render_kg_dashboard_charts(
    df_national: pd.DataFrame,
    df_regional: pd.DataFrame,
    output_filename: str = "kg_performance_dashboard.png",
) -> Path:
    """
    Render 2-panel executive dashboard chart combining:
    - Panel 1: National Weekly Milestone Progression (WK1..WKn)
    - Panel 2: Regional 20-Day Second Recharge Conversion vs 70% Target
    """
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)
    target_path = CHARTS_DIR / output_filename

    # Prepare National Weekly Data
    df_2rc = df_national[df_national["REPORT_TYPE"] == "2nd Purchase"].copy()
    weekly_agg = df_2rc.groupby("WEEK", as_index=False).agg({
        "ACTIVATION_DATE": "min",
        "KG_BASE": "sum",
        "PCT_3D": "mean",
        "PCT_7D": "mean",
        "PCT_16D": "mean",
        "PCT_20D": "mean",
    }).sort_values("ACTIVATION_DATE").reset_index(drop=True)

    # Filter out recent incomplete weeks for 20D if na or 0
    weekly_agg = weekly_agg[weekly_agg["PCT_3D"] > 0]

    # Prepare Regional Data (latest 4-week average or overall mean)
    reg_agg = df_regional.groupby("DS_EXEC_REGION", as_index=False).agg({
        "KG_BASE": "sum",
        "PCT_3D": "mean",
        "PCT_7D": "mean",
        "PCT_16D": "mean",
        "PCT_20D": "mean",
    }).sort_values("PCT_20D", ascending=True).reset_index(drop=True)

    # Figure Setup: 2 panels side by side (16:7 aspect ratio)
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6), dpi=150)
    fig.patch.set_facecolor("#FFFFFF")

    # =========================================================================
    # Panel 1: National Weekly Second Recharge Milestone Progression
    # =========================================================================
    weeks = weekly_agg["WEEK"].tolist()
    x = np.arange(len(weeks))

    ax1.set_facecolor("#FAFAFA")
    ax1.grid(True, linestyle="--", alpha=0.4, color="#D0D7DE", zorder=1)

    # Plot Milestones
    ax1.plot(x, weekly_agg["PCT_20D"] * 100, marker="o", color=BRAND_DEEP_GREEN, linewidth=2.5, label="20-Day Rate", zorder=3)
    ax1.plot(x, weekly_agg["PCT_16D"] * 100, marker="s", color=BRAND_GREEN, linewidth=2, label="16-Day Rate", zorder=3)
    ax1.plot(x, weekly_agg["PCT_7D"] * 100, marker="^", color=BRAND_YELLOW, linewidth=2, label="7-Day Rate", zorder=3)
    ax1.plot(x, weekly_agg["PCT_3D"] * 100, marker="d", color="#3B82F6", linewidth=1.5, label="3-Day Rate", zorder=3)

    # Target Reference Lines
    ax1.axhline(70, color=BRAND_RED, linestyle=":", linewidth=1.8, label="20-Day Target (70%)", zorder=2)
    ax1.axhline(35, color="#D97706", linestyle=":", linewidth=1.4, label="7-Day Target (35%)", zorder=2)

    ax1.set_xticks(x)
    ax1.set_xticklabels(weeks, rotation=45, ha="right", fontsize=9, fontweight="bold")
    ax1.set_ylabel("Conversion Rate (%)", fontsize=10, fontweight="bold", color="#1F2937")
    ax1.set_title("National 2nd Recharge Milestone Trends by Cohort Week", fontsize=11, fontweight="bold", color="#006341", pad=12)
    ax1.set_ylim(0, 85)
    ax1.legend(loc="upper left", frameon=True, facecolor="#FFFFFF", edgecolor="#D0D7DE", fontsize=8)

    # =========================================================================
    # Panel 2: Regional 20-Day Conversion Leaderboard vs 70% Target
    # =========================================================================
    ax2.set_facecolor("#FAFAFA")
    ax2.grid(True, axis="x", linestyle="--", alpha=0.4, color="#D0D7DE", zorder=1)

    regions = reg_agg["DS_EXEC_REGION"].tolist()
    rates = (reg_agg["PCT_20D"] * 100).tolist()
    y = np.arange(len(regions))

    # Bar colors: highlight top region with Deep Green, others with Brand Green
    colors = [BRAND_DEEP_GREEN if r >= 45 else (BRAND_GREEN if r >= 35 else "#F59E0B") for r in rates]

    bars = ax2.barh(y, rates, color=colors, height=0.6, zorder=3)
    ax2.axvline(70, color=BRAND_RED, linestyle="--", linewidth=1.8, label="National Target (70%)", zorder=4)

    # Add data labels
    for bar in bars:
        width = bar.get_width()
        ax2.text(
            width + 1.0,
            bar.get_y() + bar.get_height() / 2,
            f"{width:.1f}%",
            va="center",
            ha="left",
            fontsize=9,
            fontweight="bold",
            color="#1F2937",
        )

    ax2.set_yticks(y)
    ax2.set_yticklabels(regions, fontsize=9, fontweight="bold", color="#1F2937")
    ax2.set_xlabel("20-Day Conversion Rate (%)", fontsize=10, fontweight="bold", color="#1F2937")
    ax2.set_title("Regional 2nd Recharge 20-Day Conversion Ranking", fontsize=11, fontweight="bold", color="#006341", pad=12)
    ax2.set_xlim(0, 80)
    ax2.legend(loc="lower right", frameon=True, facecolor="#FFFFFF", edgecolor="#D0D7DE", fontsize=8)

    plt.tight_layout(pad=2.5)
    plt.savefig(target_path, dpi=150, bbox_inches="tight")
    plt.close()

    print(f"[CHART] Successfully rendered executive visual dashboard: {target_path}")
    return target_path

