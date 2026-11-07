"""Validate report values and render dashboard email assets."""

from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patheffects as path_effects
import matplotlib.ticker as mticker
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

from services.performance.paths import PREVIEW_DIR
from services.performance_email import DailySummary, KPI_LABELS, format_report_period


def _validate_report_data(
    purchases: pd.DataFrame,
    usage: pd.DataFrame,
    required_metrics: set[str],
    label: str,
) -> None:
    if purchases.empty:
        raise ValueError(
            f"{label} purchase query returned no rows; email was not sent."
        )
    if usage.empty:
        raise ValueError(f"{label} usage query returned no rows; email was not sent.")
    if not {"ID_DATE", "METRIC", "VALUE"}.issubset(purchases.columns):
        raise ValueError(f"{label} purchase query returned an invalid shape.")
    if not {"ID_DATE", "METRIC", "VALUE"}.issubset(usage.columns):
        raise ValueError(f"{label} usage query returned an invalid shape.")

    purchase_values = pd.to_numeric(purchases["VALUE"], errors="coerce")
    usage_values = pd.to_numeric(usage["VALUE"], errors="coerce")
    if purchase_values.isna().all() or usage_values.isna().all():
        raise ValueError(f"{label} purchase or usage values are not numeric.")
    actual_metrics = set(purchases["METRIC"].dropna().astype(str))
    missing = required_metrics - actual_metrics
    if missing:
        raise ValueError(
            f"{label} data is missing required metrics: {', '.join(sorted(missing))}; "
            "email was not sent."
        )


def _kpi_metrics(
    data_purchases: pd.DataFrame,
    data_usage: pd.DataFrame,
    voice_purchases: pd.DataFrame,
    voice_usage: pd.DataFrame,
) -> dict[str, tuple[str, str]]:
    def totals(
        purchases: pd.DataFrame,
        usage: pd.DataFrame,
        *,
        usage_divisor: float = 1,
    ) -> dict[str, float]:
        values = pd.to_numeric(purchases["VALUE"], errors="coerce")
        metric_values = (
            purchases.assign(_VALUE=values).groupby("METRIC")["_VALUE"].sum()
        )
        usage_values = pd.to_numeric(
            usage.loc[usage["METRIC"] == "USAGE", "VALUE"], errors="coerce"
        )
        if usage_values.empty or usage_values.isna().all():
            raise ValueError("Usage metric has no numeric values; email was not sent.")
        return {
            "Revenue": float(metric_values["C.REVENUE"]),
            "Purchase Quantity": float(metric_values["B.QUANTITY"]),
            "Usage": float(usage_values.sum()) / usage_divisor,
        }

    data = totals(data_purchases, data_usage, usage_divisor=1024)
    voice = totals(voice_purchases, voice_usage)

    def format_value(label: str, value: float, *, voice_value: bool) -> str:
        if label == "Revenue":
            return f"{value:,.0f} ETB"
        if label in ("Bundle Subscribers", "Purchase Quantity"):
            return f"{value:,.0f}"
        if label == "Usage":
            return f"{value:,.1f} GB" if not voice_value else f"{value:,.0f} minutes"
        raise ValueError(f"Unknown KPI label: {label}")

    return {
        label: (
            format_value(label, data[label], voice_value=False),
            format_value(label, voice[label], voice_value=True),
        )
        for label in KPI_LABELS
    }


def _daily_performance_summary(
    reporting_date: str,
    data_purchases: pd.DataFrame,
    data_usage: pd.DataFrame,
    voice_purchases: pd.DataFrame,
    voice_usage: pd.DataFrame,
    data_covered_dates: set[str],
    voice_covered_dates: set[str],
) -> DailySummary:
    report_day = datetime.strptime(reporting_date, "%Y%m%d").date()
    previous_day = report_day - timedelta(days=1)

    def totals(
        purchases: pd.DataFrame,
        usage: pd.DataFrame,
        day: str,
        covered_dates: set[str],
        *,
        voice: bool,
    ) -> dict[str, float | None]:
        if day not in covered_dates:
            return {
                "Revenue": None,
                "Bundle Subscribers": None,
                "Purchase Quantity": None,
                "Usage": None,
            }
        day_number = int(day)
        purchase_day = purchases.loc[
            pd.to_numeric(purchases["ID_DATE"], errors="raise") == day_number
        ]
        usage_day = usage.loc[
            pd.to_numeric(usage["ID_DATE"], errors="raise") == day_number
        ]

        def metric_total(frame: pd.DataFrame, metric: str) -> float | None:
            values = frame.loc[frame["METRIC"] == metric, "VALUE"]
            if values.empty:
                return None
            numeric_values = pd.to_numeric(values, errors="coerce")
            if numeric_values.isna().any():
                raise ValueError(
                    f"Daily {metric} values for {day} contain non-numeric data."
                )
            return float(numeric_values.sum())

        usage_value = metric_total(usage_day, "USAGE")
        return {
            "Revenue": metric_total(purchase_day, "C.REVENUE"),
            "Bundle Subscribers": metric_total(purchase_day, "A.SUBS"),
            "Purchase Quantity": metric_total(purchase_day, "B.QUANTITY"),
            "Usage": (
                usage_value if voice or usage_value is None else usage_value / 1024
            ),
        }

    current_data = totals(
        data_purchases,
        data_usage,
        report_day.strftime("%Y%m%d"),
        data_covered_dates,
        voice=False,
    )
    previous_data = totals(
        data_purchases,
        data_usage,
        previous_day.strftime("%Y%m%d"),
        data_covered_dates,
        voice=False,
    )
    current_voice = totals(
        voice_purchases,
        voice_usage,
        report_day.strftime("%Y%m%d"),
        voice_covered_dates,
        voice=True,
    )
    previous_voice = totals(
        voice_purchases,
        voice_usage,
        previous_day.strftime("%Y%m%d"),
        voice_covered_dates,
        voice=True,
    )

    def paired_values(
        current: dict[str, float | None], previous: dict[str, float | None]
    ) -> dict[str, tuple[float | None, float | None]]:
        return {
            label: (current[label], previous[label])
            for label in ("Revenue", "Bundle Subscribers", "Purchase Quantity", "Usage")
        }

    return {
        "reporting_date": format_report_period(
            report_day.strftime("%Y%m%d"), report_day.strftime("%Y%m%d")
        ),
        "comparison_date": format_report_period(
            previous_day.strftime("%Y%m%d"), previous_day.strftime("%Y%m%d")
        ),
        "data": paired_values(current_data, previous_data),
        "voice": paired_values(current_voice, previous_voice),
    }


def _render_dashboard(
    image_path: Path,
    service: str,
    reporting_period: str,
    start_date: str,
    end_date: str,
    purchases: pd.DataFrame,
    usage: pd.DataFrame,
) -> None:
    if service not in {"Data", "Voice"}:
        raise ValueError(f"Unknown dashboard service: {service}")
    image_path.parent.mkdir(parents=True, exist_ok=True)
    images: list[Image.Image] = []

    required_columns = {"ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"}
    if not required_columns.issubset(
        purchases.columns
    ) or not required_columns.issubset(usage.columns):
        raise ValueError(f"{service} dashboard data has an invalid shape.")

    dates = pd.date_range(
        pd.to_datetime(start_date, format="%Y%m%d"),
        pd.to_datetime(end_date, format="%Y%m%d"),
        freq="D",
    )
    purchase_data = purchases.copy()
    usage_data = usage.copy()
    for frame in (purchase_data, usage_data):
        frame["_DATE"] = pd.to_datetime(
            pd.to_numeric(frame["ID_DATE"], errors="raise").astype("int64").astype(str),
            format="%Y%m%d",
        )
        frame["_VALUE"] = pd.to_numeric(frame["VALUE"], errors="coerce").fillna(0)

    purchase_data = purchase_data[
        purchase_data["METRIC"].isin(("A.SUBS", "B.QUANTITY", "C.REVENUE"))
    ].copy()
    usage_data = usage_data[usage_data["METRIC"] == "USAGE"].copy()

    if service == "Voice":
        from services.voice.report import VOICE_BUNDLES_BY_NAME

        category_by_bundle = {
            name: bundle["category"] for name, bundle in VOICE_BUNDLES_BY_NAME.items()
        }
        display_by_bundle = {
            name: bundle["display_name"]
            for name, bundle in VOICE_BUNDLES_BY_NAME.items()
        }
        purchase_data["BUNDLE_CATEGORY"] = (
            purchase_data["BUNDLE_NAME"].map(category_by_bundle).fillna("Other")
        )
    else:
        from services.data.report import display_bundle_name

        bundle_names = set(purchase_data["BUNDLE_NAME"]) | set(
            usage_data["BUNDLE_NAME"]
        )
        display_by_bundle = {
            str(name): display_bundle_name(str(name)) for name in bundle_names
        }
        if "BUNDLE_CATEGORY" not in purchase_data:
            raise ValueError(f"{service} dashboard purchase data has no categories.")

    revenue_by_date = (
        purchase_data.loc[purchase_data["METRIC"] == "C.REVENUE"]
        .groupby("_DATE")["_VALUE"]
        .sum()
        .reindex(dates, fill_value=0)
    )
    quantity_by_date = (
        purchase_data.loc[purchase_data["METRIC"] == "B.QUANTITY"]
        .groupby("_DATE")["_VALUE"]
        .sum()
        .reindex(dates, fill_value=0)
    )
    usage_by_date = (
        usage_data.groupby("_DATE")["_VALUE"].sum().reindex(dates, fill_value=0)
    )
    revenue_by_category = (
        purchase_data.loc[purchase_data["METRIC"] == "C.REVENUE"]
        .groupby("BUNDLE_CATEGORY", dropna=False)["_VALUE"]
        .sum()
        .sort_values(ascending=False)
    )
    revenue_by_category.index = revenue_by_category.index.fillna("Other")
    revenue_by_bundle = (
        purchase_data.loc[purchase_data["METRIC"] == "C.REVENUE"]
        .groupby("BUNDLE_NAME")["_VALUE"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
    )
    usage_by_bundle = (
        usage_data.groupby("BUNDLE_NAME")["_VALUE"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
    )

    kpi_values = (
        revenue_by_date.sum(),
        purchase_data.loc[purchase_data["METRIC"] == "A.SUBS", "_VALUE"].sum(),
        quantity_by_date.sum(),
        usage_by_date.sum(),
    )
    if not any(value != 0 for value in kpi_values):
        raise RuntimeError(
            f"{service} dashboard source metrics are empty or zero; "
            "workbook will not be emailed."
        )

    chart_specs = (
        (
            "Revenue Trend | ETB",
            "line",
            dates,
            revenue_by_date,
            "#00A859",
            "Revenue (ETB)",
        ),
        (
            (
                "Bundle Purchase Quantity"
                if service == "Data"
                else "Voice Purchase Quantity"
            ),
            "column",
            dates,
            quantity_by_date,
            "#F6C445",
            "Quantity",
        ),
        (
            "Traffic Trend | GB" if service == "Data" else "Voice Minutes Trend",
            "line",
            dates,
            usage_by_date / 1024 if service == "Data" else usage_by_date,
            "#00A859",
            "Traffic (GB)" if service == "Data" else "Minutes",
        ),
        (
            "Revenue Mix by Category",
            "doughnut",
            revenue_by_category.index,
            revenue_by_category,
            "#00A859",
            "",
        ),
        (
            "Top Bundles by Revenue",
            "bar",
            revenue_by_bundle.index,
            revenue_by_bundle,
            "#006341",
            "Revenue (ETB)",
        ),
        (
            (
                "Top Bundles by Traffic"
                if service == "Data"
                else "Top Bundles by Voice Minutes"
            ),
            "bar",
            usage_by_bundle.index,
            usage_by_bundle / 1024 if service == "Data" else usage_by_bundle,
            "#F6C445",
            "Traffic (GB)" if service == "Data" else "Minutes",
        ),
    )
    try:
        for title, chart_type, labels, values, color, axis_label in chart_specs:
            figure, axis = plt.subplots(figsize=(7.5, 3.5), dpi=120)
            try:
                figure.patch.set_facecolor("white")
                axis.set_facecolor("white")
                if chart_type == "doughnut":
                    nonzero = values[values > 0]
                    if nonzero.empty:
                        axis.text(0.5, 0.5, "No revenue data", ha="center", va="center")
                        axis.set_axis_off()
                    else:
                        axis.pie(
                            nonzero.values,
                            autopct="%1.0f%%",
                            startangle=90,
                            wedgeprops={"width": 0.42, "edgecolor": "white"},
                            textprops={
                                "color": "white",
                                "weight": "bold",
                                "path_effects": [
                                    path_effects.withStroke(
                                        linewidth=2,
                                        foreground="#263238",
                                    )
                                ],
                            },
                        )
                        axis.set_aspect("equal")
                        legend_handles = [
                            plt.Line2D(
                                [0],
                                [0],
                                marker="o",
                                linestyle="",
                                markerfacecolor=wedge.get_facecolor(),
                                markeredgecolor="none",
                                markersize=8,
                                label=str(category),
                            )
                            for category, wedge in zip(nonzero.index, axis.patches)
                        ]
                        axis.legend(
                            handles=legend_handles,
                            loc="center left",
                            bbox_to_anchor=(1.02, 0.5),
                            frameon=False,
                            fontsize=8,
                            title="Category",
                            title_fontsize=9,
                        )
                elif chart_type in ("line", "column"):
                    axis_labels = pd.to_datetime(labels).strftime("%d %b")
                    if chart_type == "line":
                        axis.plot(
                            axis_labels,
                            values.values,
                            color=color,
                            marker="o",
                            linewidth=2,
                        )
                    else:
                        axis.bar(axis_labels, values.values, color=color, width=0.65)
                    tick_step = max(1, (len(axis_labels) + 11) // 12)
                    tick_positions = list(range(0, len(axis_labels), tick_step))
                    axis.set_xticks(tick_positions)
                    axis.set_xticklabels(
                        [axis_labels[position] for position in tick_positions]
                    )
                    axis.tick_params(axis="x", labelrotation=35, labelsize=8)
                    axis.set_ylabel(axis_label)
                    axis.yaxis.set_major_formatter(
                        mticker.StrMethodFormatter("{x:,.0f}")
                    )
                    axis.grid(axis="y", linestyle="--", alpha=0.2)
                    axis.set_axisbelow(True)
                else:
                    display_labels = [
                        display_by_bundle.get(str(name), str(name).split(".", 1)[-1])
                        for name in labels
                    ]
                    axis.barh(display_labels[::-1], values.values[::-1], color=color)
                    axis.set_xlabel(axis_label)
                    axis.xaxis.set_major_formatter(
                        mticker.StrMethodFormatter("{x:,.0f}")
                    )
                    axis.tick_params(axis="x", labelrotation=45, labelsize=8)
                    for tick in axis.get_xticklabels():
                        tick.set_horizontalalignment("right")
                    axis.grid(axis="x", linestyle="--", alpha=0.2)
                    axis.set_axisbelow(True)
                    axis.tick_params(axis="y", labelsize=8)

                axis.set_title(
                    title, fontsize=12, fontweight="bold", color="#006341", pad=12
                )
                for spine in ("top", "right"):
                    axis.spines[spine].set_visible(False)
                if chart_type == "doughnut":
                    figure.tight_layout(rect=(0, 0, 0.68, 1))
                else:
                    figure.tight_layout()
                with BytesIO() as stream:
                    figure.savefig(stream, format="png", dpi=120, facecolor="white")
                    stream.seek(0)
                    with Image.open(stream) as source:
                        images.append(source.convert("RGB").copy())
            finally:
                plt.close(figure)

        columns = 2
        rows = (len(images) + columns - 1) // columns
        cell_width = max(image.width for image in images)
        cell_height = max(image.height for image in images)
        gap = 14
        header_height = 54
        canvas = Image.new(
            "RGB",
            (
                columns * cell_width + (columns + 1) * gap,
                rows * cell_height + (rows + 1) * gap + header_height,
            ),
            "#F4F7F5",
        )
        draw = ImageDraw.Draw(canvas)
        try:
            title_font = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 23)
        except OSError:
            title_font = ImageFont.load_default()
        draw.text(
            (gap, 12),
            f"{service} Performance Dashboard | {reporting_period}",
            fill="#006341",
            font=title_font,
        )
        for index, image in enumerate(images):
            x = gap + (index % columns) * (cell_width + gap)
            y = header_height + gap + (index // columns) * (cell_height + gap)
            canvas.paste(image, (x, y))
        canvas.save(image_path, format="PNG", optimize=True)
        canvas.close()
    finally:
        for image in images:
            image.close()


def _write_previews(
    html_body: str,
    start_date: str,
    end_date: str,
) -> tuple[Path, Path]:
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    suffix = f"{start_date}_{end_date}"
    desktop = PREVIEW_DIR / f"Performance_{suffix}_Desktop.html"
    mobile = PREVIEW_DIR / f"Performance_{suffix}_Mobile.html"
    desktop.write_text(html_body, encoding="utf-8")
    mobile.write_text(html_body, encoding="utf-8")
    return desktop, mobile
