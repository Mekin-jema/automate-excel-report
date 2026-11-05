"""Validate report values and render dashboard email assets."""

from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

from services.performance.paths import PREVIEW_DIR
from services.performance_email import KPI_LABELS


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
            "Bundle Subscribers": float(metric_values["A.SUBS"]),
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


def _dashboard_kpi_columns(service: str) -> tuple[int, ...]:
    return (2, 5, 8, 11) if service == "Data" else (2, 5, 8, 11)


def _render_dashboard(
    workbook_path: Path,
    image_path: Path,
    service: str,
    reporting_period: str,
) -> None:
    import win32com.client

    image_path.parent.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    images: list[Image.Image] = []
    try:
        workbook = excel.Workbooks.Open(
            str(workbook_path.resolve()),
            UpdateLinks=0,
            ReadOnly=False,
        )
        excel.CalculateFullRebuild()
        dashboard = workbook.Worksheets("Dashboard")
        if dashboard.ChartObjects().Count != 6:
            raise RuntimeError(f"{service} report should contain six dashboard charts.")

        kpi_values = [
            dashboard.Cells(7, column).Value
            for column in _dashboard_kpi_columns(service)
        ]
        if not any(
            isinstance(value, (int, float)) and value != 0 for value in kpi_values
        ):
            raise RuntimeError(
                f"{service} dashboard formulas did not calculate. "
                "Workbook will not be emailed."
            )
        workbook.Save()

        with TemporaryDirectory(prefix=f"{service.lower()}_charts_") as temp_dir:
            temp_path = Path(temp_dir)
            chart_objects = dashboard.ChartObjects()
            for index in range(1, chart_objects.Count + 1):
                chart = chart_objects.Item(index).Chart
                exported = temp_path / f"chart_{index}.png"
                if not chart.Export(str(exported), "PNG") or not exported.is_file():
                    raise RuntimeError(
                        f"Could not export {service} dashboard chart {index}."
                    )
                with Image.open(exported) as source:
                    images.append(source.convert("RGB").copy())

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
    finally:
        for image in images:
            image.close()
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()


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
