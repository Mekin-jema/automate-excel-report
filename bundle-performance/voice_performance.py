"""Generate the Oracle-backed voice performance Excel report."""

import argparse
from datetime import datetime

from services.voice.database import load_voice_data
from services.voice.report import create_voice_excel


def main():
    parser = argparse.ArgumentParser(description="Voice performance Excel report")
    parser.add_argument("--start", required=True, help="Start date in YYYYMMDD format")
    parser.add_argument("--end", required=True, help="End date in YYYYMMDD format")
    parser.add_argument(
        "--source",
        choices=("oracle", "duckdb"),
        default="oracle",
        help="Data source: production Oracle or local synthetic DuckDB sample data",
    )
    args = parser.parse_args()

    if any(
        len(value) != 8 or not value.isdigit()
        for value in (args.start, args.end)
    ):
        parser.error("Dates must use YYYYMMDD format.")
    try:
        datetime.strptime(args.start, "%Y%m%d")
        datetime.strptime(args.end, "%Y%m%d")
    except ValueError:
        parser.error("Dates must be valid calendar dates in YYYYMMDD format.")
    if args.start > args.end:
        parser.error("--start must be earlier than or equal to --end.")

    purchases, usage, five_birr_usage = load_voice_data(
        args.start, args.end, source=args.source
    )
    report_path = create_voice_excel(
        purchases, usage, five_birr_usage, args.start, args.end
    )
    print(f"Voice performance report created: {report_path} (source: {args.source})")


if __name__ == "__main__":
    main()
