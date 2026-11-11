#!/usr/bin/env python3
"""
CLI entry point for KG Recharge Tracker Email and Excel Automation.
===================================================================

Usage examples:
---------------
1. Generate Excel report and standalone HTML email previews using DuckDB:
   uv run python send_kg_email.py --source duckdb --preview-only

2. Connect to Oracle EDW (with automatic fallback to DuckDB):
   uv run python send_kg_email.py --source oracle --preview-only

3. Dispatch email to configured SMTP recipient:
   uv run python send_kg_email.py --source duckdb --send
"""

import argparse
from pathlib import Path
import sys

# Ensure root and src are on sys.path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(BASE_DIR / "src"))

from config.settings import DEFAULT_SOURCE, REPORT_EMAIL_TO
from kg.services.runner import run_kg_pipeline


def parse_args():
    parser = argparse.ArgumentParser(
        description="Daily Gross Add (KG) First and Second Recharge Cohort Reporter"
    )
    parser.add_argument(
        "--source",
        choices=["duckdb", "oracle"],
        default=DEFAULT_SOURCE,
        help="Data source: 'duckdb' for local offline data or 'oracle' for EDW production. Defaults to duckdb.",
    )
    parser.add_argument(
        "--send",
        action="store_true",
        help="Send executive email via SMTP with Excel attachment. Defaults to False (preview only).",
    )
    parser.add_argument(
        "--preview-only",
        action="store_true",
        default=False,
        help="Generate Excel report and HTML previews without sending email.",
    )
    parser.add_argument(
        "--recipient",
        type=str,
        default=None,
        help=f"Target recipient email address override (default: {REPORT_EMAIL_TO}).",
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help="Custom output filename for the generated Excel workbook.",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    preview_mode = args.preview_only or (not args.send)

    result = run_kg_pipeline(
        source=args.source,
        send=args.send,
        preview_only=preview_mode,
        recipient=args.recipient,
        output_filename=args.output,
    )
    return 0 if result["status"] == "success" else 1


if __name__ == "__main__":
    sys.exit(main())

