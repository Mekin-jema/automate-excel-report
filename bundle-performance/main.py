import argparse

from services.database import load_bundle_data
from services.report import create_excel, create_pdf
from services.email_service import send_outlook_email


def main():

    parser = argparse.ArgumentParser(
        description="Bundle Performance Reporting"
    )

    parser.add_argument("--start", required=True)
    parser.add_argument("--end", required=True)
    parser.add_argument(
        "--source",
        choices=["duckdb", "oracle", "auto"],
        default="auto",
        help="Database engine to use: 'duckdb' (sample data), 'oracle' (production), or 'auto' (fallback)."
    )

    parser.add_argument(
        "--send-email",
        action="store_true"
    )

    args = parser.parse_args()

    if (
        len(args.start) != 8
        or len(args.end) != 8
        or not args.start.isdigit()
        or not args.end.isdigit()
    ):
        parser.error("Dates must use YYYYMMDD format.")

    print(f"Loading data (source: {args.source})...")

    usage, purchases, actual_src = load_bundle_data(
        args.start,
        args.end,
        source=args.source
    )

    print(f"Queries completed using {actual_src}.")
    print(f"Loaded {len(purchases)} purchase rows, {len(usage)} usage rows.")

    excel = create_excel(
        usage,
        purchases,
        args.start,
        args.end
    )

    pdf = create_pdf(
        usage,
        purchases,
        args.start,
        args.end
    )

    print(f"Excel created: {excel}")
    print(f"PDF created: {pdf}")

    if args.send_email:

        send_outlook_email(
            subject=(
                f"Bundle Performance Report "
                f"{args.start} - {args.end}"
            ),
            body=(
                "Hello,\n\n"
                "Please find attached the Bundle Performance Report."
                "\n\nBest regards,\nMekin Jemal"
            ),
            attachments=[excel, pdf]
        )

        print("Email sent successfully.")


if __name__ == "__main__":
    main()
