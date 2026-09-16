"""
Student Certificate Generator CLI Entry Point.

Automates generating individual student certificates from an Excel file
and a DOCX certificate template, exporting pixel-perfect PDFs named <REG_NO>.pdf.
"""

import sys
import argparse
from pathlib import Path

# Ensure proper unicode handling on Windows console
if sys.platform == "win32":
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    if hasattr(sys.stderr, "reconfigure"):
        try:
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

from app.certificate_generator import CertificateGenerator


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Automated Student Certificate Mail-Merge and PDF Generator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Automatic discovery from current working directory:
  python main.py

  # With optional explicit files:
  python main.py --excel Student_Details.xlsx --template certificate_template.docx

  # Non-interactive / CI automation:
  python main.py --yes --overwrite
"""
    )

    parser.add_argument(
        "--root-dir", "-d",
        type=str,
        default=".",
        help="Root/input directory to inspect (defaults to current working directory)."
    )

    parser.add_argument(
        "--excel", "-e",
        type=str,
        default=None,
        help="Path to student Excel file (optional; automatically detected if omitted)."
    )

    parser.add_argument(
        "--template", "-t",
        type=str,
        default=None,
        help="Path to DOCX certificate template (optional; automatically detected if omitted)."
    )

    parser.add_argument(
        "--output", "-o",
        type=str,
        default=None,
        help="Output certificates directory (defaults to <root-dir>/output/certificates/)."
    )

    parser.add_argument(
        "--yes", "-y",
        action="store_true",
        help="Automatically confirm and generate without interactive prompts."
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing output PDFs without prompting."
    )

    parser.add_argument(
        "--engine",
        choices=["auto", "libreoffice", "word"],
        default="auto",
        help="PDF conversion engine: 'libreoffice', 'word', or 'auto' (default: auto, prefers LibreOffice)."
    )

    parser.add_argument(
        "--font-size",
        type=float,
        default=20.0,
        help="Font size (in pt) for student name and registration number (default: 20.0 pt)."
    )

    parser.add_argument(
        "--no-bold",
        action="store_true",
        help="Do not bold the student name and registration number (bold is applied by default)."
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Perform discovery and validation checks only, without generating certificates."
    )

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    engine_param = None if args.engine == "auto" else args.engine

    generator = CertificateGenerator(
        root_dir=args.root_dir,
        excel_path=args.excel,
        template_path=args.template,
        output_dir=args.output,
        auto_confirm=args.yes,
        overwrite=args.overwrite,
        force_engine=engine_param,
        font_size=args.font_size,
        bold=not args.no_bold
    )

    if args.dry_run:
        print("Running in Dry-Run mode (Validation Only)...")
        success = generator.discover_and_validate()
        return 0 if success else 1

    success = generator.run()
    return 0 if success else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        sys.exit(130)
    except Exception as exc:
        print(f"\nFATAL ERROR: {exc}")
        sys.exit(1)
