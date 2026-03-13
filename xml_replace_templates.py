"""
XML-only custom field replacement for Word templates.

This script isolates the XML replacement pass used in mass_update_templates.py.
It is useful for debugging templates where custom fields live in text boxes/shapes.
"""

import argparse
import csv
import shutil
import sys
from pathlib import Path
from typing import List, Tuple

from mass_update_templates import apply_xml_replacements, load_lookup_table

DEFAULT_EXCEL_PATH = Path(
    r"C:\Users\Tim\OneDrive - quillarrowlaw.com\Documents\ClioTemplates_CustomFields_MassUpdate\CustomField_LookupTable.xlsx"
)
DEFAULT_SHEET_NAME = "LookupTable"


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Run the XML-only custom field replacement pass on .docx files. "
            "Useful for text boxes/shapes not exposed by python-docx."
        )
    )
    parser.add_argument(
        "--excel",
        default=str(DEFAULT_EXCEL_PATH),
        help="Path to Excel lookup table (.xlsx).",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET_NAME,
        help="Worksheet name containing Old/New values.",
    )
    parser.add_argument(
        "--input-dir",
        required=True,
        help="Directory containing Word .docx templates.",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        help="Directory to write XML-updated templates.",
    )
    parser.add_argument(
        "--old-col",
        help="Old value column header or 1-based index (e.g., 'Old' or '1').",
    )
    parser.add_argument(
        "--new-col",
        help="New value column header or 1-based index (e.g., 'New' or '2').",
    )
    parser.add_argument(
        "--ignore-case",
        action="store_true",
        help="Case-insensitive matching.",
    )
    parser.add_argument(
        "--skip-headers-footers",
        action="store_true",
        help="Skip replacements in headers and footers.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview matches without writing files.",
    )
    parser.add_argument(
        "--report",
        help=(
            "CSV report path for per-document XML replacements. "
            "Defaults to xml_replacement_report.csv in the output directory."
        ),
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir).resolve()
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Input directory not found: {input_dir}", file=sys.stderr)
        return 1

    output_dir = Path(args.output_dir).resolve()
    if not args.dry_run:
        output_dir.mkdir(parents=True, exist_ok=True)

    report_path = (
        Path(args.report).resolve()
        if args.report
        else (input_dir if args.dry_run else output_dir) / "xml_replacement_report.csv"
    )

    # XML replacement uses literal matching (not regex) to avoid XML pattern errors.
    replacements = load_lookup_table(
        Path(args.excel).resolve(),
        args.sheet,
        args.old_col,
        args.new_col,
        use_regex=False,
        ignore_case=args.ignore_case,
    )

    docx_files = [
        path for path in input_dir.rglob("*.docx") if not path.name.startswith("~$")
    ]
    if not docx_files:
        print(f"No .docx files found under {input_dir}", file=sys.stderr)
        return 1

    print(f"Loaded {len(replacements)} replacement rows.")
    print(f"Processing {len(docx_files)} .docx files from {input_dir}")
    if args.dry_run:
        print("Dry run enabled; no files will be written.")

    total_files_changed = 0
    total_replacements = 0
    report_rows: List[Tuple[str, str, str, int]] = []

    for file_path in docx_files:
        rel_path = file_path.relative_to(input_dir)
        target_path = output_dir / rel_path

        if args.dry_run:
            counts = apply_xml_replacements(
                file_path,
                replacements,
                ignore_case=args.ignore_case,
                include_headers_footers=not args.skip_headers_footers,
                apply_changes=False,
            )
        else:
            target_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(file_path, target_path)
            counts = apply_xml_replacements(
                target_path,
                replacements,
                ignore_case=args.ignore_case,
                include_headers_footers=not args.skip_headers_footers,
                apply_changes=True,
            )

        count = sum(counts)
        total_replacements += count

        for idx, rep_count in enumerate(counts):
            if rep_count:
                replacement = replacements[idx]
                report_rows.append(
                    (
                        str(rel_path),
                        replacement.old_text,
                        replacement.new_text,
                        rep_count,
                    )
                )

        if count:
            total_files_changed += 1
            status = "updated"
        else:
            status = "no changes"
        print(f"{rel_path} -> {status} ({count} replacements)")

    print(
        f"Done. Files updated: {total_files_changed}/{len(docx_files)}. "
        f"Total replacements: {total_replacements}."
    )
    if not args.dry_run:
        print(f"Output directory: {output_dir}")

    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["document", "old_value", "new_value", "count"])
        writer.writerows(report_rows)
    print(f"Report written to: {report_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
