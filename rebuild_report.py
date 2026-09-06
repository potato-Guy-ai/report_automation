"""
Regenerate a finished report as a brand-new, correctly aligned document.

The tool reads every piece of information out of an existing report
(title, date & time, venue, description, poster and event photos), lets
you override whatever is wrong, and regenerates a fresh document from
the template.

Usage:
    python rebuild_report.py report.docx --new-date 23-02-2026 --new-venue "CS BLOCK B"
    python rebuild_report.py report.docx --new-date 23/02/2026 --poster poster.jpg
    python rebuild_report.py report.docx -o rebuilt.docx
"""

import argparse
import io
from pathlib import Path

from docx import Document

from engine.template_loader import load_template
from engine.content_inserter import replace_placeholders
from engine.extractor import extract_report_data, build_generation_data
from engine.corrector import load_document


def build_parser():

    parser = argparse.ArgumentParser(
        description=(
            "Extract the contents of an existing finished report and "
            "regenerate it from the template as a clean, properly "
            "aligned document. Override any extracted value with the "
            "options below."
        )
    )

    parser.add_argument(
        "input",
        help="Path to the finished report (.docx).",
    )

    parser.add_argument(
        "-o",
        "--output",
        help=(
            "Output path. Defaults to rebuilt_<name>.docx "
            "next to the input file."
        ),
    )

    parser.add_argument(
        "--template",
        default=(
            Path(__file__).parent / "templates" / "final_sample_report.docx"
        ),
        help="Template to regenerate from (defaults to final_sample_report.docx).",
    )

    parser.add_argument(
        "--new-date",
        help=(
            "The corrected event date, e.g. 23-02-2026. The extracted "
            "'Date & Time' date token is replaced in its own format."
        ),
    )

    parser.add_argument(
        "--date-time",
        help="Replace the whole 'Date & Time' value, e.g. '23-02-2026 & 01.30 am to 5 pm'.",
    )

    parser.add_argument(
        "--new-venue",
        help="The corrected venue value, e.g. 'CS BLOCK B'.",
    )

    parser.add_argument(
        "--title",
        help="The corrected event title.",
    )

    parser.add_argument(
        "--description",
        help="Path to a text file containing the corrected report description "
        "(blank lines split into paragraphs).",
    )

    parser.add_argument(
        "--poster",
        help="Optional replacement poster image (used for medium + full poster).",
    )

    parser.add_argument(
        "--photo1",
        help="Optional replacement for the first event photo.",
    )

    parser.add_argument(
        "--photo2",
        help="Optional replacement for the second event photo.",
    )

    parser.add_argument(
        "--preview",
        action="store_true",
        help="Show the extracted values and quit without writing anything.",
    )

    return parser


def main():

    args = build_parser().parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Input file not found: {input_path}")
        raise SystemExit(1)

    template_path = Path(args.template)
    if not template_path.exists():
        print(f"Template not found: {template_path}")
        raise SystemExit(1)

    document = load_document(input_path)

    try:
        extracted = extract_report_data(
            document,
            new_date=args.new_date,
            new_venue=args.new_venue,
        )
    except ValueError as error:
        print(f"Error: {error}")
        raise SystemExit(1)

    print()
    print("=" * 64)
    print(" EXTRACTED REPORT DATA")
    print("=" * 64)
    print()
    print(f" Date & Time : {extracted['date_time']!r}")
    print(f" Venue       : {extracted['venue']!r}")
    print(f" Title       : {extracted['title']!r}")
    preview_paragraphs = (extracted["description"] or "").split("\n\n")
    print(f" Description : {len(preview_paragraphs)} paragraph(s)")
    for paragraph in preview_paragraphs:
        print(f"   - {paragraph[:80]}")
    print(f" Images      : poster={bool(extracted['poster_image'])} "
          f"photo1={bool(extracted['photo1_image'])} "
          f"photo2={bool(extracted['photo2_image'])}")
    print()

    if args.preview:
        print(" Preview only. Nothing was written.")
        return

    if (
        not extracted["date_time"]
        and not args.date_time
    ):
        print("Could not extract a 'Date & Time' value.")
        print("Nothing was written.")
        raise SystemExit(1)

    if args.date_time:
        extracted["date_time"] = args.date_time
    if args.title:
        extracted["title"] = args.title
    if args.description:
        extracted["description"] = Path(args.description).read_text(
            encoding="utf-8"
        )

    image_overrides = {}

    for key, value in (
        ("poster", args.poster),
        ("photo1", args.photo1),
        ("photo2", args.photo2),
    ):
        if value:
            image_overrides[key] = Path(value)

    data = build_generation_data(extracted, image_overrides)

    generated = load_template(template_path)
    generated = replace_placeholders(generated, data)

    output_path = (
        Path(args.output)
        if args.output
        else input_path.parent / f"rebuilt_{input_path.stem}.docx"
    )

    generated.save(output_path)

    print(f" Regenerated report written to : {output_path}")
    print(" Original file was left untouched.")
    print()


if __name__ == "__main__":
    main()