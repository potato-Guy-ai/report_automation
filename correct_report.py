"""
CLI for correcting finished reports.

Usage:
    python correct_report.py report.docx --new-date 23-02-2026
    python correct_report.py report.docx --new-date 23-02-2026 --yes
    python correct_report.py report.docx --new-date 23/02/2026 -o corrected.docx
"""

import argparse
from pathlib import Path

from engine.corrector import (
    apply_image_replacements,
    apply_occurrences,
    default_output_path,
    detect_corrections,
    find_image_sections,
    load_document,
)


# ============================================================
# ARGUMENTS
# ============================================================

def build_parser():

    parser = argparse.ArgumentParser(
        description=(
            "Correct the event date in a finished report. "
            "Only dates that match the primary 'Date & Time' date are "
            "changed, each in its own original format."
        )
    )

    parser.add_argument(
        "input",
        help="Path to the finished report (.docx).",
    )

    parser.add_argument(
        "--new-date",
        required=True,
        help=(
            "The corrected event date, e.g. 23-02-2026, "
            "23/02/2026 or 23rd February 2026."
        ),
    )

    parser.add_argument(
        "-o",
        "--output",
        help=(
            "Output path. Defaults to corrected_<name>.docx "
            "next to the input file."
        ),
    )

    parser.add_argument(
        "--poster",
        help=(
            "Optional replacement poster image. Replaces both the "
            "medium poster and the full poster."
        ),
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
        "--yes",
        action="store_true",
        help="Apply all detected corrections without prompting.",
    )

    return parser


# ============================================================
# PREVIEW
# ============================================================

def print_preview(result, image_sections=None, image_args=None):

    primary = result["primary"]

    print()
    print("=" * 64)
    print(" CORRECTION PREVIEW")
    print("=" * 64)

    if primary is None:

        print()
        print(" Could not locate a 'Date & Time' field or parse a date.")
        print(" Nothing was changed.")
        print()
        return False

    print()
    print(f" Primary event date : {primary['token']}")
    print(f" Location           : {primary['location']}")
    print()

    changed = [m for m in result["matches"] if m["changed"]]

    if not changed:

        print(" No dates match the primary event date.")

    for index, match in enumerate(changed, start=1):

        print(f" {index:2d}. {match['original']}  ->  {match['new']}")
        print(f"     {match['location']}")
        print(f"     ...{match['snippet']}...")
        print()

    if changed:
        print(f" {len(changed)} date(s) will be updated.")

    image_labels = {
        "poster": "Poster (medium + full)",
        "photo1": "Event photo 1",
        "photo2": "Event photo 2",
    }

    if image_sections and image_args:

        print()
        print(" Image replacements:")

        for arg_key in ("poster", "photo1", "photo2"):

            if arg_key not in image_args:
                continue

            section_key = {
                "poster": "poster_medium",
                "photo1": "photo1",
                "photo2": "photo2",
            }[arg_key]

            if section_key in image_sections:

                print(
                    f"   {image_labels[arg_key]:24s}"
                    f"-> will be replaced"
                )

            else:

                print(
                    f"   {image_labels[arg_key]:24s}"
                    f"-> no matching image found, left as-is"
                )

    print()

    return bool(changed) or bool(image_sections and image_args)


# ============================================================
# MAIN
# ============================================================

def main():

    args = build_parser().parse_args()

    input_path = Path(args.input)

    if not input_path.exists():
        print(f"Input file not found: {input_path}")
        raise SystemExit(1)

    document = load_document(input_path)

    try:
        date_result = detect_corrections(document, args.new_date)
    except ValueError as error:
        print(f"Error: {error}")
        raise SystemExit(1)

    image_args = {
        key: Path(value)
        for key, value in (
            ("poster", args.poster),
            ("photo1", args.photo1),
            ("photo2", args.photo2),
        )
        if value
    }

    image_sections = (
        find_image_sections(document)
        if image_args
        else {}
    )

    has_changes = print_preview(
        date_result,
        image_sections,
        image_args,
    )

    if not has_changes:
        return

    if not args.yes:
        answer = input("Apply these corrections? [y/N] ").strip().lower()
        if answer not in {"y", "yes"}:
            print("Cancelled. No changes written.")
            return

    confirmed = [
        match
        for match in date_result["matches"]
        if match["changed"]
    ]

    total = apply_occurrences(document, confirmed)

    if image_args:

        image_mapping = {}

        for arg_key, path in image_args.items():
            if path.exists():
                image_mapping[arg_key] = open(path, "rb")

        images_replaced = apply_image_replacements(
            document,
            image_mapping,
        )

        for stream in image_mapping.values():
            stream.close()

    else:

        images_replaced = 0

    output_path = (
        Path(args.output)
        if args.output
        else default_output_path(input_path)
    )

    document.save(output_path)

    print()
    print(f" Applied {total} date correction(s).")
    if image_args:
        print(f" Replaced {images_replaced} image(s).")
    print(f" Output : {output_path}")
    print(" Original file was left untouched.")
    print()


if __name__ == "__main__":
    main()