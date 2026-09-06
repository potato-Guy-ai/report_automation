"""
Surgical date correction for finished reports.

This module is intentionally self-contained (depends only on python-docx)
so that it can be added to an existing project without touching the report
generation engine.

It is built as a small pipeline so additional correction fields
(venue, event title, faculty name, ...) can be plugged in later by adding
new handlers and registering them in FIELD_REGISTRY.

Date correction flow
--------------------
1. Detect the primary event date from the "Date & Time" field.
2. Find every date token in the document (body + tables) that represents
   the same calendar day as the primary event date.
3. Preview the proposed changes (location + old text -> new text).
4. User confirms the exact list.
5. Apply only the confirmed changes, preserving each occurrence's original
   date format and all run-level formatting (bold, size, font).

Only the matched date tokens are modified. The time part inside
"Date & Time" ("9 AM -- 10 AM") and unrelated dates on other days are never
touched.
"""

from datetime import date as _date
from pathlib import Path
import re

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from PIL import Image


# ============================================================
# DATE PARSING
# ============================================================

_MONTH_FULL = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

_MONTH_INDEX = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

# 1) 12-02-2025 / 12.02.2025 / 12/02/2025  (numeric, day-first assumed)
# 2) 12th February 2025 / 12 Feb 2025      (day + month name + year)
# 3) February 12, 2025 / Feb 12, 2025      (month name + day + year)
_DATE_PATTERN = re.compile(
    r"(\d{1,2}([-/.])\d{1,2}\2\d{2,4})"
    r"|(\d{1,2})(st|nd|rd|th)?\s+([A-Za-z]{3,})\s+(\d{2,4})"
    r"|([A-Za-z]{3,})\s+(\d{1,2})(st|nd|rd|th)?\s*,?\s*(\d{2,4})",
    re.IGNORECASE,
)


def _month_by_name(name):
    return _MONTH_INDEX.get(name[:3].lower())


def _ordinal_suffix(day):
    if 10 <= day % 100 <= 20:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


def _coerce_year(year_str):
    year = int(year_str)
    if year < 100:
        year += 2000
    return year


def _pad(value, width):
    text = str(value)
    if width and width > len(text):
        return text.zfill(width)
    return text


def parse_date_match(match):
    """Return (canonical_date, format_descriptor) or None."""
    if match.group(1):
        separator = match.group(2)
        day_str, month_str, year_str = match.group(1).split(separator)
        day = int(day_str)
        month = int(month_str)
        year = _coerce_year(year_str)
        fmt = {
            "kind": "numeric",
            "sep": separator,
            "day_width": len(day_str),
            "month_width": len(month_str),
            "year_width": len(year_str),
        }
    elif match.group(3):
        day = int(match.group(3))
        month = _month_by_name(match.group(5))
        year = _coerce_year(match.group(6))
        fmt = {
            "kind": "daymonth",
            "suffix_present": bool(match.group(4)),
            "month_width": len(match.group(5)),
            "year_width": len(match.group(6)),
        }
    elif match.group(7):
        month = _month_by_name(match.group(7))
        day = int(match.group(8))
        year = _coerce_year(match.group(10))
        fmt = {
            "kind": "monthday",
            "suffix_present": bool(match.group(9)),
            "comma_present": "," in match.group(0),
            "month_width": len(match.group(7)),
            "year_width": len(match.group(10)),
        }
    else:
        return None

    if month is None:
        return None

    try:
        canonical = (year, month, day)
        _date(year, month, day)
    except ValueError:
        return None

    return canonical, fmt


def parse_date_token(text):
    """
    Parse the first valid date in a text string.

    Returns a canonical (year, month, day) tuple, or None.
    """
    for match in _DATE_PATTERN.finditer(text):
        parsed = parse_date_match(match)
        if parsed:
            canonical, _ = parsed
            return canonical
    return None


# ============================================================
# NEW DATE RENDERING (keeps each occurrence's original format)
# ============================================================

def render_new_date(fmt, new_date):
    """Render (year, month, day) using the original token's format."""
    year, month, day = new_date

    if fmt["kind"] == "numeric":
        rendered = (
            _pad(day, fmt["day_width"])
            + fmt["sep"]
            + _pad(month, fmt["month_width"])
            + fmt["sep"]
            + _pad(year, fmt["year_width"])
        )
        return rendered

    month_name = (
        _MONTH_FULL[month - 1]
        if fmt["month_width"] > 3
        else _MONTH_FULL[month - 1][:3]
    )

    day_part = str(day)
    if fmt.get("suffix_present"):
        day_part += _ordinal_suffix(day)

    year_part = str(year) if fmt["year_width"] == 4 else _pad(year % 100, 2)

    if fmt["kind"] == "daymonth":
        return f"{day_part} {month_name} {year_part}"

    if fmt["kind"] == "monthday":
        comma = "," if fmt.get("comma_present") else ""
        return f"{month_name} {day_part}{comma} {year_part}"

    return str(year)


# ============================================================
# DOCUMENT SCANNING
# ============================================================

def _iter_document_paragraphs(document):
    """Yield (location_label, paragraph) for body + table paragraphs."""
    # Keep references to seen elements alive so their Python object ids
    # cannot be reused by another paragraph proxy (id() reuse would
    # otherwise silently drop table cell paragraphs).
    seen = []

    for index, paragraph in enumerate(document.paragraphs):
        if paragraph._p in seen:
            continue
        seen.append(paragraph._p)
        yield f"Body paragraph {index}", paragraph

    for table_index, table in enumerate(document.tables):
        for row_index, row in enumerate(table.rows):
            for cell_index, cell in enumerate(row.cells):
                for paragraph in cell.paragraphs:
                    if paragraph._p in seen:
                        continue
                    seen.append(paragraph._p)
                    yield (
                        f"Table {table_index}, cell ({row_index},{cell_index})",
                        paragraph,
                    )


def find_primary_event_date(document):
    """
    Locate the event date inside the "Date & Time" field.

    Returns None, or a dict with 'date', 'token', 'fmt',
    'location', 'paragraph' and 'context'.
    """
    for location, paragraph in _iter_document_paragraphs(document):
        text = "".join(run.text for run in paragraph.runs)
        if "date & time" not in text.lower():
            continue
        for match in _DATE_PATTERN.finditer(text):
            parsed = parse_date_match(match)
            if parsed:
                canonical, fmt = parsed
                return {
                    "date": canonical,
                    "token": match.group(0),
                    "fmt": fmt,
                    "location": location,
                    "paragraph": paragraph,
                    "context": text.strip(),
                }
    return None


def find_matching_occurrences(document, target_date):
    """
    Find every date token across the document whose calendar date
    equals target_date. Returns a list of occurrence dicts.
    """
    occurrences = []

    for location, paragraph in _iter_document_paragraphs(document):
        text = "".join(run.text for run in paragraph.runs)
        for match in _DATE_PATTERN.finditer(text):
            parsed = parse_date_match(match)
            if not parsed:
                continue
            canonical, fmt = parsed
            if canonical != target_date:
                continue

            start, end = match.span()
            original = match.group(0)
            context_text = text.strip()

            first = max(0, start - 25)
            last = min(len(text), end + 25)
            snippet = text[first:last].replace("\n", " ")

            occurrences.append({
                "location": location,
                "paragraph": paragraph,
                "start": start,
                "end": end,
                "original": original,
                "fmt": fmt,
                "context": context_text,
                "snippet": snippet,
            })

    return occurrences


# ============================================================
# FULL DETECTION PIPELINE
# ============================================================

def detect_corrections(document, new_date_str):
    """
    Detect primary event date + all matching occurrences.

    Returns a dict with:
      primary   : dict or None
      new_date  : canonical (year, month, day) tuple
      matches   : list of occurrence dicts with a computed 'new' text
    """
    new_date = parse_date_token(new_date_str)
    if new_date is None:
        raise ValueError(
            f"Could not parse the corrected date: {new_date_str!r}"
        )

    primary = find_primary_event_date(document)
    if primary is None:
        return {
            "primary": None,
            "new_date": new_date,
            "matches": [],
        }

    occurrences = find_matching_occurrences(document, primary["date"])

    for occurrence in occurrences:
        occurrence["new"] = render_new_date(
            occurrence["fmt"],
            new_date,
        )
        occurrence["changed"] = occurrence["new"] != occurrence["original"]

    return {
        "primary": primary,
        "new_date": new_date,
        "matches": occurrences,
    }


# ============================================================
# RUN-AWARE APPLICATION (preserves run-level formatting)
# ============================================================

def _apply_occurrence(paragraph, start, end, new_text):
    """
    Replace the text span [start, end) of the paragraph's concatenated
    run text with new_text, writing back into the existing runs so each
    run keeps its own formatting (bold, size, font, ...).
    """
    runs = [run for run in paragraph.runs]

    offsets = []
    position = 0
    for run in runs:
        offsets.append((position, position + len(run.text), run))
        position += len(run.text)

    if not offsets:
        return

    start = max(0, min(start, position))
    end = max(start, min(end, position))

    start_index = None
    end_index = None

    for index, (run_start, run_end, _) in enumerate(offsets):
        if start_index is None and start <= run_end and start >= run_start:
            start_index = index
        if end_index is None and end <= run_end and end >= run_start:
            end_index = index

    if start_index is None or end_index is None:
        return

    if start_index == end_index:
        run_start, _, run = offsets[start_index]
        local_start = start - run_start
        local_end = end - run_start
        run.text = (
            run.text[:local_start]
            + new_text
            + run.text[local_end:]
        )
        return

    # First overlapping run carries the replacement text.
    first_run_start, _, first_run = offsets[start_index]
    local_start = start - first_run_start
    first_run.text = first_run.text[:local_start] + new_text

    # Runs in between are emptied.
    for middle in range(start_index + 1, end_index):
        offsets[middle][2].text = ""

    # Last overlapping run keeps its tail.
    _, last_run_end, last_run = offsets[end_index]
    local_end = end - offsets[end_index][0]
    last_run.text = last_run.text[local_end:]


def apply_occurrences(document, occurrences):
    """
    Apply a list of occurrence dicts (as produced by detect_corrections)
    that the user confirmed. Only occurrences where 'changed' is True are
    modified.
    """
    applied = 0
    for occurrence in occurrences:
        if not occurrence.get("changed"):
            continue
        _apply_occurrence(
            occurrence["paragraph"],
            occurrence["start"],
            occurrence["end"],
            occurrence["new"],
        )
        applied += 1
    return applied


# ============================================================
# SAVING HELPERS
# ============================================================

def load_document(source):
    """Load a document from a path or a file-like object (e.g. BytesIO)."""
    return Document(source)


def default_output_path(input_path):
    """'report.docx' -> 'corrected_report.docx' in the same folder."""
    input_path = Path(input_path)
    return input_path.parent / f"corrected_{input_path.stem}.docx"


# ============================================================
# IMAGE SECTIONS (poster / photos / full poster)
# ============================================================

def _paragraph_has_image(paragraph):
    """True when the paragraph contains a picture (drawing/pict/object)."""
    element = paragraph._element
    return bool(
        element.findall(".//" + qn("w:drawing"))
        or element.findall(".//" + qn("w:pict"))
        or element.findall(".//" + qn("w:object"))
    )


def _image_extents(paragraph):
    """
    Return (width, height) in EMU of the first image in the paragraph,
    or None when the paragraph has no image.
    """
    extents = paragraph._element.findall(
        ".//{http://schemas.openxmlformats.org/drawingml/2006/main}ext"
    )
    if not extents:
        return None
    try:
        return int(extents[0].get("cx")), int(extents[0].get("cy"))
    except (TypeError, ValueError):
        return None


def _paragraph_text(paragraph):
    return "".join(run.text for run in paragraph.runs).strip()


def find_image_sections(document):
    """
    Locate the poster / photo / full-poster image paragraphs in a finished
    report that follows the sample layout:

        ... 'Poster:' heading
             [poster medium image]
             'Figure 1: Poster'
             'Photos' heading
             [photo 1 image]
             [photo 2 image]
             [signature image]
             'Principal'
             [full poster image]

    Returns a dict of section_key -> paragraph. Unused keys are omitted so
    optional replacement gracefully falls back to "leave as-is".
    """
    paragraphs = document.paragraphs
    texts = [_paragraph_text(p) for p in paragraphs]
    has_image = [_paragraph_has_image(p) for p in paragraphs]

    sections = {}

    # --- Poster medium ------------------------------------------------
    poster_medium = None

    poster_label = next(
        (i for i, text in enumerate(texts)
         if text.lower().startswith("poster")),
        None,
    )
    if poster_label is not None and poster_label + 1 < len(paragraphs):
        if has_image[poster_label + 1]:
            poster_medium = paragraphs[poster_label + 1]

    if poster_medium is None:
        figure_caption = next(
            (i for i, text in enumerate(texts)
             if "figure 1" in text.lower()),
            None,
        )
        if figure_caption is not None and figure_caption - 1 >= 0:
            if has_image[figure_caption - 1]:
                poster_medium = paragraphs[figure_caption - 1]

    if poster_medium is not None:
        sections["poster_medium"] = poster_medium

    # --- Signature image (never replaced) ----------------------------
    signature = None

    principal_labels = [
        i for i, text in enumerate(texts)
        if "principal" in text.lower()
    ]
    for index in principal_labels:
        if index - 1 >= 0 and has_image[index - 1]:
            signature = paragraphs[index - 1]
            break
        if index + 1 < len(paragraphs) and has_image[index + 1]:
            signature = paragraphs[index + 1]
            break

    # --- Full poster --------------------------------------------------
    poster_full = None

    for index, paragraph in enumerate(paragraphs):
        if not has_image[index]:
            continue
        if paragraph is poster_medium or paragraph is signature:
            continue
        if paragraph.paragraph_format.page_break_before:
            poster_full = paragraph
            break

    if poster_full is None:
        for index in range(len(paragraphs) - 1, -1, -1):
            if not has_image[index]:
                continue
            paragraph = paragraphs[index]
            if paragraph is poster_medium or paragraph is signature:
                continue
            poster_full = paragraph
            break

    if poster_full is not None:
        sections["poster_full"] = poster_full

    # --- Photos -------------------------------------------------------
    photos_label = next(
        (i for i, text in enumerate(texts)
         if text.lower() == "photos"),
        None,
    )

    photos_start = photos_label + 1 if photos_label is not None else 0

    photos = []
    for index in range(photos_start, len(paragraphs)):
        paragraph = paragraphs[index]
        if paragraph is signature or paragraph is poster_medium:
            continue
        if paragraph is poster_full:
            continue
        if "principal" in texts[index].lower():
            break
        if has_image[index]:
            photos.append(paragraph)

    for key, paragraph in zip(
        ("photo1", "photo2"),
        photos[:2],
    ):
        sections[key] = paragraph

    return sections


def replace_paragraph_image(paragraph, image_stream):
    """
    Replace the image(s) in one paragraph with a new image.

    The new image is fitted to the size of the image that was already
    there, so the surrounding layout is preserved. The paragraph's own
    alignment and formatting are left untouched.
    """
    old_extents = _image_extents(paragraph)

    for run in paragraph.runs:
        for tag in ("drawing", "pict", "object"):
            for child in run._element.findall(qn("w:" + tag)):
                run._element.remove(child)

    run = (
        paragraph.runs[0]
        if paragraph.runs
        else paragraph.add_run()
    )

    if old_extents is None:
        run.add_picture(
            image_stream,
            width=Inches(5.7),
            height=Inches(4.5),
        )
        return

    old_width_emu, old_height_emu = old_extents

    with Image.open(image_stream) as image:

        width_px, height_px = image.size
        aspect_ratio = width_px / height_px

    image_stream.seek(0)

    # Fit by width first, then by height, same as the report engine.
    width_inches = old_width_emu / 914400
    height_inches = width_inches / aspect_ratio

    if height_inches > old_height_emu / 914400:
        height_inches = old_height_emu / 914400
        width_inches = height_inches * aspect_ratio

    run.add_picture(
        image_stream,
        width=Inches(width_inches),
        height=Inches(height_inches),
    )


def apply_image_replacements(document, image_mapping):
    """
    Replace images in a finished report.

    image_mapping keys: 'poster' (replaces both the medium and full
    poster images), 'photo1', 'photo2'. Any key without a matching
    section in the document is skipped and the existing image is kept.

    Returns the number of images actually replaced.
    """
    sections = find_image_sections(document)

    targets = []

    if "poster" in image_mapping:
        for key in ("poster_medium", "poster_full"):
            if key in sections:
                targets.append((sections[key], image_mapping["poster"]))

    for key in ("photo1", "photo2"):
        if key in image_mapping and key in sections:
            targets.append((sections[key], image_mapping[key]))

    replaced = 0
    for paragraph, stream in targets:
        stream.seek(0)
        replace_paragraph_image(paragraph, stream)
        replaced += 1

    return replaced


# ============================================================
# FIELD REGISTRY (extensible for future fields)
# ============================================================

FIELD_REGISTRY = {
    "date": {
        "label": "Date",
        "detect": detect_corrections,
    },
    # Future fields will be registered here, e.g.:
    #   "venue": {"label": "Venue", "detect": detect_venue},
    #   "title": {"label": "Report title", "detect": detect_title},
    #   "faculty": {"label": "Faculty name", "detect": detect_faculty},
}