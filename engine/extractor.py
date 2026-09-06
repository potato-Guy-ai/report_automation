"""
Extract the contents of a finished report so it can be regenerated
from the template as a brand-new, correctly aligned document.

The extraction is intentionally "best effort": it looks for the same
anchors the generator uses (Event Title, Date & Time line, Venue line,
Poster / Photos sections) and pulls the embedded images out as bytes.
The caller can override any extracted value before regeneration.

Depends only on python-docx + the correction helpers in engine.corrector.
"""

import re
import io

from docx.oxml.ns import qn

from engine.corrector import (
    _iter_document_paragraphs,
    _paragraph_has_image,
    find_image_sections,
    find_primary_venue,
    parse_date_match,
    parse_date_token,
    render_new_date,
    _DATE_PATTERN,
)

# Headings that end the description / separate report sections.
_SECTION_BOUNDARIES = (
    "poster",
    "figure 1",
    "enclosure",
    "photos",
    "principal",
)

_DATE_TIME_LABEL = re.compile(
    r"\bdate\s*&\s*time\s*:\s*(?P<value>[^\n]+)",
    re.IGNORECASE,
)


def _paragraph_text(paragraph):
    return "".join(run.text for run in paragraph.runs).strip()


def _image_bytes(paragraph):
    """
    Return the bytes of the first embedded picture in the paragraph,
    or None. Prefers the largest referenced image when several blips
    are present.
    """
    if paragraph is None:
        return None
    try:
        related_parts = paragraph.part.related_parts
    except AttributeError:
        return None

    candidates = []
    for blip in paragraph._element.findall(".//" + qn("a:blip")):
        r_id = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
        if not r_id:
            continue
        part = related_parts.get(r_id)
        if part is not None:
            candidates.append(part)

    if not candidates:
        return None

    candidates.sort(key=lambda part: len(part.blob), reverse=True)
    return candidates[0].blob


def _find_date_time_value(document):
    """
    Return the raw value after the 'Date & Time :' label, or None.

    When the venue label shares the same line (some reports put both on
    one paragraph), the value is cut off before the venue.
    """
    for _, paragraph in _iter_document_paragraphs(document):
        text = "".join(run.text for run in paragraph.runs)
        match = _DATE_TIME_LABEL.search(text)
        if not match:
            continue
        value = match.group("value").strip()
        if not value:
            continue
        value = re.split(r"\bvenue\b", value, maxsplit=1, flags=re.IGNORECASE)[0]
        value = value.strip()
        if value:
            return value
    return None


def _find_title(document, body_paragraphs):
    """
    Return the event title: the first non-empty paragraph after the
    'REPORT ON' header.
    """
    report_on = next(
        (
            index
            for index, paragraph in enumerate(body_paragraphs)
            if _paragraph_text(paragraph).lower() == "report on"
        ),
        None,
    )
    if report_on is None:
        return None

    start = report_on + 1
    for paragraph in body_paragraphs[start:]:
        text = _paragraph_text(paragraph)
        if not text:
            continue
        if _paragraph_has_image(paragraph):
            continue
        if _DATE_TIME_LABEL.search(text):
            continue
        if text.lower().startswith(_SECTION_BOUNDARIES):
            continue
        return text

    return None


def _find_description(document, body_paragraphs, start_from):
    """
    Return the report body text: the paragraphs that appear before the
    Poster / Photos / Enclosure sections.

    The description starts at the first non-empty paragraph after the
    metadata block (title / date & time / venue) and extends until the
    first image or section heading.
    """
    start = None
    for index in range(start_from, len(body_paragraphs)):
        text = _paragraph_text(body_paragraphs[index])
        if not text:
            continue
        if _paragraph_has_image(body_paragraphs[index]):
            continue
        if _DATE_TIME_LABEL.search(text):
            continue
        if text.lower().startswith(_SECTION_BOUNDARIES):
            continue
        start = index
        break

    if start is None:
        return ""

    stop = len(body_paragraphs)
    for index in range(start, len(body_paragraphs)):
        text = _paragraph_text(body_paragraphs[index])
        if _paragraph_has_image(body_paragraphs[index]):
            stop = index
            break
        if text.lower().startswith(_SECTION_BOUNDARIES):
            stop = index
            break

    parts = []
    for index in range(start, stop):
        text = _paragraph_text(body_paragraphs[index])
        if text:
            parts.append(text)

    return "\n\n".join(parts)


def extract_report_data(document, new_date=None, new_venue=None):
    """
    Pull every piece of information out of a finished report.

    Returns a dict with the raw values the generator consumes:
      date_time, date_token, venue, title, description,
      poster_image, photo1_image, photo2_image (bytes or None).

    When new_date / new_venue are given, the extracted date/venue are
    replaced by the corrected values.
    """
    body_paragraphs = document.paragraphs

    sections = find_image_sections(document)

    date_time = _find_date_time_value(document)

    venue_primary = find_primary_venue(document)

    title = _find_title(document, body_paragraphs)

    start_from = None
    if title is not None:
        for index, paragraph in enumerate(body_paragraphs):
            if _paragraph_text(paragraph) == title:
                start_from = index + 1
                break

    if start_from is None:
        report_on = next(
            (
                index
                for index, paragraph in enumerate(body_paragraphs)
                if _paragraph_text(paragraph).lower() == "report on"
            ),
            None,
        )
        if report_on is not None:
            start_from = report_on + 1
    if start_from is None:
        start_from = 0

    description = _find_description(
        document,
        body_paragraphs,
        start_from,
    )

    extracted = {
        "date_time": date_time,
        "venue": (
            venue_primary["value"]
            if venue_primary is not None
            else None
        ),
        "title": title,
        "description": description,
        "poster_image": _image_bytes(
            sections.get("poster_medium")
        ),
        "photo1_image": _image_bytes(
            sections.get("photo1")
        ),
        "photo2_image": _image_bytes(
            sections.get("photo2")
        ),
    }

    # ---------------- overrides --------------------------------

    if new_date is not None and date_time is not None:
        for match in _DATE_PATTERN.finditer(date_time):
            parsed = parse_date_match(match)
            if not parsed:
                continue
            _canonical, fmt = parsed
            new_canonical = parse_date_token(new_date)
            if new_canonical is None:
                raise ValueError(
                    f"Could not parse the corrected date: {new_date!r}"
                )
            replacement = render_new_date(fmt, new_canonical)
            extracted["date_time"] = (
                date_time[: match.start()]
                + replacement
                + date_time[match.end():]
            )
            break

    if new_venue is not None:
        extracted["venue"] = new_venue

    return extracted


def rewrite_dates_in_text(text, old_date_str, new_date_str):
    """
    Rewrite every date mention inside `text` whose calendar date equals
    the date read from old_date_str so it reflects the date read from
    new_date_str. Each mention keeps its original format (so
    "January 30, 2026", "30 Jan 2026" or "30-01-2026" are each
    reformatted consistently with the new date).

    No-op (returns text unchanged) when either string has no parsable
    date or the dates are identical.
    """
    if not text:
        return text

    old_canonical = parse_date_token(old_date_str)
    new_canonical = parse_date_token(new_date_str)

    if old_canonical is None or new_canonical is None:
        return text

    if old_canonical == new_canonical:
        return text

    parts = []
    last = 0

    for match in _DATE_PATTERN.finditer(text):

        parsed = parse_date_match(match)
        if parsed is None:
            continue

        canonical, fmt = parsed
        if canonical != old_canonical:
            continue

        parts.append(text[last:match.start()])
        parts.append(render_new_date(fmt, new_canonical))
        last = match.end()

    parts.append(text[last:])

    return "".join(parts)


def build_generation_data(extracted, image_overrides=None):
    """
    Turn an extraction result into the data dict the generator expects.

    image_overrides may map 'poster' / 'photo1' / 'photo2' to a path or
    a file-like object; extracted bytes are used for any key not listed.
    """
    image_overrides = image_overrides or {}

    def resolve(key, optional=False):
        if key in image_overrides and image_overrides[key] is not None:
            return image_overrides[key]
        bytes_value = extracted[f"{key}_image"]
        if bytes_value is None:
            if optional:
                return None
            raise ValueError(
                f"No {key} image available for regeneration."
            )
        return io.BytesIO(bytes_value)

    poster_medium = resolve("poster")

    if "poster_full" in image_overrides and image_overrides["poster_full"] is not None:
        poster_full = image_overrides["poster_full"]
    else:
        poster_full = resolve("poster")

    return {
        "DATE_TIME": extracted["date_time"] or "",
        "VENUE": extracted["venue"] or "",
        "EVENT_TITLE": extracted["title"] or "",
        "DESCRIPTION": extracted["description"] or "",
        "POSTER_MEDIUM": poster_medium,
        "EVENT_PHOTO_1": resolve("photo1"),
        "EVENT_PHOTO_2": resolve("photo2", optional=True),
        "POSTER_FULL": poster_full,
    }