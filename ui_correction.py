"""
Streamlit UI for the "Correct existing report" workflow.

Kept as a separate module so it can be added/removed without affecting
the template-based generation flow in app.py.
"""

import streamlit as st
from pathlib import Path
import tempfile
from io import BytesIO

from docx import Document

from engine.corrector import (
    apply_occurrences,
    apply_image_replacements,
    detect_corrections,
    find_image_sections,
)


def render_correction_ui():

    st.subheader("Correct Existing Report")

    st.write(
        "Upload a finished report. The tool detects the event "
        "date from the `Date & Time` field and lists every date "
        "in the document that matches it, so you can confirm "
        "exactly what gets changed. Time, venue and everything "
        "else are left untouched."
    )

    uploaded = st.file_uploader(
        "Finished report (.docx)",
        type=["docx"],
    )

    new_date = st.text_input(
        "Corrected event date",
        placeholder="e.g. 23-02-2026",
        help=(
            "Only dates equal to the primary event date are updated, "
            "each in its own original format "
            "(12-02-2025 -> 23-02-2026, "
            "12th February 2025 -> 23rd February 2026)."
        ),
    )

    st.subheader("Replace images (optional)")

    st.write(
        "Leave a field empty to keep the existing image in the "
        "report unchanged."
    )

    poster = st.file_uploader(
        "Replace Poster (medium + full poster)",
        type=["jpg", "jpeg", "png"],
    )

    photo1 = st.file_uploader(
        "Replace Event Photo 1",
        type=["jpg", "jpeg", "png"],
    )

    photo2 = st.file_uploader(
        "Replace Event Photo 2",
        type=["jpg", "jpeg", "png"],
    )

    detect = st.button(
        "Detect dates",
        use_container_width=True,
    )

    if uploaded and new_date.strip() and detect:

        try:

            blob = uploaded.getvalue()

            document = Document(BytesIO(blob))

            result = detect_corrections(
                document,
                new_date.strip(),
            )

            image_sections = find_image_sections(document)

            st.session_state["correct_blob"] = blob
            st.session_state["correct_name"] = uploaded.name
            st.session_state["correct_new_date"] = new_date.strip()
            st.session_state["correct_result"] = result
            st.session_state["correct_image_sections"] = image_sections

        except Exception as e:

            st.exception(e)

    result = st.session_state.get("correct_result")

    if result is None and not (uploaded and new_date.strip() and detect):

        st.info(
            "Upload a report and enter the corrected date, "
            "then press 'Detect dates'."
        )
        return

    if result is None:
        return

    primary = result.get("primary")

    if primary is None:

        st.error(
            "Could not locate a 'Date & Time' field or parse "
            "a date from it."
        )
        return

    changed = [
        match
        for match in result["matches"]
        if match.get("changed")
    ]

    st.info(
        "Detected event date: "
        f"**{primary['token']}** "
        f"({primary['location']})"
    )

    image_sections = st.session_state.get("correct_image_sections", {})

    section_labels = {
        "poster_medium": "Poster",
        "photo1": "Event Photo 1",
        "photo2": "Event Photo 2",
        "poster_full": "Full Poster",
    }

    if image_sections:

        found_labels = [
            section_labels[key]
            for key in ("poster_medium", "photo1", "photo2", "poster_full")
            if key in image_sections
        ]

        st.write(
            "Images found in the report: "
            + ", ".join(f"**{label}**" for label in found_labels)
            + ". Upload replacements above to swap them."
        )

    else:

        st.write(
            "No poster/photo images were found in this report, "
            "so image replacement will be skipped."
        )

    confirmed = []

    if not changed:

        st.success(
            "No dates in the report need changing. "
            "You can still replace images below."
        )

    else:

        st.subheader("Dates to update")

        for index, match in enumerate(changed):

            widget_key = f"correct_apply_{index}"

            if widget_key not in st.session_state:
                st.session_state[widget_key] = True

            checked = st.checkbox(
                (
                    f"**{match['original']}** → **{match['new']}**"
                    f"  —  {match['location']}"
                ),
                key=widget_key,
                help=match["snippet"],
            )

            if checked:
                confirmed.append(match)

        st.caption(
            f"{len(confirmed)} of {len(changed)} "
            "date(s) selected for correction."
        )

    blob = st.session_state.get("correct_blob")
    original_name = st.session_state.get("correct_name", "report.docx")
    correction_date = st.session_state.get("correct_new_date", new_date)

    generate = st.button(
        "Generate corrected report",
        type="primary",
        use_container_width=True,
    )

    if not (generate and blob):
        return

    # ------------------------------------------------------------
    # Collect optional image replacements
    # ------------------------------------------------------------

    image_mapping = {}

    if poster is not None:
        image_mapping["poster"] = BytesIO(poster.getbuffer())
    if photo1 is not None:
        image_mapping["photo1"] = BytesIO(photo1.getbuffer())
    if photo2 is not None:
        image_mapping["photo2"] = BytesIO(photo2.getbuffer())

    # ------------------------------------------------------------
    # Confirm when no poster / photos were provided
    # ------------------------------------------------------------

    confirmed_no_images = st.session_state.get(
        "correct_no_images_ok",
        False,
    )

    if not image_mapping and not confirmed_no_images:

        st.warning(
            "No poster or photos were provided. The existing "
            "poster and photos in the report will be left unchanged."
        )

        if st.button(
            "Yes, continue without poster/photos",
            use_container_width=True,
        ):
            st.session_state["correct_no_images_ok"] = True
        else:
            return

    try:

        with st.spinner("Correcting report..."):

            document = Document(BytesIO(blob))

            result = detect_corrections(
                document,
                correction_date,
            )

            confirmed_originals = {
                match["original"]
                for match in confirmed
            }

            to_apply = [
                match
                for match in result["matches"]
                if (
                    match.get("changed")
                    and match["original"] in confirmed_originals
                )
            ]

            apply_occurrences(document, to_apply)

            images_replaced = 0

            if image_mapping:

                images_replaced = apply_image_replacements(
                    document,
                    image_mapping,
                )

            with tempfile.TemporaryDirectory() as temp_dir:

                temp_dir = Path(temp_dir)

                file_stem = Path(original_name).stem
                corrected_path = (
                    temp_dir / f"corrected_{file_stem}.docx"
                )

                document.save(corrected_path)

                corrected_bytes = corrected_path.read_bytes()

        st.session_state["correct_no_images_ok"] = False

        message = "✅ Report corrected successfully!"

        if image_mapping:
            message += f" ({images_replaced} image(s) replaced)"

        st.success(message)

        st.download_button(
            label="⬇️ Download Corrected Report",
            data=corrected_bytes,
            file_name=f"corrected_{file_stem}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

    except Exception as e:

        st.error(
            "Something went wrong while correcting "
            "the report."
        )

        st.exception(e)