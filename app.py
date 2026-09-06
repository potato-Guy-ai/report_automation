import streamlit as st
from pathlib import Path
import tempfile

from engine.template_loader import load_template
from engine.content_inserter import replace_placeholders
from engine.preview import render_docx_preview, show_preview
from ui_correction import render_correction_ui


# ============================================================
# CONFIGURATION
# ============================================================

BASE = Path(__file__).parent

TEMPLATE = BASE / "templates" / "final_sample_report.docx"


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Report Automation",
    page_icon="📄",
    layout="centered"
)


# ============================================================
# HEADER
# ============================================================

st.title("📄 Report Automation System")

st.write(
    "Create a professional event report automatically using the "
    "provided Word template, or pull the contents out of an "
    "already-finished report, fix the wrong values and regenerate "
    "a clean, properly aligned document."
)

st.divider()


# ============================================================
# MODE SELECTION
# ============================================================

mode = st.radio(
    "Mode",
    [
        "Generate from template",
        "Regenerate from existing report",
    ],
    horizontal=True,
)

if mode == "Regenerate from existing report":
    render_correction_ui()
    st.stop()


# ============================================================
# EVENT DETAILS
# ============================================================

st.subheader("Event Details")

event_title = st.text_input(
    "Event Title",
    placeholder="Example: Automation Programming Workshop"
)

date_time = st.text_input(
    "Date & Time",
    placeholder="Example: 30-01-2026 & 01.30 pm to 5 pm"
)

venue = st.text_input(
    "Venue",
    placeholder="Example: CSE LAB – 1"
)


# ============================================================
# REPORT CONTENT
# ============================================================

st.subheader("Report Content")

description = st.text_area(
    "Event Description",
    placeholder=(
        "Enter the complete report content here...\n\n"
        "You can write multiple paragraphs and as much content "
        "as required."
    ),
    height=300
)


st.divider()


# ============================================================
# IMAGES
# ============================================================

st.subheader("Event Images")

poster = st.file_uploader(
    "Upload Poster",
    type=["jpg", "jpeg", "png"],
    help="This poster will be used for both the medium and full poster sections."
)

photo1 = st.file_uploader(
    "Upload Event Photo 1",
    type=["jpg", "jpeg", "png"]
)

photo2 = st.file_uploader(
    "Upload Event Photo 2",
    type=["jpg", "jpeg", "png"]
)


st.divider()


# ============================================================
# IMAGE PREVIEW
# ============================================================

if poster:
    st.write("**Poster Preview**")
    st.image(
        poster,
        width=300
    )

if photo1:
    st.write("**Event Photo 1 Preview**")
    st.image(
        photo1,
        width=300
    )

if photo2:
    st.write("**Event Photo 2 Preview**")
    st.image(
        photo2,
        width=300
    )


# ============================================================
# GENERATE REPORT
# ============================================================

st.divider()

generate = st.button(
    "🚀 Generate Report",
    type="primary",
    use_container_width=True
)


if generate:

    # --------------------------------------------------------
    # VALIDATION
    # --------------------------------------------------------

    missing_fields = []

    if not event_title.strip():
        missing_fields.append("Event Title")

    if not date_time.strip():
        missing_fields.append("Date & Time")

    if not venue.strip():
        missing_fields.append("Venue")

    if not description.strip():
        missing_fields.append("Event Description")

    if not poster:
        missing_fields.append("Poster")

    if not photo1:
        missing_fields.append("Event Photo 1")

    if missing_fields:

        st.error(
            "Please provide: "
            + ", ".join(missing_fields)
        )

        st.stop()


    # --------------------------------------------------------
    # TEMPORARY IMAGE STORAGE
    # --------------------------------------------------------

    with tempfile.TemporaryDirectory() as temp_dir:

        temp_dir = Path(temp_dir)

        poster_path = temp_dir / "poster.jpg"
        photo1_path = temp_dir / "photo1.jpg"
        photo2_path = temp_dir / "photo2.jpg"

        poster_path.write_bytes(
            poster.getbuffer()
        )

        photo1_path.write_bytes(
            photo1.getbuffer()
        )

        photo2_value = None

        if photo2 is not None:
            photo2_path.write_bytes(
                photo2.getbuffer()
            )
            photo2_value = photo2_path


        # ----------------------------------------------------
        # REPORT DATA
        # ----------------------------------------------------

        data = {

            "DATE_TIME": date_time,

            "VENUE": venue,

            "EVENT_TITLE": event_title,

            "DESCRIPTION": description,

            "POSTER_MEDIUM": poster_path,

            "EVENT_PHOTO_1": photo1_path,

            "EVENT_PHOTO_2": photo2_value,

            "POSTER_FULL": poster_path,
        }


        # ----------------------------------------------------
        # GENERATE DOCUMENT
        # ----------------------------------------------------

        try:

            with st.spinner(
                "Generating your report..."
            ):

                document = load_template(
                    TEMPLATE
                )

                document = replace_placeholders(
                    document,
                    data
                )


                output_path = temp_dir / "generated_report.docx"

                document.save(
                    output_path
                )


            # ------------------------------------------------
            # SUCCESS
            # ------------------------------------------------

            st.success(
                "✅ Report generated successfully!"
            )


            # ------------------------------------------------
            # DOWNLOAD
            # ------------------------------------------------

            report_bytes = output_path.read_bytes()

            st.download_button(
                label="⬇️ Download Report",
                data=report_bytes,
                file_name="generated_report.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
                use_container_width=True
            )

            # ------------------------------------------------
            # PREVIEW
            # ------------------------------------------------

            pages = render_docx_preview(report_bytes, temp_dir)

            show_preview(pages)


        except Exception as e:

            st.error(
                "Something went wrong while generating "
                "the report."
            )

            st.exception(e)