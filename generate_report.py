from pathlib import Path

from engine.template_loader import load_template
from engine.content_inserter import replace_placeholders


# ============================================================
# PATHS
# ============================================================

BASE = Path(__file__).parent

TEMPLATE = BASE / "templates" / "final_sample_report.docx"

OUTPUT_DIR = BASE / "output"
OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

OUTPUT = OUTPUT_DIR / "generated_report.docx"


# ============================================================
# REPORT DATA
# ============================================================

data = {

    "DATE_TIME": (
        "30-01-2026 & 01.30 am to 5 pm"
    ),

    "VENUE": (
        "CSE LAB – 1"
    ),

    "EVENT_TITLE": (
        "AUTOMATION PROGRAMMING WORKSHOP"
    ),

    "DESCRIPTION": (
        "The Automation Programming Workshop was organized "
        "to introduce students to automation concepts and "
        "programming techniques.\n\n"

        "The session covered fundamental concepts, practical "
        "demonstrations and real-world applications of "
        "automation programming.\n\n"

        "Students actively participated in the session and "
        "gained practical exposure through demonstrations "
        "and interactive activities."
    ),

    # --------------------------------------------------------
    # IMAGES
    # --------------------------------------------------------

    "POSTER_MEDIUM": (
        BASE / "images" / "poster.jpeg"
    ),

    "EVENT_PHOTO_1": (
        BASE / "images" / "photo1.jpeg"
    ),

    "EVENT_PHOTO_2": (
        BASE / "images" / "photo2.jpeg"
    ),

    "POSTER_FULL": (
        BASE / "images" / "poster.jpeg"
    ),
}


# ============================================================
# GENERATE REPORT
# ============================================================

document = load_template(TEMPLATE)

document = replace_placeholders(
    document,
    data
)

document.save(OUTPUT)


print()
print("==========================================")
print(" REPORT GENERATED SUCCESSFULLY")
print("==========================================")
print(f"Output : {OUTPUT}")
print()