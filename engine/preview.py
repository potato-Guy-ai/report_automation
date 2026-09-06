"""
Render a .docx as page preview images for the Streamlit UI.

Uses Microsoft Word (COM) to export the document to PDF, then MuPDF to
rasterise each page into a PNG. Falls back to returning None when Word
is unavailable so callers can show a simpler text preview instead.
"""

from pathlib import Path

import streamlit as st


def render_docx_preview(
    docx_bytes,
    temp_dir,
    max_pages=10,
    zoom=1.6,
):
    """
    Return a list of PNG bytes, one per page of the document.

    temp_dir must be a Path to a writable temporary directory.
    Returns None when rendering is not possible (Word missing/failed).
    """
    try:
        import win32com.client
        import pymupdf
    except ImportError:
        return None

    try:

        docx_path = Path(temp_dir) / "preview_report.docx"
        pdf_path = Path(temp_dir) / "preview_report.pdf"

        docx_path.write_bytes(docx_bytes)

    except Exception:
        return None

    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        document = word.Documents.Open(
            str(docx_path),
            ReadOnly=True,
            AddToRecentFiles=False,
        )

        try:
            document.ExportAsFixedFormat(
                OutputFileName=str(pdf_path),
                ExportFormat=17,
            )
        finally:
            document.Close(False)
    except Exception:
        return None
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        word = None

    if not pdf_path.exists():
        return None

    try:
        import pymupdf

        pages = []
        with pymupdf.open(pdf_path) as pdf:
            for page in pdf:
                if len(pages) >= max_pages:
                    break
                matrix = pymupdf.Matrix(zoom, zoom)
                pixmap = page.get_pixmap(matrix=matrix)
                pages.append(pixmap.tobytes("png"))
        return pages
    except Exception:
        return None


def show_preview(pages):
    """
    Display a document preview storyboard inside Streamlit.

    Falls back to a friendly message when pages could not be rendered.
    """
    if not pages:
        st.caption(
            "Page preview could not be rendered on this machine. "
            "You can still download the document."
        )
        return

    total = len(pages)
    st.subheader(f"Preview ({total} page{'s' if total != 1 else ''})")

    st.caption(
        "Scroll through the pages below, then download the report above."
    )

    columns = st.columns(2)

    for index, png in enumerate(pages):
        with columns[index % 2]:
            st.image(
                png,
                caption=f"Page {index + 1}",
                use_container_width=True,
            )