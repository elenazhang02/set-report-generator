"""
SET Program Execution Report — Streamlit Web App
Run: streamlit run app.py
"""

import streamlit as st
import tempfile
import os
import datetime
import sys
import io
import contextlib

# Ensure the module can find its logos/ directory via __file__
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_supplier_ppt as v2

st.set_page_config(
    page_title="SET Execution Report Generator",
    page_icon="📊",
    layout="centered",
)

st.title("SET Execution Report Generator")
st.write("Upload your LPM image CSV export to generate a branded PowerPoint deck.")

# ── Inputs ─────────────────────────────────────────────────────────────────────

csv_file = st.file_uploader("LPM image CSV export", type=["csv"])

supplier_name = st.text_input(
    "Supplier name",
    placeholder="e.g. CAMPARI AMERICA",
    help=(
        "The supplier portion of the tracker_name field — everything before the program name. "
        "Leave blank to use the full tracker_name as the program name."
    ),
)

col1, col2 = st.columns(2)
with col1:
    include_supplier_logo = st.checkbox(
        "Include supplier logo",
        value=True,
        help="Pulled from local library or web (~5–15 s if not cached).",
    )
with col2:
    include_brand_logos = st.checkbox(
        "Include brand logos on overview",
        value=False,
        help="Adds brand logo strip to the overview slide (~10–30 s extra).",
    )

# ── Generate ───────────────────────────────────────────────────────────────────

if st.button("Generate PowerPoint", type="primary", disabled=not csv_file):
    log_buf = io.StringIO()

    with st.status("Building your deck…", expanded=True) as status:
        try:
            st.write("Reading CSV and deriving metadata…")

            with tempfile.TemporaryDirectory() as tmpdir:
                # Save the uploaded CSV to a real file path
                csv_path = os.path.join(tmpdir, csv_file.name)
                with open(csv_path, "wb") as f:
                    f.write(csv_file.getvalue())

                # Patch module-level globals the original CLI sets via config block
                v2.SUPPLIER_NAME = supplier_name.strip()
                v2.OUTPUT_DIR = tmpdir

                st.write("Fetching store images (parallel download)…")
                generated_at = datetime.datetime.now().astimezone()

                with contextlib.redirect_stdout(log_buf):
                    out_path = v2.build_deck(
                        csv_path,
                        include_supplier_logo,
                        include_brand_logos,
                        generated_at,
                    )

                st.write("Packaging file…")
                with open(out_path, "rb") as f:
                    pptx_bytes = f.read()
                filename = os.path.basename(out_path)
                slide_count = filename  # used below in success message

            status.update(label="Done!", state="complete")

            st.success(f"Deck ready — {filename}")
            st.download_button(
                label="Download PowerPoint",
                data=pptx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
            )

        except Exception as exc:
            status.update(label="Something went wrong", state="error")
            st.error(str(exc))

    # Show raw log in a collapsed expander for debugging
    log_text = log_buf.getvalue()
    if log_text:
        with st.expander("Build log"):
            st.code(log_text, language=None)
