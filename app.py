"""
CliftonStrengths Pipeline — Web App
2b Limitless Internship | AI & Automation

A drag-and-drop front end for clifton_strengths_pipeline.py. Anyone on the
team can open this page, drop in their CliftonStrengths files (PDF reports,
Excel/CSV trackers, or pasted-text .txt files — any mix), and download two
files back: the data spreadsheet and the wheel-chart spreadsheet.

Run locally:
    streamlit run app.py

Deploy for the whole team: see README.md in this folder.
"""

import contextlib
import io
import tempfile
from pathlib import Path

import streamlit as st

from clifton_strengths_pipeline import (
    collect_reports,
    build_workbook,
    build_wheels_workbook,
)

st.set_page_config(page_title="CliftonStrengths Pipeline", page_icon="🌀", layout="centered")

st.title("🌀 CliftonStrengths Pipeline")
st.write(
    "Upload everyone's CliftonStrengths files below — PDF reports, an Excel/CSV "
    "tracker, or pasted-text files, any mix, all at once. This replaces manually "
    "typing 34 rank numbers per person."
)

with st.expander("What files can I upload?"):
    st.markdown(
        "- **PDF** — the Gallup emailed report for one person\n"
        "- **XLSX / CSV** — a tracker with one column per theme name and rank "
        "numbers underneath, or an export in a few other common layouts\n"
        "- **TXT** — one person's results pasted from the Gallup portal, one file per person\n\n"
        "You can upload files for multiple people at once, in any combination of these formats."
    )

generate_charts = st.checkbox("Also generate the wheel-chart file", value=True)

uploaded_files = st.file_uploader(
    "Drop files here",
    type=["pdf", "xlsx", "xls", "csv", "txt"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"{len(uploaded_files)} file(s) ready.")

    if st.button("Process files", type="primary"):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            input_dir = tmp_dir / "input"
            input_dir.mkdir()
            for f in uploaded_files:
                (input_dir / f.name).write_bytes(f.getvalue())

            log_buffer = io.StringIO()
            with st.spinner("Reading files and building the spreadsheet..."):
                with contextlib.redirect_stdout(log_buffer), contextlib.redirect_stderr(log_buffer):
                    reports = collect_reports(input_dir)

            if not reports:
                st.error("Couldn't parse any of the uploaded files. Details below.")
                st.code(log_buffer.getvalue() or "No details available.")
            else:
                data_path = tmp_dir / "CliftonStrengths_Data.xlsx"
                wheels_path = tmp_dir / "CliftonStrengths_Data_Wheels.xlsx"

                merged = build_workbook(reports, data_path)

                st.success(f"Processed {len(reports)} file(s) — {len(merged)} people total.")

                log_text = log_buffer.getvalue().strip()
                if "SKIPPED" in log_text or "WARNING" in log_text:
                    with st.expander("Some rows needed attention — click to view"):
                        st.code(log_text)

                st.download_button(
                    "⬇️ Download data spreadsheet",
                    data=data_path.read_bytes(),
                    file_name="CliftonStrengths_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                if generate_charts:
                    with st.spinner("Generating wheel charts..."):
                        build_wheels_workbook(merged, wheels_path)
                    st.download_button(
                        "⬇️ Download wheel charts",
                        data=wheels_path.read_bytes(),
                        file_name="CliftonStrengths_Data_Wheels.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                with st.expander("Preview parsed people"):
                    st.write(sorted(merged.keys()))
else:
    st.info("Upload at least one file to get started.")
