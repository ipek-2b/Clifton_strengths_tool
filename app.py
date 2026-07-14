"""
CliftonStrengths Pipeline — Web App
2b Limitless Internship | AI & Automation

A drag-and-drop front end for clifton_strengths_pipeline.py. Anyone on the
team can open this page, drop in their CliftonStrengths files (PDF reports,
Excel/CSV trackers, or pasted-text .txt files — any mix), and download four
files back: the data spreadsheet, the wheel-chart spreadsheet, the workshop
Word workbook, and printable table name cards.

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
    build_workshop_docx,
    build_name_cards_docx,
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

col1, col2, col3 = st.columns(3)
with col1:
    generate_charts = st.checkbox("Wheel-chart spreadsheet", value=True)
with col2:
    generate_workbook = st.checkbox("Workshop Word workbook", value=True)
with col3:
    generate_cards = st.checkbox("Table name cards", value=True)

uploaded_files = st.file_uploader(
    "Drop files here",
    type=["pdf", "xlsx", "xls", "csv", "txt"],
    accept_multiple_files=True,
)

# Everything produced by "Process files" is stashed in session_state. Streamlit
# reruns the whole script on ANY button click — including the download buttons
# below — so without this, clicking "Download data spreadsheet" would wipe out
# the other download buttons and force starting over from "Process files".
if "results" not in st.session_state:
    st.session_state.results = None

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
                st.session_state.results = None
                st.error("Couldn't parse any of the uploaded files. Details below.")
                st.code(log_buffer.getvalue() or "No details available.")
            else:
                data_path = tmp_dir / "CliftonStrengths_Data.xlsx"
                wheels_path = tmp_dir / "CliftonStrengths_Data_Wheels.xlsx"
                workbook_path = tmp_dir / "Workshop_Workbook.docx"
                cards_path = tmp_dir / "Table_Name_Cards.docx"

                merged = build_workbook(reports, data_path)

                log_text = log_buffer.getvalue().strip()

                results = {
                    "num_reports": len(reports),
                    "people": sorted(merged.keys()),
                    "log_text": log_text,
                    "data_bytes": data_path.read_bytes(),
                    "wheels_bytes": None,
                    "workbook_bytes": None,
                    "cards_bytes": None,
                }

                if generate_charts:
                    with st.spinner("Generating wheel charts..."):
                        build_wheels_workbook(merged, wheels_path)
                    results["wheels_bytes"] = wheels_path.read_bytes()

                if generate_workbook:
                    with st.spinner("Building the workshop workbook..."):
                        build_workshop_docx(merged, workbook_path)
                    results["workbook_bytes"] = workbook_path.read_bytes()

                if generate_cards:
                    with st.spinner("Building table name cards..."):
                        build_name_cards_docx(merged, cards_path)
                    results["cards_bytes"] = cards_path.read_bytes()

                st.session_state.results = results
else:
    st.session_state.results = None
    st.info("Upload at least one file to get started.")

# Render results (and download buttons) from session_state, independent of the
# "Process files" click that produced them, so they survive reruns triggered by
# clicking a download button.
results = st.session_state.results
if results:
    st.success(f"Processed {results['num_reports']} file(s) — {len(results['people'])} people total.")

    if "SKIPPED" in results["log_text"] or "WARNING" in results["log_text"]:
        with st.expander("Some rows needed attention — click to view"):
            st.code(results["log_text"])

    st.download_button(
        "⬇️ Download data spreadsheet",
        data=results["data_bytes"],
        file_name="CliftonStrengths_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if results["wheels_bytes"] is not None:
        st.download_button(
            "⬇️ Download wheel charts",
            data=results["wheels_bytes"],
            file_name="CliftonStrengths_Data_Wheels.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if results["workbook_bytes"] is not None:
        st.download_button(
            "⬇️ Download workshop workbook",
            data=results["workbook_bytes"],
            file_name="Workshop_Workbook.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    if results["cards_bytes"] is not None:
        st.download_button(
            "⬇️ Download table name cards",
            data=results["cards_bytes"],
            file_name="Table_Name_Cards.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    with st.expander("Preview parsed people"):
        st.write(results["people"])
