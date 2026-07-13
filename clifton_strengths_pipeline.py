"""
CliftonStrengths 34 -> Spreadsheet Pipeline
2b Limitless Internship | AI & Automation

Extracts each person's rank-ordered 34 CliftonStrengths themes and compiles
everyone into a single spreadsheet, replacing manual retyping of 34 data
points per person.

Accepts THREE input formats, since it's unconfirmed which one the team
actually has for each person — drop any mix of these into one folder:

  1. PDF   — the Gallup emailed report (e.g. "Kucukerdogan-Ipek-ALL_34.pdf")
  2. XLSX/CSV — a spreadsheet export or manually-built tracker. Handles:
       - 2b Limitless's actual format: one column per theme name ("Strategic",
         "Competition", ... "Includer"), one row per person, cell = that
         person's rank number (1-34) for that theme — this is what the team
         manually types in from the PDF today.
       - "long" format: one row per theme, with Name/Theme/Rank-ish columns
       - "wide" format: one row per person, 34 columns labelled "Rank 1".."Rank 34",
         each holding a theme name
       - a bare single column of 34 theme names (one person, name = filename)
  3. TXT   — pasted text copied straight from the Gallup portal page.
       Put one person's pasted results per .txt file. Works whether the
       paste kept the "1. Strategic" numbering or is just 34 theme names
       on separate lines in rank order.

USAGE
-----
    python3 clifton_strengths_pipeline.py --input "path/to/reports_folder" --output "CliftonStrengths_Data.xlsx"
    python3 clifton_strengths_pipeline.py --input "path/to/single_report.pdf" --output "CliftonStrengths_Data.xlsx"

Re-running on the same output file merges in new/updated people (matched by
name) and leaves everyone else untouched, so you can keep dropping files in
the folder over time and re-run.

OUTPUT — TWO SEPARATE FILES
----------------------------
1. The data workbook (--output), laid out exactly like the team's own tracker:
     - "CliftonStrengths Data"     — domain banner row, then 34 theme-name columns,
                                      one row per person, cell = that person's RANK
                                      NUMBER (1-34) for that theme
     - "Domain Summary (Top 10)"   — top-10 theme counts per domain, leading domain
2. The chart workbook (--charts-output, defaults to "<output>_Wheels.xlsx"):
     - "Strengths Wheels"          — a domain-quadrant radar chart per person, 2 per
                                      row, each with a name banner. A manual page
                                      break is inserted after every row of charts so
                                      one never gets sliced across a printed page.
   Skip this file entirely with --no-charts for faster runs on large batches.

REQUIREMENTS
------------
pip install pdfplumber openpyxl pandas matplotlib --break-system-packages
"""

import argparse
import re
import sys
import tempfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pdfplumber
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

# ---------------------------------------------------------------------------
# Fixed Gallup CliftonStrengths 34 -> Domain mapping (does not vary by person)
# ---------------------------------------------------------------------------
DOMAIN_MAP = {
    "Achiever": "Executing", "Arranger": "Executing", "Belief": "Executing",
    "Consistency": "Executing", "Deliberative": "Executing", "Discipline": "Executing",
    "Focus": "Executing", "Responsibility": "Executing", "Restorative": "Executing",
    "Activator": "Influencing", "Command": "Influencing", "Communication": "Influencing",
    "Competition": "Influencing", "Maximiser": "Influencing", "Maximizer": "Influencing",
    "Self-Assurance": "Influencing", "Significance": "Influencing", "Woo": "Influencing",
    "Adaptability": "Relationship Building", "Connectedness": "Relationship Building",
    "Developer": "Relationship Building", "Empathy": "Relationship Building",
    "Harmony": "Relationship Building", "Includer": "Relationship Building",
    "Individualisation": "Relationship Building", "Individualization": "Relationship Building",
    "Positivity": "Relationship Building", "Relator": "Relationship Building",
    "Analytical": "Strategic Thinking", "Context": "Strategic Thinking",
    "Futuristic": "Strategic Thinking", "Ideation": "Strategic Thinking",
    "Input": "Strategic Thinking", "Intellection": "Strategic Thinking",
    "Learner": "Strategic Thinking", "Strategic": "Strategic Thinking",
}
DOMAIN_ORDER = ["Executing", "Influencing", "Relationship Building", "Strategic Thinking"]

# Canonical column order used for BOTH the spreadsheet and the wheel chart: themes
# grouped by domain (in DOMAIN_ORDER), alphabetical within each domain — this is the
# exact column order the team's own tracker uses.
DOMAIN_THEMES_ORDERED = {d: [] for d in DOMAIN_ORDER}
for _theme, _domain in DOMAIN_MAP.items():
    if _theme in ("Maximizer", "Individualization"):  # skip US-spelling duplicates
        continue
    DOMAIN_THEMES_ORDERED[_domain].append(_theme)
for _d in DOMAIN_THEMES_ORDERED:
    DOMAIN_THEMES_ORDERED[_d].sort()

ORDERED_THEMES = [theme for d in DOMAIN_ORDER for theme in DOMAIN_THEMES_ORDERED[d]]

# Sort longest-first so "Self-Assurance" matches before a stray "Self" would.
_THEME_ALTERNATION = "|".join(sorted(DOMAIN_MAP.keys(), key=len, reverse=True))
_RANK_PATTERN = re.compile(rf'\b(\d{{1,2}})\.\s*({_THEME_ALTERNATION})\b')
_BARE_THEME_PATTERN = re.compile(rf'^({_THEME_ALTERNATION})$', re.IGNORECASE)


def _normalise_theme(raw: str) -> str | None:
    """Match a loose string (extra spaces, punctuation, wrong case) to a canonical theme name."""
    cleaned = re.sub(r'[^A-Za-z\- ]', '', str(raw)).strip()
    for theme in DOMAIN_MAP:
        if cleaned.lower() == theme.lower():
            return theme
    return None


# ---------------------------------------------------------------------------
# Format 1: PDF (Gallup emailed report)
# ---------------------------------------------------------------------------
def extract_from_pdf(pdf_path: Path) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        page1_text = pdf.pages[0].extract_text() or ""

    header_line = page1_text.splitlines()[0]
    if "|" in header_line:
        name, date = [part.strip() for part in header_line.split("|", 1)]
    else:
        name, date = header_line.strip(), ""

    ranks = {}
    for num, theme in _RANK_PATTERN.findall(page1_text):
        n = int(num)
        if 1 <= n <= 34 and n not in ranks:
            ranks[n] = theme.strip()

    _require_complete(ranks, pdf_path.name)
    return {"name": name, "date": date, "ranks": ranks, "source_file": pdf_path.name}


# ---------------------------------------------------------------------------
# Format 2: pasted text from the Gallup portal (.txt, one person per file)
# ---------------------------------------------------------------------------
def extract_from_text(txt_path: Path) -> dict:
    text = txt_path.read_text(encoding="utf-8", errors="ignore")
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    name = txt_path.stem.replace("_", " ").replace("-", " ").strip()
    date = ""
    # If the first line looks like "NAME | DATE" or "NAME - DATE", use it and drop it.
    if lines and ("|" in lines[0] or re.search(r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', lines[0])):
        header = lines[0]
        if "|" in header:
            parts = [p.strip() for p in header.split("|", 1)]
            name, date = parts[0], parts[1] if len(parts) > 1 else ""
        lines = lines[1:]

    numbered_text = "\n".join(lines)
    ranks = {}
    for num, theme in _RANK_PATTERN.findall(numbered_text):
        n = int(num)
        if 1 <= n <= 34 and n not in ranks:
            ranks[n] = theme.strip()

    # Fallback: no "N. Theme" numbering present — assume each line is one theme,
    # already in rank order (this is what a plain copy-paste of a portal list looks like).
    if len(ranks) < 34:
        ranks = {}
        rank_n = 1
        for line in lines:
            theme = _normalise_theme(line)
            if theme is None:
                # allow trailing junk on the line, e.g. "1 Strategic" already handled above;
                # try stripping a leading rank number/bullet before giving up
                stripped = re.sub(r'^[\d]{1,2}[\.\)]?\s*', '', line)
                theme = _normalise_theme(stripped)
            if theme:
                ranks[rank_n] = theme
                rank_n += 1
            if rank_n > 34:
                break

    _require_complete(ranks, txt_path.name)
    return {"name": name, "date": date, "ranks": ranks, "source_file": txt_path.name}


# ---------------------------------------------------------------------------
# Format 3: Excel / CSV export from the Gallup portal
# ---------------------------------------------------------------------------
def _load_raw_grid(path: Path) -> list:
    """Every cell value, row by row, with NO assumption about which row is the header."""
    if path.suffix.lower() == ".csv":
        import csv
        with open(path, newline="", encoding="utf-8-sig") as f:
            return [row for row in csv.reader(f)]
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    return [[c for c in row] for row in ws.iter_rows(values_only=True)]


def _extract_team_tracker_format(path: Path) -> list:
    """
    THIS IS THE ACTUAL 2b LIMITLESS LAYOUT (confirmed from their live tracker):
    one column per CliftonStrengths theme, grouped visually under domain banners
    (EXECUTING / INFLUENCING / RELATIONSHIP BUILDING / STRATEGIC THINKING), with
    the theme names themselves living a few rows down — the tracker has THREE
    stacked header rows (domain banner, domain description, theme name) before
    the data starts, and often a totals row at the bottom. This scans for
    whichever row actually contains the 34 theme names, wherever it is, and
    ignores rows that don't look like a real person (blank name, totals row).
    """
    grid = _load_raw_grid(path)

    header_row_idx, best_count = None, 0
    for i in range(min(10, len(grid))):
        count = sum(1 for cell in grid[i] if _normalise_theme(cell))
        if count > best_count:
            header_row_idx, best_count = i, count

    if header_row_idx is None or best_count < 25:
        return []  # doesn't look like this format — let the caller try something else

    header_row = grid[header_row_idx]
    theme_cols = [(idx, _normalise_theme(v)) for idx, v in enumerate(header_row) if _normalise_theme(v)]
    theme_col_idxs = {idx for idx, _ in theme_cols}

    name_idx = None
    for idx, v in enumerate(header_row):
        if isinstance(v, str) and re.search(r'name|team member|employee', v, re.IGNORECASE):
            name_idx = idx
            break
    if name_idx is None:
        left_of_themes = [idx for idx in range(len(header_row)) if idx not in theme_col_idxs]
        name_idx = left_of_themes[0] if left_of_themes else 0

    reports = []
    for row in grid[header_row_idx + 1:]:
        if name_idx >= len(row):
            continue
        name_val = row[name_idx]
        if not isinstance(name_val, str) or not name_val.strip():
            continue  # skips blank rows and the numeric totals row at the bottom

        ranks = {}
        for idx, theme in theme_cols:
            if idx >= len(row) or row[idx] is None:
                continue
            try:
                rank_num = int(float(row[idx]))
            except (ValueError, TypeError):
                continue
            if 1 <= rank_num <= 34:
                ranks[rank_num] = theme

        if len(ranks) < 34:
            missing = [n for n in range(1, 35) if n not in ranks]
            print(f"WARNING '{name_val.strip()}' in {path.name}: missing rank(s) {missing} — row skipped",
                  file=sys.stderr)
            continue

        reports.append({"name": name_val.strip(), "date": "", "ranks": ranks, "source_file": path.name})

    return reports


def extract_from_table(path: Path) -> list:
    """Returns a LIST of report dicts — a single spreadsheet export may contain many people."""
    team_tracker_reports = _extract_team_tracker_format(path)
    if team_tracker_reports:
        return team_tracker_reports

    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)

    cols_lower = {c: str(c).strip().lower() for c in df.columns}

    def find_col(*keywords):
        for orig, low in cols_lower.items():
            if any(k in low for k in keywords):
                return orig
        return None

    name_col = find_col("name")
    theme_col = find_col("theme", "talent", "strength")
    rank_col = find_col("rank", "sequence", "position", "order")

    reports = []

    if name_col and theme_col and rank_col:
        # "Long" format: one row per (person, theme) pair.
        for person, group in df.groupby(name_col):
            group_sorted = group.sort_values(rank_col)
            ranks = {}
            for i, (_, r) in enumerate(group_sorted.iterrows(), start=1):
                theme = _normalise_theme(r[theme_col])
                rank_val = int(r[rank_col]) if str(r[rank_col]).strip().isdigit() else i
                if theme:
                    ranks[rank_val] = theme
            _require_complete(ranks, f"{path.name} ({person})")
            reports.append({"name": str(person).strip(), "date": "", "ranks": ranks, "source_file": path.name})
        return reports

    # "Wide" format: one row per person, 34 theme columns (e.g. "Rank 1".."Rank 34",
    # "1".."34", or "Theme 1".."Theme 34") — the cell holds the THEME NAME for that rank.
    rank_like_cols = []
    for c in df.columns:
        m = re.search(r'(\d{1,2})\s*$', str(c))
        if m and 1 <= int(m.group(1)) <= 34:
            rank_like_cols.append((int(m.group(1)), c))

    if len(rank_like_cols) >= 30:  # tolerate a couple missing/misnamed columns
        rank_like_cols.sort()
        for _, row in df.iterrows():
            person = str(row[name_col]).strip() if name_col else f"Person row {_ + 2}"
            ranks = {}
            for n, col in rank_like_cols:
                theme = _normalise_theme(row[col])
                if theme:
                    ranks[n] = theme
            _require_complete(ranks, f"{path.name} ({person})")
            reports.append({"name": person, "date": "", "ranks": ranks, "source_file": path.name})
        return reports

    # Bare single column of 34 theme names, in rank order, for ONE person.
    first_col = df.columns[0]
    values = [v for v in df[first_col].tolist() if pd.notna(v)]
    ranks = {}
    for i, v in enumerate(values, start=1):
        theme = _normalise_theme(v)
        if theme:
            ranks[i] = theme
        if i > 34:
            break
    _require_complete(ranks, path.name)
    reports.append({"name": path.stem.replace("_", " ").replace("-", " "), "date": "",
                     "ranks": ranks, "source_file": path.name})
    return reports


def _require_complete(ranks: dict, source: str):
    missing = [n for n in range(1, 35) if n not in ranks]
    if missing:
        raise ValueError(
            f"{source}: could not identify rank(s) {missing} out of 34 — "
            "layout may differ from what this parser expects. Send a sample so the parser can be adjusted."
        )


def leading_domain(ranks: dict) -> str:
    """Matches the Gallup report's own 'You lead with X themes' — domain of the #1 theme."""
    return DOMAIN_MAP[ranks[1]]


def domain_counts_top10(ranks: dict) -> dict:
    counts = {d: 0 for d in DOMAIN_ORDER}
    for n in range(1, 11):
        counts[DOMAIN_MAP[ranks[n]]] += 1
    return counts


# ---------------------------------------------------------------------------
# Domain "wheel" chart — one quadrant per domain, spokes = that domain's
# themes, spoke length = strength (rank 1 reaches the rim, rank 34 sits at
# the centre). Matches the style of wheel the team already builds by hand.
# ---------------------------------------------------------------------------
DOMAIN_CHART_COLORS = {
    "Executing": "#8064A2",              # purple
    "Influencing": "#F79646",            # orange
    "Relationship Building": "#4F81BD",  # blue
    "Strategic Thinking": "#9BBB59",     # green
}


def _wheel_layout() -> list:
    """34 (domain, theme, angle_in_radians) entries, each domain filling a fixed 90° quadrant,
    starting at 12 o'clock and going clockwise: Executing, Influencing, Relationship Building,
    Strategic Thinking."""
    layout = []
    for q, domain in enumerate(DOMAIN_ORDER):
        themes = DOMAIN_THEMES_ORDERED[domain]
        n = len(themes)
        for i, theme in enumerate(themes):
            angle_deg = q * 90 + (i + 0.5) / n * 90
            layout.append((domain, theme, np.radians(angle_deg)))
    return layout


_WHEEL_LAYOUT = _wheel_layout()


def generate_wheel_png(name: str, rank_of_theme: dict, out_path: Path):
    """rank_of_theme: {theme_name: rank_number (1-34)}. Saves a PNG to out_path."""
    fig = plt.figure(figsize=(3.6, 3.6), dpi=150)
    ax = fig.add_subplot(111, projection="polar")
    ax.set_theta_zero_location("N")
    ax.set_theta_direction(-1)
    ax.set_ylim(0, 34)
    ax.set_yticklabels([])
    ax.grid(color="#DDDDDD", linewidth=0.5)
    ax.spines["polar"].set_visible(False)

    by_domain = {}
    for domain, theme, angle in _WHEEL_LAYOUT:
        by_domain.setdefault(domain, []).append((theme, angle))

    for q, domain in enumerate(DOMAIN_ORDER):
        entries = by_domain[domain]
        quadrant_start = np.radians(q * 90)
        quadrant_end = np.radians((q + 1) * 90)
        # wedge: start at centre along the quadrant's lower boundary, out to each
        # spoke tip in turn, back to centre along the quadrant's upper boundary
        thetas = [quadrant_start] + [angle for _, angle in entries] + [quadrant_end]
        radii = [0.0] + [34 - rank_of_theme[theme] + 1 for theme, _ in entries] + [0.0]
        color = DOMAIN_CHART_COLORS[domain]
        ax.fill(thetas, radii, color=color, alpha=0.55, linewidth=0.8, edgecolor=color)

    tick_angles = [angle for _, _, angle in _WHEEL_LAYOUT]
    tick_labels = [theme for _, theme, _ in _WHEEL_LAYOUT]
    ax.set_xticks(tick_angles)
    ax.set_xticklabels(tick_labels, fontsize=4.3)
    ax.tick_params(axis="x", pad=2)

    fig.tight_layout(pad=0.3)
    fig.savefig(out_path, transparent=True)
    plt.close(fig)


def build_wheels_workbook(merged: dict, output_path: Path):
    """Builds a SEPARATE workbook containing just the 'Strengths Wheels' chart grid —
    2 people per row, each with a name banner above their wheel. A manual page break
    is inserted after every row of 2 charts so a wheel never gets sliced across a
    printed/exported page."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        wb = Workbook()
        ws = wb.active
        ws.title = "Strengths Wheels"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        banner_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        banner_font = Font(name="Arial", bold=True, size=12)

        CHART_COLS = 7       # columns each chart+banner spans
        CHART_ROW_SPAN = 20  # rows each chart occupies below its banner
        GAP_ROWS = 2         # blank rows between one chart-row and the next
        GAP_COLS = 1
        BLOCK_ROWS = 1 + CHART_ROW_SPAN + GAP_ROWS  # banner + chart + gap

        for col in range(1, 2 * CHART_COLS + GAP_COLS + 2):
            ws.column_dimensions[get_column_letter(col)].width = 9

        people = sorted(merged.items())
        last_grid_row = -1
        for i, (name, row) in enumerate(people):
            rank_of_theme = {theme: row[theme] for theme in ORDERED_THEMES}

            png_path = tmp_dir / f"wheel_{i}.png"
            generate_wheel_png(name, rank_of_theme, png_path)

            grid_row, grid_col = divmod(i, 2)
            top_row = grid_row * BLOCK_ROWS + 1
            left_col = grid_col * (CHART_COLS + GAP_COLS) + 1

            # start a fresh printed page for every new row of charts
            if grid_row != last_grid_row and grid_row > 0:
                ws.row_breaks.append(Break(id=top_row - 1))
            last_grid_row = grid_row

            ws.merge_cells(start_row=top_row, start_column=left_col,
                            end_row=top_row, end_column=left_col + CHART_COLS - 1)
            banner_cell = ws.cell(row=top_row, column=left_col, value=name)
            banner_cell.fill = banner_fill
            banner_cell.font = banner_font
            banner_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[top_row].height = 22

            img = XLImage(str(png_path))
            img.width = 330
            img.height = 330
            ws.add_image(img, f"{get_column_letter(left_col)}{top_row + 1}")

        wb.save(output_path)
    return len(people)


def collect_reports(input_path: Path) -> list:
    if input_path.is_dir():
        files = sorted(
            [p for p in input_path.iterdir()
             if p.suffix.lower() in (".pdf", ".txt", ".csv", ".xlsx", ".xls")]
        )
    else:
        files = [input_path]

    if not files:
        print(f"No supported files (.pdf/.txt/.csv/.xlsx) found at {input_path}", file=sys.stderr)
        sys.exit(1)

    reports = []
    for path in files:
        try:
            suffix = path.suffix.lower()
            if suffix == ".pdf":
                reports.append(extract_from_pdf(path))
            elif suffix == ".txt":
                reports.append(extract_from_text(path))
            elif suffix in (".csv", ".xlsx", ".xls"):
                reports.extend(extract_from_table(path))
            print(f"OK: {path.name}")
        except Exception as exc:  # noqa: BLE001 - surface per-file errors, keep going
            print(f"SKIPPED {path.name}: {exc}", file=sys.stderr)
    return reports


# ---------------------------------------------------------------------------
# Spreadsheet output
# ---------------------------------------------------------------------------
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
BODY_FONT = Font(name="Arial")
DOMAIN_FILL = {
    "Executing": PatternFill(start_color="E2D5F0", end_color="E2D5F0", fill_type="solid"),
    "Influencing": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
    "Relationship Building": PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"),
    "Strategic Thinking": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
}


DOMAIN_BANNER_FILL = {d: PatternFill(start_color=DOMAIN_CHART_COLORS[d].lstrip("#"),
                                      end_color=DOMAIN_CHART_COLORS[d].lstrip("#"),
                                      fill_type="solid")
                       for d in DOMAIN_ORDER}
INFO_COLUMNS = ["Name", "Report Date", "Source File", "Leading Domain"]


def build_workbook(reports: list, output_path: Path):
    """Builds the data workbook in the SAME layout as the team's own tracker:
    a domain-banner row, then a row of the 34 theme names as column headers,
    then one row per person with their RANK NUMBER (not the theme name) under
    each theme column."""
    existing_rows = {}
    if output_path.exists():
        wb_existing = load_workbook(output_path)
        if "CliftonStrengths Data" in wb_existing.sheetnames:
            ws_existing = wb_existing["CliftonStrengths Data"]
            headers = [c.value for c in ws_existing[2]]  # theme-name row is row 2
            for row in ws_existing.iter_rows(min_row=3, values_only=True):
                row_dict = dict(zip(headers, row))
                if row_dict.get("Name"):
                    existing_rows[row_dict["Name"]] = row_dict

    wb = Workbook()
    ws = wb.active
    ws.title = "CliftonStrengths Data"

    all_headers = INFO_COLUMNS + ORDERED_THEMES

    # Row 1: domain banner, merged across each domain's theme columns
    col = len(INFO_COLUMNS) + 1
    for domain in DOMAIN_ORDER:
        span = len(DOMAIN_THEMES_ORDERED[domain])
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + span - 1)
        cell = ws.cell(row=1, column=col, value=domain.upper())
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = DOMAIN_BANNER_FILL[domain]
        cell.alignment = Alignment(horizontal="center")
        col += span
    ws.row_dimensions[1].height = 20

    # Row 2: actual column headers (theme names)
    for col_idx, h in enumerate(all_headers, start=1):
        cell = ws.cell(row=2, column=col_idx, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "E3"

    merged = dict(existing_rows)
    for r in reports:
        row = {"Name": r["name"], "Report Date": r["date"], "Source File": r["source_file"],
               "Leading Domain": leading_domain(r["ranks"])}
        for rank_num, theme in r["ranks"].items():
            row[theme] = rank_num
        merged[r["name"]] = row

    for row_idx, (name, row) in enumerate(sorted(merged.items()), start=3):
        for col_idx, h in enumerate(all_headers, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=row.get(h, ""))
            cell.font = BODY_FONT
            cell.alignment = Alignment(horizontal="center") if h in DOMAIN_MAP else Alignment()
            domain = DOMAIN_MAP.get(h)
            if domain:
                cell.fill = DOMAIN_FILL[domain]

    for col_idx, h in enumerate(all_headers, start=1):
        width = 9 if h in DOMAIN_MAP else max(16, len(h) + 2)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws2 = wb.create_sheet("Domain Summary (Top 10)")
    summary_headers = ["Name"] + DOMAIN_ORDER + ["Leading Domain"]
    for col, h in enumerate(summary_headers, start=1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    ws2.freeze_panes = "A2"

    row_idx = 2
    for name, row in sorted(merged.items()):
        ranks = {row[theme]: theme for theme in ORDERED_THEMES}
        counts = domain_counts_top10(ranks)
        ws2.cell(row=row_idx, column=1, value=name).font = BODY_FONT
        for col_offset, d in enumerate(DOMAIN_ORDER, start=2):
            ws2.cell(row=row_idx, column=col_offset, value=counts[d]).font = BODY_FONT
        ws2.cell(row=row_idx, column=len(DOMAIN_ORDER) + 2, value=leading_domain(ranks)).font = BODY_FONT
        row_idx += 1
    for col_idx in range(1, len(summary_headers) + 1):
        ws2.column_dimensions[get_column_letter(col_idx)].width = 22

    wb.save(output_path)
    return merged


def main():
    parser = argparse.ArgumentParser(description="Extract CliftonStrengths 34 results (PDF/Excel/CSV/pasted text) into a spreadsheet + a separate wheel-chart workbook.")
    parser.add_argument("--input", required=True, help="A single file or a folder mixing .pdf/.txt/.csv/.xlsx files.")
    parser.add_argument("--output", required=True, help="Path to the output data .xlsx file.")
    parser.add_argument("--charts-output", default=None,
                         help="Path to the SEPARATE wheel-chart .xlsx file. Defaults to '<output>_Wheels.xlsx'.")
    parser.add_argument("--no-charts", action="store_true",
                         help="Skip generating the wheel-chart file (faster for large batches).")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)
    charts_output = Path(args.charts_output) if args.charts_output else \
        output_path.with_name(output_path.stem + "_Wheels" + output_path.suffix)

    reports = collect_reports(input_path)
    print(f"Parsed {len(reports)} report(s) successfully.")

    if not reports and not output_path.exists():
        print("Nothing parsed — no output file exists to merge into either.", file=sys.stderr)
        return

    merged = build_workbook(reports, output_path)
    print(f"Wrote {len(merged)} total row(s) to {output_path}")

    if not args.no_charts:
        chart_count = build_wheels_workbook(merged, charts_output)
        print(f"Wrote {chart_count} wheel chart(s) to {charts_output}")


if __name__ == "__main__":
    main()
