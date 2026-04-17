"""
core.py
=======
Shared processing logic used by both:
  - app.py          (Streamlit web app)
  - process_placements.py  (terminal script)

Do not run this file directly.
"""

import io
import re
import zipfile
from datetime import date, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Excel helper: row/col numbers -> range string e.g. "A1:S1" ──────────────

def merge_range(row1, col1, row2, col2) -> str:
    return f"{get_column_letter(col1)}{row1}:{get_column_letter(col2)}{row2}"


# ── School name cleaning ─────────────────────────────────────────────────────

def clean_school(name) -> str:
    if not isinstance(name, str):
        return name
    s = name.strip()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'\s*-\s*', ' - ', s)
    s = s.title()
    return s


# ── CSV loading ──────────────────────────────────────────────────────────────

def load_csv(source) -> pd.DataFrame:
    """
    source can be a file path (str) or a file-like object
    (e.g. an uploaded file from Streamlit).
    """
    raw = pd.read_csv(source)
    raw.columns = [
        "timestamp", "email", "full_name", "pref_name",
        "student_id", "school", "days_raw",
        *raw.columns[7:]
    ]
    raw["display_name"] = raw.apply(
        lambda r: str(r["pref_name"]).strip()
        if pd.notna(r["pref_name"])
           and str(r["pref_name"]).strip() not in ("", "nan")
        else str(r["full_name"]).strip(),
        axis=1,
    )
    raw["full_name"] = raw["full_name"].str.strip()
    raw["email"]     = raw["email"].str.strip()
    raw["school"]    = raw["school"].apply(clean_school)
    return raw


# ── Date detection ───────────────────────────────────────────────────────────

def extract_year(raw: pd.DataFrame) -> int:
    years = []
    for ts in raw["timestamp"].dropna():
        try:
            years.append(int(str(ts).split("/")[2].split(" ")[0]))
        except (IndexError, ValueError):
            pass
    if not years:
        raise ValueError(
            "Could not read a year from the Timestamp column. "
            "Check this is a Google Form CSV export."
        )
    return max(set(years), key=years.count)


def parse_form_day(raw_string: str, year: int):
    s = raw_string.strip()
    s = re.sub(r"(\d+)(st|nd|rd|th)\b", r"\1", s)
    s = re.sub(
        r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+",
        "", s,
    ).strip()
    s = f"{s} {year}"
    for fmt in ("%B %d %Y", "%d %B %Y"):
        try:
            p = pd.to_datetime(s, format=fmt)
            return date(p.year, p.month, p.day)
        except Exception:
            pass
    return None


def detect_placement_dates(raw: pd.DataFrame, year: int):
    all_tokens = set()
    for cell in raw["days_raw"].dropna():
        for token in str(cell).split(","):
            all_tokens.add(token.strip())

    sampled = set()
    for token in all_tokens:
        d = parse_form_day(token, year)
        if d and d.weekday() < 5:
            sampled.add(d)

    if not sampled:
        raise ValueError(
            "No valid placement dates found in the CSV. "
            "Check that the attendance days column is present."
        )

    ordered_dates = []
    cursor = min(sampled)
    while len(ordered_dates) < 15:
        if cursor.weekday() < 5:
            ordered_dates.append(cursor)
        cursor += timedelta(days=1)

    day_lookup = {}
    for token in all_tokens:
        d = parse_form_day(token, year)
        if d and d in ordered_dates:
            day_lookup[token] = d

    return ordered_dates, day_lookup


# ── Schedule labels ──────────────────────────────────────────────────────────

def build_schedule_labels(ordered_dates: list):
    short_labels = [
        f"{d.strftime('%a')} {d.day} {d.strftime('%b')}"
        for d in ordered_dates
    ]
    week_of = [i // 5 + 1 for i in range(15)]
    week_labels = []
    for w in range(1, 4):
        idx = [i for i, x in enumerate(week_of) if x == w]
        d1, d2 = ordered_dates[idx[0]], ordered_dates[idx[-1]]
        week_labels.append(
            f"Week {w}  ({d1.strftime('%b')} {d1.day}"
            f" - {d2.strftime('%b')} {d2.day})"
        )
    range_label = (
        f"{ordered_dates[0].strftime('%b')} {ordered_dates[0].day} - "
        f"{ordered_dates[14].strftime('%b')} {ordered_dates[14].day} "
        f"{ordered_dates[14].year}"
    )
    return short_labels, week_of, week_labels, range_label


# ── Attendance matrix for one school ─────────────────────────────────────────

def build_matrix(school: str, raw: pd.DataFrame, long_df: pd.DataFrame,
                 ordered_dates: list) -> pd.DataFrame:
    stu = (
        raw[raw["school"] == school]
        [["display_name", "full_name", "email", "student_id"]]
        .drop_duplicates()
        .sort_values("full_name")
        .reset_index(drop=True)
    )
    school_long = long_df[long_df["school"] == school]
    if not school_long.empty:
        pivot = (
            school_long[["display_name", "date"]]
            .drop_duplicates()
            .assign(present="✓")
            .pivot(index="display_name", columns="date", values="present")
        )
    else:
        pivot = pd.DataFrame(index=[], columns=ordered_dates)

    mat = stu.merge(pivot.reset_index(), on="display_name", how="left")
    for d in ordered_dates:
        if d not in mat.columns:
            mat[d] = ""
    mat[ordered_dates] = mat[ordered_dates].fillna("")
    return mat


# ── Expand CSV to long format ────────────────────────────────────────────────

def expand_long(raw: pd.DataFrame, day_lookup: dict) -> pd.DataFrame:
    rows = []
    for _, s in raw.iterrows():
        if pd.isna(s["days_raw"]):
            continue
        for token in str(s["days_raw"]).split(","):
            token = token.strip()
            if token in day_lookup:
                rows.append({
                    "school":       s["school"],
                    "display_name": s["display_name"],
                    "full_name":    s["full_name"],
                    "email":        s["email"],
                    "student_id":   s["student_id"],
                    "date":         day_lookup[token],
                })
    return pd.DataFrame(rows).drop_duplicates()


# ── Excel colour palette ─────────────────────────────────────────────────────

NAVY      = "0D47A1"
DARK_NAVY = "1F3864"
MID_NAVY  = "1565C0"
DARK_GREY = "2C3E50"
TICK_BG   = "00897B"
ALT_ROW   = "F5F5F5"

WK_BG  = ["FFF9C4", "FCE4EC", "E8EAF6"]
WK_HDR = ["F9A825", "AD1457", "5C6BC0"]

COUNT_GREEN = ("2ECC71", "27AE60")
COUNT_AMBER = ("F39C12", "E67E22")
COUNT_RED   = ("E74C3C", "C0392B")


def make_border(colour="CCCCCC"):
    s = Side(style="thin", color=colour)
    return Border(left=s, right=s, top=s, bottom=s)

def solid_fill(hex_colour):
    return PatternFill("solid", fgColor=hex_colour)

def style_cell(cell, bg=None, fg="000000", bold=False, size=10,
               h_align="center", border_colour="CCCCCC",
               wrap=False, italic=False):
    if bg:
        cell.fill = solid_fill(bg)
    cell.font      = Font(name="Arial", bold=bold, italic=italic,
                          color=fg, size=size)
    cell.alignment = Alignment(horizontal=h_align, vertical="center",
                               wrap_text=wrap)
    cell.border    = make_border(border_colour)


# ── Build one Excel workbook ─────────────────────────────────────────────────

def build_workbook(school, mat, ordered_dates, short_labels,
                   week_of, week_labels, range_label) -> Workbook:

    n_students    = len(mat)
    fixed_headers = ["Preferred Name", "Full Name", "Email Address", "Student ID"]
    n_fixed       = 4
    n_days        = 15
    total_cols    = n_fixed + n_days

    wb = Workbook()

    # ── Sheet 1: Attendance Schedule ─────────────────────────
    ws = wb.active
    ws.title = "Attendance Schedule"
    ws.sheet_view.showGridLines = False

    ws.merge_cells(merge_range(1, 1, 1, total_cols))
    c = ws.cell(1, 1, school)
    style_cell(c, bg=NAVY, fg="FFFFFF", bold=True, size=14, border_colour="FFFFFF")
    ws.row_dimensions[1].height = 32

    ws.merge_cells(merge_range(2, 1, 2, total_cols))
    c = ws.cell(2, 1,
        f"Placement Attendance  |  {n_students} student(s)  |  "
        f"3-Week Period: {range_label}  |  10 days per student")
    style_cell(c, bg=MID_NAVY, fg="E3F2FD", size=11, wrap=True, border_colour="FFFFFF")
    ws.row_dimensions[2].height = 22

    for col in range(1, n_fixed + 1):
        style_cell(ws.cell(3, col), bg=DARK_NAVY, border_colour="FFFFFF")
    for w in range(3):
        cols = [i for i, x in enumerate(week_of) if x == w + 1]
        cs, ce = n_fixed + cols[0] + 1, n_fixed + cols[-1] + 1
        ws.merge_cells(merge_range(3, cs, 3, ce))
        c = ws.cell(3, cs, week_labels[w])
        style_cell(c, bg=WK_HDR[w], fg="FFFFFF", bold=True, border_colour="FFFFFF")
    ws.row_dimensions[3].height = 22

    for ci, h in enumerate(fixed_headers + short_labels, 1):
        c = ws.cell(4, ci, h)
        style_cell(c, bg=DARK_NAVY, fg="FFFFFF", bold=True,
                   size=11, border_colour="FFFFFF", wrap=True)
    ws.row_dimensions[4].height = 36

    data_start = 5
    for r_idx, row in mat.iterrows():
        rn     = data_start + r_idx
        row_bg = ALT_ROW if r_idx % 2 == 1 else None
        for ci, val in enumerate(
            [row["display_name"], row["full_name"], row["email"], row["student_id"]], 1
        ):
            c = ws.cell(rn, ci, val)
            style_cell(c, bg=row_bg, h_align="left" if ci <= 3 else "center")
        for di, d in enumerate(ordered_dates):
            ci  = n_fixed + di + 1
            val = row[d]
            c   = ws.cell(rn, ci, val)
            if val == "✓":
                style_cell(c, bg=TICK_BG, fg="FFFFFF", bold=True, size=11)
            else:
                style_cell(c, bg=WK_BG[week_of[di] - 1])
        ws.row_dimensions[rn].height = 18

    tr = data_start + n_students
    ws.merge_cells(merge_range(tr, 1, tr, n_fixed))
    c = ws.cell(tr, 1, "TOTAL STUDENTS PER DAY")
    style_cell(c, bg=DARK_GREY, fg="FFFFFF", bold=True,
               h_align="left", border_colour="FFFFFF")
    for di, d in enumerate(ordered_dates):
        n_att = (mat[d] == "✓").sum()
        c = ws.cell(tr, n_fixed + di + 1, n_att)
        style_cell(c, bg=DARK_NAVY, fg="FFFFFF", bold=True, border_colour="FFFFFF")
    ws.row_dimensions[tr].height = 22

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 13
    for di in range(n_days):
        ws.column_dimensions[get_column_letter(n_fixed + di + 1)].width = 10
    ws.freeze_panes = ws.cell(data_start, n_fixed + 1)

    # ── Sheet 2: Daily Summary ───────────────────────────────
    ws2 = wb.create_sheet("Daily Summary")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:D1")
    c = ws2.cell(1, 1, f"{school} -- Daily Student Count")
    style_cell(c, bg=NAVY, fg="FFFFFF", bold=True, size=13, border_colour="FFFFFF")
    ws2.row_dimensions[1].height = 28

    for ci, h in enumerate(
        ["Week", "Date", "Day of Week", "Students Attending"], 1
    ):
        c = ws2.cell(2, ci, h)
        style_cell(c, bg=DARK_NAVY, fg="FFFFFF", bold=True,
                   size=11, border_colour="FFFFFF")
    ws2.row_dimensions[2].height = 22

    for di, d in enumerate(ordered_dates):
        r   = 3 + di
        w   = week_of[di] - 1
        n   = (mat[d] == "✓").sum()
        ws2.cell(r, 1, week_labels[w])
        ws2.cell(r, 2, short_labels[di])
        ws2.cell(r, 3, d.strftime("%A"))
        ws2.cell(r, 4, n)
        for ci in range(1, 4):
            style_cell(ws2.cell(r, ci), bg=WK_BG[w], h_align="left")
        bg_c, bdr_c = (COUNT_GREEN if n >= 5
                       else COUNT_AMBER if n >= 2
                       else COUNT_RED)
        style_cell(ws2.cell(r, 4), bg=bg_c, fg="FFFFFF",
                   bold=True, border_colour=bdr_c)
        ws2.row_dimensions[r].height = 20

    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 14
    ws2.column_dimensions["C"].width = 16
    ws2.column_dimensions["D"].width = 22

    note = 3 + n_days + 1
    ws2.merge_cells(f"A{note}:D{note}")
    c = ws2.cell(note, 1,
                 "Colour key:  Green = 5+ students   "
                 "Amber = 2-4 students   Red = 0-1 students")
    c.font      = Font(name="Arial", size=9, color="555555", italic=True)
    c.alignment = Alignment(horizontal="left")

    return wb


# ── Save workbook to bytes (for in-memory download) ──────────────────────────

def workbook_to_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Build a ZIP of multiple workbooks in memory ──────────────────────────────

def build_zip(schools_data: list) -> bytes:
    """
    schools_data: list of (filename, wb_bytes)
    Returns a ZIP file as bytes.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, wb_bytes in schools_data:
            zf.writestr(filename, wb_bytes)
    return buf.getvalue()
