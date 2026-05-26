"""
core.py
=======
Shared processing logic for the Placement Schedule Generator.
Used by both app.py (Streamlit) and process_placements.py (terminal).
"""

import io
import re
import zipfile
from datetime import date, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Excel helper ─────────────────────────────────────────────

def merge_range(row1, col1, row2, col2) -> str:
    return f"{get_column_letter(col1)}{row1}:{get_column_letter(col2)}{row2}"


# ── File loading ─────────────────────────────────────────────

def read_file(source) -> pd.DataFrame:
    """Read a CSV or Excel file from a path or file-like object."""
    name = getattr(source, "name", str(source))
    if str(name).lower().endswith(".csv"):
        return pd.read_csv(source)
    return pd.read_excel(source)


# ── InPlace loading and filtering ────────────────────────────

def load_inplace(source) -> pd.DataFrame:
    """
    Load the InPlace export. Expects these columns (position-independent):
      Student, Email, Student Code, Agency,
      Start Date, Status, Placement Allocation Groups
    """
    df = read_file(source)

    required = ["Student", "Email", "Student Code", "Agency",
                "Start Date", "Status", "Placement Allocation Groups"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"InPlace file is missing these columns: {missing}\n"
            f"Found columns: {list(df.columns)}"
        )

    df["Email"]         = df["Email"].astype(str).str.strip().str.lower()
    df["Agency"]        = df["Agency"].astype(str).str.strip()
    df["Student"]       = df["Student"].astype(str).str.strip()
    df["Student Code"]  = df["Student Code"].astype(str).str.strip()
    df["Start Date"]    = pd.to_datetime(df["Start Date"], errors="coerce")
    df["Status"]        = df["Status"].astype(str).str.strip()

    return df


def get_available_years(inplace_df: pd.DataFrame) -> list:
    """Extract unique years from the Start Date column, sorted descending."""
    years = (
        inplace_df["Start Date"]
        .dropna()
        .dt.year
        .unique()
    )
    return sorted([int(y) for y in years], reverse=True)


def filter_inplace(inplace_df: pd.DataFrame, year: int):
    """
    Apply all filters to the InPlace data:

      Year filter (two-pronged — handles unplaced students with blank Start Date):
        Keep row if EITHER:
          a) Start Date year == selected year, OR
          b) Start Date is blank/null
        AND Requirement Groups contains the selected year as a 4-digit number.
        This ensures unplaced students (no Start Date yet) are included as long
        as their Requirement Group is for the correct year.

      Then:
        Exclude agencies starting with 'EDU -'
        Exclude Status == 'Completed'

    Returns:
      filtered_df   — the filtered InPlace DataFrame
      window_start  — earliest Start Date in this year (date), or None
      window_end    — latest End Date in this year (date), or None
    """
    import re as _re
    df = inplace_df.copy()

    # Year in Requirement Groups (catches all rows including unplaced)
    req_has_year = df["Requirement Groups"].astype(str).str.contains(
        str(year), na=False
    )
    # Also accept Placement Allocation Groups as a fallback
    alloc_has_year = df["Placement Allocation Groups"].astype(str).str.contains(
        str(year), na=False
    )
    group_year_ok = req_has_year | alloc_has_year

    # Start Date: matches year OR is blank
    start_year_ok = (
        df["Start Date"].dt.year == year
    ) | df["Start Date"].isna()

    # Both conditions must be true
    df = df[group_year_ok & start_year_ok]

    # Capture the placement window from rows that DO have a Start Date
    placed = df[df["Start Date"].notna()]
    window_start = placed["Start Date"].min().date() if len(placed) > 0 else None
    end_col = pd.to_datetime(placed["End Date"], errors="coerce") if len(placed) > 0 else pd.Series([], dtype="datetime64[ns]")
    window_end = end_col.max().date() if len(placed) > 0 and pd.notna(end_col.max()) else None

    # Exclude EDU - agencies
    df = df[~df["Agency"].str.startswith("EDU -", na=False)]

    # Exclude Completed status (blank is allowed)
    df = df[df["Status"].str.lower() != "completed"]

    return df.reset_index(drop=True), window_start, window_end


def clean_school_name(name) -> str:
    """Normalise school name — same logic as before."""
    if not isinstance(name, str) or name.strip() == "":
        return name
    s = name.strip()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'\s*-\s*', ' - ', s)
    s = s.title()
    return s


# ── Google Form loading ──────────────────────────────────────

def load_form(source) -> pd.DataFrame:
    """
    Load the Google Form CSV/Excel export.
    Uses column positions 0-7 (same structure every year).
    """
    df = read_file(source)

    # Rename by position — robust to minor column name changes
    col_map = {
        df.columns[0]: "timestamp",
        df.columns[1]: "email",
        df.columns[2]: "full_name",
        df.columns[3]: "pref_name",
        df.columns[4]: "student_id",
        df.columns[5]: "school_form",   # ignored — we use InPlace Agency
        df.columns[6]: "days_raw",
    }
    df = df.rename(columns=col_map)

    df["email"]      = df["email"].astype(str).str.strip().str.lower()
    df["pref_name"]  = df["pref_name"].astype(str).str.strip()
    df["pref_name"]  = df["pref_name"].where(
        ~df["pref_name"].isin(["", "nan", "NaN"]), other=""
    )
    df["days_raw"]   = df["days_raw"].astype(str).str.strip()
    df["days_raw"]   = df["days_raw"].where(df["days_raw"] != "nan", other="")

    return df


# ── Date detection ───────────────────────────────────────────

def extract_year_from_form(form_df: pd.DataFrame) -> int:
    years = []
    for ts in form_df["timestamp"].dropna():
        try:
            years.append(int(str(ts).split("/")[2].split(" ")[0]))
        except (IndexError, ValueError):
            pass
    if not years:
        raise ValueError("Could not detect year from the Form timestamp column.")
    return max(set(years), key=years.count)


def parse_form_day(raw_string: str, year: int):
    s = raw_string.strip()
    s = re.sub(r"(\d+)(st|nd|rd|th)\b", r"\1", s)
    s = re.sub(
        r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+",
        "", s
    ).strip()
    s = f"{s} {year}"
    for fmt in ("%B %d %Y", "%d %B %Y"):
        try:
            p = pd.to_datetime(s, format=fmt)
            return date(p.year, p.month, p.day)
        except Exception:
            pass
    return None


def detect_placement_dates(form_df: pd.DataFrame, year: int):
    """
    Parse placement dates from the form using the timestamp year as the
    placement year. Students always submit one month before a mid-year
    placement, so the timestamp year always matches the placement year.
    This also acts as a natural per-year filter — only tokens that parse
    into a weekday in the correct year are included.
    """
    # Filter form rows to only those submitted in the selected year
    year_mask = pd.to_datetime(form_df["timestamp"], dayfirst=True, errors="coerce").dt.year == year
    form_year = form_df[year_mask]

    all_tokens = set()
    for cell in form_year["days_raw"]:
        if cell:
            for token in cell.split(","):
                all_tokens.add(token.strip())

    sampled = set()
    for token in all_tokens:
        d = parse_form_day(token, year)
        if d and d.weekday() < 5:
            sampled.add(d)

    if not sampled:
        raise ValueError(
            f"No valid placement dates found in the Form for {year}. "
            f"Check that the form responses include submissions from {year}."
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


# ── Master join ──────────────────────────────────────────────

def build_master(inplace_df: pd.DataFrame,
                 form_df: pd.DataFrame,
                 day_lookup: dict) -> tuple:
    """
    Join InPlace (filtered) with Form responses on email.

    Only form rows where at least one day falls within the day_lookup
    (i.e. within the selected year's placement window) are treated as
    valid submissions. This handles students who submitted the form in
    multiple years — only the relevant year's submission is used.

    Returns:
        master_df       — one row per student, with all fields + date columns
        no_placement_df — form submissions with no InPlace record
    """
    # Build form lookup: email -> {pref_name, days set}
    # Since form is already filtered by timestamp year in detect_placement_dates,
    # only the current year's submissions reach here.
    # If a student submitted multiple times in the same year, keep the one
    # with the most days in the placement window.
    form_lookup = {}
    for _, row in form_df.iterrows():
        email = row["email"]
        days = set()
        if row["days_raw"]:
            for token in row["days_raw"].split(","):
                token = token.strip()
                if token in day_lookup:
                    days.add(day_lookup[token])

        if days:
            existing = form_lookup.get(email, {})
            if len(days) >= len(existing.get("days", set())):
                form_lookup[email] = {
                    "pref_name": row["pref_name"],
                    "days":      days,
                    "submitted": True,
                }

    # Students in form but not in InPlace
    inplace_emails = set(inplace_df["Email"].str.lower().str.strip())
    no_placement_rows = []
    for _, row in form_df.iterrows():
        if row["email"] not in inplace_emails:
            no_placement_rows.append({
                "student_name": row["full_name"],
                "pref_name":    row["pref_name"],
                "email":        row["email"],
                "student_id":   row["student_id"],
                "agency":       "— No placement assigned —",
                "submitted":    True,
                "days":         form_lookup.get(row["email"], {}).get("days", set()),
            })
    no_placement_df = pd.DataFrame(no_placement_rows)

    # Build master from InPlace
    master_rows = []
    for _, ip_row in inplace_df.iterrows():
        email      = ip_row["Email"].lower().strip()
        form_data  = form_lookup.get(email, {"pref_name": "", "days": set(), "submitted": False})
        master_rows.append({
            "student_name": ip_row["Student"],
            "pref_name":    form_data["pref_name"],   # from form only
            "email":        email,
            "student_id":   ip_row["Student Code"],
            "agency":       clean_school_name(ip_row["Agency"]),
            "submitted":    form_data["submitted"],
            "days":         form_data["days"],
        })

    master_df = pd.DataFrame(master_rows)
    return master_df, no_placement_df


# ── Build attendance matrix for one school ───────────────────

def build_matrix_from_master(school: str,
                              master_df: pd.DataFrame,
                              ordered_dates: list) -> pd.DataFrame:
    """
    Returns a DataFrame for one school:
    - Submitted students first (sorted by name), with tick marks
    - Non-submitted students at the bottom, flagged
    """
    school_df = master_df[master_df["agency"] == school].copy()

    submitted     = school_df[school_df["submitted"]].sort_values("student_name")
    not_submitted = school_df[~school_df["submitted"]].sort_values("student_name")

    rows = []

    for _, s in submitted.iterrows():
        row = {
            "display_name": s["pref_name"] if s["pref_name"] else s["student_name"],
            "full_name":    s["student_name"],
            "email":        s["email"],
            "student_id":   s["student_id"],
            "note":         "",
        }
        for d in ordered_dates:
            row[d] = "✓" if d in s["days"] else ""
        rows.append(row)

    for _, s in not_submitted.iterrows():
        row = {
            "display_name": s["student_name"],
            "full_name":    s["student_name"],
            "email":        s["email"],
            "student_id":   s["student_id"],
            "note":         "No preference submitted",
        }
        for d in ordered_dates:
            row[d] = ""
        rows.append(row)

    return pd.DataFrame(rows).reset_index(drop=True)


# ── Colour palette ───────────────────────────────────────────

NAVY      = "0D47A1"
DARK_NAVY = "1F3864"
MID_NAVY  = "1565C0"
DARK_GREY = "2C3E50"
TICK_BG   = "00897B"
ALT_ROW   = "F5F5F5"
NO_SUB_BG = "FFF3CD"   # amber tint for "no preference submitted" rows
NO_SUB_FG = "856404"

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


# ── Build one Excel workbook ─────────────────────────────────

def build_workbook(school, mat, ordered_dates, short_labels,
                   week_of, week_labels, range_label) -> Workbook:

    n_students    = len(mat)
    # Columns: Preferred Name | Full Name | Email | Student ID | Note | [days]
    fixed_headers = ["Preferred Name", "Full Name", "Email Address",
                     "Student ID", "Note"]
    n_fixed       = len(fixed_headers)   # 5
    n_days        = 15
    total_cols    = n_fixed + n_days

    wb = Workbook()

    # ── Sheet 1: Attendance Schedule ────────────────────────
    ws = wb.active
    ws.title = "Attendance Schedule"
    ws.sheet_view.showGridLines = False

    # Row 1: school title
    ws.merge_cells(merge_range(1, 1, 1, total_cols))
    c = ws.cell(1, 1, school)
    style_cell(c, bg=NAVY, fg="FFFFFF", bold=True, size=14,
               border_colour="FFFFFF")
    ws.row_dimensions[1].height = 32

    # Row 2: subtitle
    ws.merge_cells(merge_range(2, 1, 2, total_cols))
    c = ws.cell(2, 1,
        f"Placement Attendance  |  {n_students} student(s)  |  "
        f"3-Week Period: {range_label}  |  10 days per student")
    style_cell(c, bg=MID_NAVY, fg="E3F2FD", size=11,
               wrap=True, border_colour="FFFFFF")
    ws.row_dimensions[2].height = 22

    # Row 3: week group headers
    for col in range(1, n_fixed + 1):
        style_cell(ws.cell(3, col), bg=DARK_NAVY, border_colour="FFFFFF")
    for w in range(3):
        cols = [i for i, x in enumerate(week_of) if x == w + 1]
        cs = n_fixed + cols[0] + 1
        ce = n_fixed + cols[-1] + 1
        ws.merge_cells(merge_range(3, cs, 3, ce))
        c = ws.cell(3, cs, week_labels[w])
        style_cell(c, bg=WK_HDR[w], fg="FFFFFF", bold=True,
                   border_colour="FFFFFF")
    ws.row_dimensions[3].height = 22

    # Row 4: column headers
    for ci, h in enumerate(fixed_headers + short_labels, 1):
        c = ws.cell(4, ci, h)
        style_cell(c, bg=DARK_NAVY, fg="FFFFFF", bold=True,
                   size=11, border_colour="FFFFFF", wrap=True)
    ws.row_dimensions[4].height = 36

    # Rows 5+: student data
    data_start = 5
    for r_idx, row in mat.iterrows():
        rn          = data_start + r_idx
        no_sub      = row["note"] == "No preference submitted"
        row_bg      = NO_SUB_BG if no_sub else (ALT_ROW if r_idx % 2 == 1 else None)
        text_colour = NO_SUB_FG if no_sub else "000000"

        for ci, val in enumerate(
            [row["display_name"], row["full_name"],
             row["email"], row["student_id"], row["note"]], 1
        ):
            c = ws.cell(rn, ci, val)
            style_cell(c, bg=row_bg, fg=text_colour,
                        h_align="left" if ci <= 4 else "center",
                        italic=no_sub and ci == 5)

        for di, d in enumerate(ordered_dates):
            ci  = n_fixed + di + 1
            val = row[d]
            c   = ws.cell(rn, ci, val)
            if val == "✓":
                style_cell(c, bg=TICK_BG, fg="FFFFFF", bold=True, size=11)
            else:
                style_cell(c, bg=NO_SUB_BG if no_sub else WK_BG[week_of[di] - 1])
        ws.row_dimensions[rn].height = 18

    # Total row (only counts submitted students)
    tr = data_start + n_students
    ws.merge_cells(merge_range(tr, 1, tr, n_fixed))
    c = ws.cell(tr, 1, "TOTAL STUDENTS PER DAY")
    style_cell(c, bg=DARK_GREY, fg="FFFFFF", bold=True,
               h_align="left", border_colour="FFFFFF")
    for di, d in enumerate(ordered_dates):
        n_att = (mat[d] == "✓").sum()
        c = ws.cell(tr, n_fixed + di + 1, n_att)
        style_cell(c, bg=DARK_NAVY, fg="FFFFFF", bold=True,
                   border_colour="FFFFFF")
    ws.row_dimensions[tr].height = 22

    # Column widths and freeze
    ws.column_dimensions["A"].width = 22   # preferred name
    ws.column_dimensions["B"].width = 22   # full name
    ws.column_dimensions["C"].width = 30   # email
    ws.column_dimensions["D"].width = 13   # student id
    ws.column_dimensions["E"].width = 26   # note
    for di in range(n_days):
        ws.column_dimensions[get_column_letter(n_fixed + di + 1)].width = 10
    ws.freeze_panes = ws.cell(data_start, n_fixed + 1)

    # ── Sheet 2: Daily Summary ───────────────────────────────
    ws2 = wb.create_sheet("Daily Summary")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:D1")
    c = ws2.cell(1, 1, f"{school} -- Daily Student Count")
    style_cell(c, bg=NAVY, fg="FFFFFF", bold=True, size=13,
               border_colour="FFFFFF")
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

    note_row = 3 + n_days + 1
    ws2.merge_cells(f"A{note_row}:D{note_row}")
    c = ws2.cell(note_row, 1,
                 "Colour key:  Green = 5+ students   "
                 "Amber = 2-4 students   Red = 0-1 students")
    c.font      = Font(name="Arial", size=9, color="555555", italic=True)
    c.alignment = Alignment(horizontal="left")

    return wb


# ── Build "No Placement Assigned" workbook ───────────────────

def build_no_placement_workbook(no_placement_df: pd.DataFrame,
                                 ordered_dates: list,
                                 short_labels: list,
                                 week_of: list,
                                 week_labels: list,
                                 range_label: str) -> Workbook:
    """One workbook for students who submitted the form but have no InPlace record."""

    n_fixed    = 4   # Name | Email | Student ID | Days submitted
    n_days     = 15
    total_cols = n_fixed + n_days

    wb = Workbook()
    ws = wb.active
    ws.title = "No Placement Assigned"
    ws.sheet_view.showGridLines = False

    ws.merge_cells(merge_range(1, 1, 1, total_cols))
    c = ws.cell(1, 1, "Students Without a Placement Assignment")
    style_cell(c, bg=NAVY, fg="FFFFFF", bold=True, size=14,
               border_colour="FFFFFF")
    ws.row_dimensions[1].height = 32

    ws.merge_cells(merge_range(2, 1, 2, total_cols))
    c = ws.cell(2, 1,
        f"These students submitted the Google Form but have no agency assigned in InPlace  |  "
        f"Period: {range_label}")
    style_cell(c, bg=MID_NAVY, fg="E3F2FD", size=11,
               wrap=True, border_colour="FFFFFF")
    ws.row_dimensions[2].height = 22

    # Week headers
    for col in range(1, n_fixed + 1):
        style_cell(ws.cell(3, col), bg=DARK_NAVY, border_colour="FFFFFF")
    for w in range(3):
        cols = [i for i, x in enumerate(week_of) if x == w + 1]
        cs = n_fixed + cols[0] + 1
        ce = n_fixed + cols[-1] + 1
        ws.merge_cells(merge_range(3, cs, 3, ce))
        c = ws.cell(3, cs, week_labels[w])
        style_cell(c, bg=WK_HDR[w], fg="FFFFFF", bold=True,
                   border_colour="FFFFFF")
    ws.row_dimensions[3].height = 22

    # Column headers
    headers = ["Full Name", "Email Address", "Student ID",
               "Preferred Name"] + short_labels
    for ci, h in enumerate(headers, 1):
        c = ws.cell(4, ci, h)
        style_cell(c, bg=DARK_NAVY, fg="FFFFFF", bold=True,
                   size=11, border_colour="FFFFFF", wrap=True)
    ws.row_dimensions[4].height = 36

    data_start = 5
    for r_idx, row in no_placement_df.reset_index(drop=True).iterrows():
        rn     = data_start + r_idx
        row_bg = ALT_ROW if r_idx % 2 == 1 else None

        for ci, val in enumerate(
            [row.get("student_name", ""), row.get("email", ""),
             row.get("student_id", ""), row.get("pref_name", "")], 1
        ):
            c = ws.cell(rn, ci, val)
            style_cell(c, bg=row_bg, h_align="left")

        for di, d in enumerate(ordered_dates):
            ci  = n_fixed + di + 1
            val = "✓" if d in row.get("days", set()) else ""
            c   = ws.cell(rn, ci, val)
            if val == "✓":
                style_cell(c, bg=TICK_BG, fg="FFFFFF", bold=True, size=11)
            else:
                style_cell(c, bg=WK_BG[week_of[di] - 1])
        ws.row_dimensions[rn].height = 18

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 22
    for di in range(n_days):
        ws.column_dimensions[get_column_letter(n_fixed + di + 1)].width = 10
    ws.freeze_panes = ws.cell(data_start, n_fixed + 1)

    return wb


# ── Save to bytes ────────────────────────────────────────────

def workbook_to_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_zip(schools_data: list) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, wb_bytes in schools_data:
            zf.writestr(filename, wb_bytes)
    return buf.getvalue()
