"""
app.py  —  Placement Schedule Generator (Streamlit web app)

HOW TO RUN LOCALLY:
  streamlit run app.py

HOW TO DEPLOY (free):
  Push to GitHub, then connect at https://streamlit.io/cloud
"""

import re
import streamlit as st
from core import (
    load_csv,
    extract_year,
    detect_placement_dates,
    build_schedule_labels,
    expand_long,
    build_matrix,
    build_workbook,
    workbook_to_bytes,
    build_zip,
)

# ── Page config ──────────────────────────────────────────────
st.set_page_config(
    page_title="Placement Schedule Generator",
    page_icon="🎓",
    layout="centered",
)

# ── Monash brand colours ─────────────────────────────────────
MONASH_BLUE       = "#006DAE"
MONASH_BLUE_DARK  = "#005490"
MONASH_BLUE_LIGHT = "#E6F2F9"
MONASH_GREEN      = "#008A25"
MONASH_DARK       = "#3C3C3C"
MONASH_GREY       = "#5A5A5A"
MONASH_LIGHT_GREY = "#E6E6E6"
MONASH_BG         = "#F6F6F6"

# ── Custom CSS ───────────────────────────────────────────────
st.markdown(f"""
<style>
html, body, [class*="css"] {{
    font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
}}
#MainMenu, footer, header {{ visibility: hidden; }}
.block-container {{
    padding-top: 0 !important;
    padding-bottom: 3rem;
    max-width: 800px;
}}

/* ── Monash header bar ── */
.monash-header {{
    background: {MONASH_BLUE};
    padding: 18px 28px;
    margin: -1rem -1rem 0 -1rem;
    display: flex;
    align-items: center;
    gap: 18px;
}}
.monash-logo-text {{
    color: white;
    font-size: 1.7rem;
    font-weight: 900;
    letter-spacing: -1px;
    border: 2.5px solid white;
    padding: 4px 10px;
    border-radius: 4px;
    white-space: nowrap;
    line-height: 1;
}}
.monash-header-text h1 {{
    color: white;
    font-size: 1.3rem;
    font-weight: 700;
    margin: 0 0 2px 0;
    letter-spacing: -0.3px;
}}
.monash-header-text p {{
    color: rgba(255,255,255,0.78);
    font-size: 0.82rem;
    margin: 0;
}}

/* ── Step section card ── */
/* We target Streamlit's container element directly */
div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"] {{
    background: white !important;
    border: 1px solid {MONASH_LIGHT_GREY} !important;
    border-top: 3px solid {MONASH_BLUE} !important;
    border-radius: 8px !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06) !important;
    padding: 4px 8px 24px !important;
}}

/* ── Step header inside card ── */
.step-header {{
    display: flex;
    align-items: center;
    gap: 12px;
    padding-bottom: 14px;
    margin-bottom: 6px;
    border-bottom: 1px solid {MONASH_LIGHT_GREY};
}}
.step-badge {{
    background: {MONASH_BLUE};
    color: white;
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 0.8px;
    padding: 4px 12px;
    border-radius: 3px;
    text-transform: uppercase;
    flex-shrink: 0;
}}
.step-title {{
    font-size: 1rem;
    font-weight: 600;
    color: {MONASH_DARK};
    margin: 0;
}}

/* ── Metric cards ── */
.metrics-row {{
    display: flex;
    gap: 12px;
    margin-top: 14px;
    margin-bottom: 12px;
}}
.metric-card {{
    flex: 1;
    background: {MONASH_BG};
    border-radius: 6px;
    padding: 14px 16px;
    border-left: 3px solid {MONASH_BLUE};
}}
.metric-label {{
    font-size: 0.67rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.7px;
    color: {MONASH_GREY};
    margin-bottom: 5px;
}}
.metric-value {{
    font-size: 1.8rem;
    font-weight: 700;
    color: {MONASH_BLUE};
    line-height: 1;
}}
.metric-sub {{
    font-size: 0.71rem;
    color: #999;
    margin-top: 3px;
}}

/* ── Period box ── */
.period-box {{
    background: {MONASH_BLUE_LIGHT};
    border: 1px solid #B3D4EA;
    border-radius: 6px;
    padding: 10px 14px;
    font-size: 0.85rem;
    color: {MONASH_BLUE_DARK};
    font-weight: 500;
    margin-top: 4px;
    margin-bottom: 8px;
}}
.period-box span {{ font-weight: 700; }}

/* ── Status pills ── */
.pill-success {{
    display: inline-flex;
    align-items: center;
    gap: 7px;
    background: #E8F5EE;
    color: #1B5E35;
    border: 1px solid #A8D5BB;
    border-radius: 4px;
    padding: 6px 14px;
    font-size: 0.83rem;
    font-weight: 600;
    margin-top: 6px;
    margin-bottom: 8px;
}}
.pill-dot {{
    width: 7px;
    height: 7px;
    border-radius: 50%;
    background: {MONASH_GREEN};
    flex-shrink: 0;
}}

/* ── Progress label ── */
.progress-label {{
    font-size: 0.8rem;
    color: {MONASH_GREY};
    font-weight: 500;
    background: white;
    padding: 6px 10px;
    border-radius: 4px;
    border: 1px solid {MONASH_LIGHT_GREY};
    margin-top: 6px;
    margin-bottom: 8px;
}}

/* ── Success box ── */
.success-box {{
    background: #E8F5EE;
    border: 1px solid #A8D5BB;
    border-left: 4px solid {MONASH_GREEN};
    border-radius: 6px;
    padding: 12px 16px;
    margin: 8px 0 8px;
    font-size: 0.9rem;
    font-weight: 600;
    color: #1B5E35;
}}

/* ── Buttons ── */
div[data-testid="stButton"] > button[kind="primary"] {{
    background: {MONASH_BLUE} !important;
    border: none !important;
    border-radius: 6px !important;
    font-size: 0.9rem !important;
    font-weight: 600 !important;
}}
div[data-testid="stButton"] > button[kind="primary"]:hover {{
    background: {MONASH_BLUE_DARK} !important;
}}
div[data-testid="stDownloadButton"] > button {{
    background: {MONASH_GREEN} !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-size: 0.9rem !important;
    font-weight: 600 !important;
}}
div[data-testid="stDownloadButton"] > button:hover {{
    background: #006B1D !important;
}}

/* ── Upload widget ── */
[data-testid="stFileUploader"] {{
    border: 2px dashed {MONASH_LIGHT_GREY} !important;
    border-radius: 8px !important;
    background: {MONASH_BG} !important;
}}

/* ── Progress bar ── */
[data-testid="stProgress"] > div > div {{
    background: {MONASH_BLUE} !important;
    border-radius: 3px !important;
}}
[data-testid="stProgress"] > div {{
    background: {MONASH_LIGHT_GREY} !important;
    border-radius: 3px !important;
    height: 8px !important;
}}

/* ── Radio ── */
[data-testid="stRadio"] label {{ font-size: 0.9rem !important; }}

/* ── Footer ── */
.monash-footer {{
    text-align: center;
    color: #aaa;
    font-size: 0.72rem;
    margin-top: 2.5rem;
    padding-top: 1rem;
    border-top: 1px solid {MONASH_LIGHT_GREY};
}}
</style>
""", unsafe_allow_html=True)

# ── Monash header ────────────────────────────────────────────
st.markdown("""
<div class="monash-header">
    <div class="monash-logo-text">M</div>
    <div class="monash-header-text">
        <h1>Placement Schedule Generator</h1>
        <p>Faculty of Education — Student Placement Office</p>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# STEP 1 — inside a bordered container
# ════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("""
    <div class="step-header">
      <span class="step-badge">Step 1</span>
      <p class="step-title">Upload your Google Form CSV</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Upload CSV",
        type="csv",
        label_visibility="collapsed",
    )

    if uploaded:
        try:
            raw = load_csv(uploaded)
        except Exception as e:
            st.error(f"Could not read the CSV: {e}")
            st.stop()

        try:
            year = extract_year(raw)
            ordered_dates, day_lookup = detect_placement_dates(raw, year)
        except ValueError as e:
            st.error(str(e))
            st.stop()

        short_labels, week_of, week_labels, range_label = build_schedule_labels(ordered_dates)

        st.markdown(f"""
        <div class="metrics-row">
          <div class="metric-card">
            <div class="metric-label">Students</div>
            <div class="metric-value">{len(raw)}</div>
            <div class="metric-sub">form responses</div>
          </div>
          <div class="metric-card">
            <div class="metric-label">Schools</div>
            <div class="metric-value">{raw['school'].nunique()}</div>
            <div class="metric-sub">unique schools</div>
          </div>
          <div class="metric-card">
            <div class="metric-label">Placement days</div>
            <div class="metric-value">15</div>
            <div class="metric-sub">working days</div>
          </div>
        </div>
        <div class="period-box">
          Detected placement period: <span>{range_label}</span>
        </div>
        """, unsafe_allow_html=True)

if not uploaded:
    st.markdown(
        f'<p style="text-align:center;color:#aaa;font-size:0.82rem;margin-top:0.4rem;">'
        f'Upload the CSV downloaded from Google Forms to continue</p>',
        unsafe_allow_html=True,
    )
    st.stop()

st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# STEP 2 — inside a bordered container
# ════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("""
    <div class="step-header">
      <span class="step-badge">Step 2</span>
      <p class="step-title">Choose schools to export</p>
    </div>
    """, unsafe_allow_html=True)

    all_schools = sorted(raw["school"].dropna().unique())

    export_mode = st.radio(
        "Export mode",
        ["All schools", "Selected schools only"],
        horizontal=True,
        label_visibility="collapsed",
    )

    if export_mode == "All schools":
        schools_to_export = all_schools
        st.markdown(
            f'<div class="pill-success"><span class="pill-dot"></span>'
            f'All {len(all_schools)} schools will be exported</div>',
            unsafe_allow_html=True,
        )
    else:
        selected = st.multiselect(
            "Search and select schools",
            options=all_schools,
            placeholder="Type a school name to search...",
            label_visibility="collapsed",
        )
        if not selected:
            st.markdown(
                f'<p style="color:{MONASH_GREY};font-size:0.85rem;margin-top:6px;">'
                f'Search and select at least one school above.</p>',
                unsafe_allow_html=True,
            )
            st.stop()
        schools_to_export = selected
        st.markdown(
            f'<div class="pill-success"><span class="pill-dot"></span>'
            f'{len(selected)} school(s) selected</div>',
            unsafe_allow_html=True,
        )

st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# STEP 3 — inside a bordered container
# ════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("""
    <div class="step-header">
      <span class="step-badge">Step 3</span>
      <p class="step-title">Generate and download</p>
    </div>
    """, unsafe_allow_html=True)

    n_schools = len(schools_to_export)
    generate  = st.button(
        f"Generate {n_schools} Excel file{'s' if n_schools != 1 else ''}",
        type="primary",
        use_container_width=True,
    )

    if generate:
        long_df      = expand_long(raw, day_lookup)
        progress_bar = st.progress(0)
        status_text  = st.empty()
        schools_data = []

        for idx, school in enumerate(schools_to_export):
            pct = (idx + 1) / n_schools
            progress_bar.progress(pct)
            status_text.markdown(
                f'<div class="progress-label">'
                f'Building {idx + 1} of {n_schools}: {school}</div>',
                unsafe_allow_html=True,
            )
            mat      = build_matrix(school, raw, long_df, ordered_dates)
            wb       = build_workbook(school, mat, ordered_dates,
                                      short_labels, week_of, week_labels, range_label)
            wb_bytes = workbook_to_bytes(wb)
            safe     = re.sub(r'[\\/:*?"<>|]', "_", school).strip()
            schools_data.append((f"{safe}.xlsx", wb_bytes))

        progress_bar.empty()
        status_text.empty()

        n = len(schools_data)
        st.markdown(
            f'<div class="success-box">'
            f'All {n} file{"s" if n != 1 else ""} generated successfully.</div>',
            unsafe_allow_html=True,
        )

        if n == 1:
            filename, wb_bytes = schools_data[0]
            st.download_button(
                label=f"Download {filename}",
                data=wb_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            zip_bytes = build_zip(schools_data)
            st.download_button(
                label=f"Download all {n} files as ZIP",
                data=zip_bytes,
                file_name=f"placement_schedules_{ordered_dates[0].year}.zip",
                mime="application/zip",
                use_container_width=True,
            )

# ── Footer ───────────────────────────────────────────────────
st.markdown(
    '<div class="monash-footer">'
    'Monash University &nbsp;·&nbsp; Faculty of Education &nbsp;·&nbsp; Placement Schedule Generator'
    '</div>',
    unsafe_allow_html=True,
)
# placeholder - reading current file
