"""
app.py  —  Placement Schedule Generator (Streamlit web app)
"""

import re
import streamlit as st
from core import (
    load_inplace,
    load_form,
    get_available_years,
    filter_inplace,
    get_status_breakdown,
    get_no_placement_students,
    detect_placement_dates,
    extract_year_from_form,
    build_schedule_labels,
    build_master,
    build_no_placement_df,
    build_matrix_from_master,
    build_workbook,
    build_no_placement_workbook,
    workbook_to_bytes,
    build_zip,
)

# ── Page config ──────────────────────────────────────────────
st.set_page_config(
    page_title="Placement Schedule Generator",
    page_icon="🎓",
    layout="centered",
)

MONASH_BLUE       = "#006DAE"
MONASH_BLUE_DARK  = "#005490"
MONASH_BLUE_LIGHT = "#E6F2F9"
MONASH_GREEN      = "#008A25"
MONASH_DARK       = "#3C3C3C"
MONASH_GREY       = "#5A5A5A"
MONASH_LIGHT_GREY = "#E6E6E6"
MONASH_BG         = "#F6F6F6"

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
}}
.monash-header-text p {{
    color: rgba(255,255,255,0.78);
    font-size: 0.82rem;
    margin: 0;
}}
div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"] {{
    background: white !important;
    border: 1px solid {MONASH_LIGHT_GREY} !important;
    border-top: 3px solid {MONASH_BLUE} !important;
    border-radius: 8px !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06) !important;
    padding: 4px 8px 24px !important;
}}
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
.info-box {{
    background: #FFF8E1;
    border: 1px solid #FFE082;
    border-left: 4px solid #F9A825;
    border-radius: 6px;
    padding: 10px 14px;
    font-size: 0.83rem;
    color: #5D4037;
    margin-top: 8px;
    margin-bottom: 8px;
}}
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
.success-box {{
    background: #E8F5EE;
    border: 1px solid #A8D5BB;
    border-left: 4px solid {MONASH_GREEN};
    border-radius: 6px;
    padding: 12px 16px;
    margin: 8px 0;
    font-size: 0.9rem;
    font-weight: 600;
    color: #1B5E35;
}}
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
[data-testid="stFileUploader"] {{
    border: 2px dashed {MONASH_LIGHT_GREY} !important;
    border-radius: 8px !important;
    background: {MONASH_BG} !important;
}}
[data-testid="stProgress"] > div > div {{
    background: {MONASH_BLUE} !important;
    border-radius: 3px !important;
}}
[data-testid="stProgress"] > div {{
    background: {MONASH_LIGHT_GREY} !important;
    border-radius: 3px !important;
    height: 8px !important;
}}
[data-testid="stRadio"] label {{ font-size: 0.9rem !important; }}
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

# ── Header ───────────────────────────────────────────────────
st.markdown("""
<div class="monash-header">
    <div class="monash-logo-text">M</div>
    <div class="monash-header-text">
        <h1>Placement Schedule Generator</h1>
        <p>Faculty of Education — Student Placement Office</p>
    </div>
</div>
<br>
""", unsafe_allow_html=True)

tab_main, tab_help, tab_inplace, tab_form = st.tabs(["Generate Schedules", "Help & Guide", "How to Download InPlace", "How to Download Form"])

with tab_help:
    st.markdown(f"""
    <div style="padding:0.5rem 0">
    <h3 style="color:{MONASH_BLUE};font-size:1.1rem;font-weight:700;margin-bottom:4px;">Quick Reference</h3>
    <p style="color:{MONASH_GREY};font-size:0.85rem;margin-bottom:1.2rem;">
        For the full step-by-step guide with screenshots, download the Word document below.
    </p>
    </div>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"""
        <div class="step-header">
          <span class="step-badge">What you need</span>
          <p class="step-title">Two files before you start</p>
        </div>
        <table style="width:100%;border-collapse:collapse;font-size:0.85rem;">
          <tr style="background:{MONASH_BLUE};color:white;">
            <th style="padding:8px 12px;text-align:left;font-weight:600;">File</th>
            <th style="padding:8px 12px;text-align:left;font-weight:600;">Where it comes from</th>
            <th style="padding:8px 12px;text-align:left;font-weight:600;">Format</th>
          </tr>
          <tr style="background:#F6F6F6;">
            <td style="padding:8px 12px;font-weight:600;color:{MONASH_BLUE};">InPlace export</td>
            <td style="padding:8px 12px;">Downloaded from the InPlace system</td>
            <td style="padding:8px 12px;">.xlsx or .csv</td>
          </tr>
          <tr>
            <td style="padding:8px 12px;font-weight:600;color:{MONASH_BLUE};">Google Form responses</td>
            <td style="padding:8px 12px;">Downloaded from Google Sheets (linked to the Form)</td>
            <td style="padding:8px 12px;">.xlsx or .csv</td>
          </tr>
        </table>
        <div style="margin-top:12px;margin-bottom:4px;background:#FFF8E1;border-left:4px solid #F9A825;border-radius:4px;padding:8px 12px;font-size:0.82rem;color:#5D4037;">
          <strong>Important:</strong> Do NOT open and re-save the CSV in Excel before uploading — this can corrupt the date format.
        </div>
        """, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"""
        <div class="step-header">
          <span class="step-badge">Steps</span>
          <p class="step-title">How to generate the schedules</p>
        </div>
        """, unsafe_allow_html=True)

        steps_html = ""
        steps = [
            ("1", "Upload both files", "In the Generate Schedules tab, upload the InPlace file on the left and the Google Form CSV on the right."),
            ("2", "Select the year", "A year dropdown appears automatically based on the InPlace data. Choose the correct placement year."),
            ("3", "Check the summary", "Confirm the detected placement period, student count, and school count look correct."),
            ("4", "Choose schools", "Select All schools to export everything, or use Selected schools only to test with one school first."),
            ("5", "Generate", "Click the Generate button and wait for the progress bar to complete. Do not close the browser."),
            ("6", "Download", "Click the Download button to save your files. Multiple schools are bundled into a ZIP file."),
        ]
        for num, title, detail in steps:
            steps_html += f"""
            <div style="display:flex;gap:12px;align-items:flex-start;margin-bottom:12px;">
              <div style="background:{MONASH_BLUE};color:white;font-size:0.75rem;font-weight:700;
                          width:24px;height:24px;border-radius:50%;display:flex;align-items:center;
                          justify-content:center;flex-shrink:0;margin-top:2px;">{num}</div>
              <div>
                <div style="font-size:0.88rem;font-weight:600;color:{MONASH_DARK};margin-bottom:2px;">{title}</div>
                <div style="font-size:0.82rem;color:{MONASH_GREY};">{detail}</div>
              </div>
            </div>"""
        st.markdown(steps_html, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"""
        <div class="step-header">
          <span class="step-badge">Common errors</span>
          <p class="step-title">Troubleshooting</p>
        </div>
        <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
          <tr style="background:{MONASH_BLUE};color:white;">
            <th style="padding:8px 12px;text-align:left;font-weight:600;">Error</th>
            <th style="padding:8px 12px;text-align:left;font-weight:600;">Fix</th>
          </tr>
          <tr style="background:#F6F6F6;">
            <td style="padding:8px 12px;font-weight:600;">Year mismatch detected</td>
            <td style="padding:8px 12px;">Download both files for the same placement year and re-upload.</td>
          </tr>
          <tr>
            <td style="padding:8px 12px;font-weight:600;">InPlace file is missing columns</td>
            <td style="padding:8px 12px;">Re-download from InPlace without editing the file first.</td>
          </tr>
          <tr style="background:#F6F6F6;">
            <td style="padding:8px 12px;font-weight:600;">No students found for selected year</td>
            <td style="padding:8px 12px;">Check you have the correct year's InPlace export uploaded.</td>
          </tr>
          <tr>
            <td style="padding:8px 12px;font-weight:600;">No valid placement dates in Form</td>
            <td style="padding:8px 12px;">Make sure the Form CSV has not been opened and re-saved in Excel.</td>
          </tr>
          <tr style="background:#F6F6F6;">
            <td style="padding:8px 12px;font-weight:600;">Page does not load</td>
            <td style="padding:8px 12px;">Wait 1 minute and refresh. Contact your IT support if it continues.</td>
          </tr>
        </table>
        <p style="margin-top:10px;margin-bottom:4px;font-size:0.8rem;color:{MONASH_GREY};">
          For anything not listed above, download the full guide below or contact your placement team administrator.
        </p>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f'<p style="font-size:0.85rem;font-weight:600;color:{MONASH_DARK};">Full User Guide (Word document)</p>',
        unsafe_allow_html=True
    )
    try:
        with open("Placement_Schedule_Generator_User_Guide.docx", "rb") as f:
            st.download_button(
                label="Download User Guide (.docx)",
                data=f.read(),
                file_name="Placement_Schedule_Generator_User_Guide.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
    except FileNotFoundError:
        st.info("Place the file 'Placement_Schedule_Generator_User_Guide.docx' in the same folder as app.py to enable this download.")

with tab_inplace:
    st.markdown("""
    <div style="padding:0.5rem 0 0.2rem">
    <h3 style="color:#2E7D32;font-size:1.1rem;font-weight:700;margin-bottom:4px;">Downloading from InPlace (iReport)</h3>
    <p style="color:#5A5A5A;font-size:0.85rem;margin-bottom:1.2rem;">
        Follow these steps each year to get the InPlace student export file.
    </p>
    </div>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"""
        <div class="step-header">
          <span class="step-badge" style="background:#E65100;">Read this first</span>
          <p class="step-title">Email must be added manually</p>
        </div>
        <div style="background:#FFF3E0;border:2px solid #F57C00;border-radius:8px;
                    padding:16px 18px;">
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">
            <span style="font-size:1.4rem;">⚠️</span>
            <span style="font-size:0.95rem;font-weight:700;color:#E65100;">
              This column is NOT included by default
            </span>
          </div>
          <p style="font-size:0.87rem;color:{MONASH_DARK};margin:0 0 12px;line-height:1.7;">
            <strong>Email</strong> does not appear in the standard
            iReport export. If you skip this step, the app will not be able to match students
            to their Google Form responses, and will show an error when you try to upload the file.
          </p>
          <div style="background:white;border-radius:6px;padding:14px 16px;">
            <p style="font-size:0.85rem;font-weight:700;color:#E65100;margin:0 0 8px;">
              Before downloading from iReport, do this:
            </p>
            <ol style="font-size:0.86rem;color:{MONASH_DARK};margin:0;padding-left:22px;line-height:2;">
              <li>Open the report in iReport</li>
              <li>Find the <strong>"Add Column"</strong> option (usually near the report filters or column settings)</li>
              <li>Tick the box next to <strong>Email</strong></li>
              <li>Apply the changes — confirm the column now appears in the report preview</li>
              <li>Then download the report as usual</li>
            </ol>
          </div>
        </div>
        """, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"""
        <div class="step-header">
          <span class="step-badge" style="background:#2E7D32;">What to download</span>
          <p class="step-title">iReport — 1st & 2nd year placement (non-method, non-HPE)</p>
        </div>
        <p style="font-size:0.85rem;color:{MONASH_GREY};margin-bottom:8px;">
          The InPlace export must include all of these columns. Do not delete or rename any columns after downloading.
        </p>
        <table style="width:100%;border-collapse:collapse;font-size:0.82rem;margin-top:8px;">
          <tr style="background:#1B5E20;color:white;">
            <th style="padding:8px 12px;text-align:left;font-weight:600;width:35%;">Column name</th>
            <th style="padding:8px 12px;text-align:left;font-weight:600;">What it is used for</th>
          </tr>
          <tr style="background:#E8F5E9;"><td style="padding:7px 12px;font-weight:600;color:#2E7D32;">Student</td><td style="padding:7px 12px;">Full student name</td></tr>
          <tr style="background:#FFF3E0;"><td style="padding:7px 12px;font-weight:600;color:#E65100;">Email&nbsp;⚠️</td><td style="padding:7px 12px;">Matches the student to their Google Form response — <strong>must be added via "Add Column" in iReport</strong></td></tr>
          <tr style="background:#E8F5E9;"><td style="padding:7px 12px;font-weight:600;color:#2E7D32;">Agency</td><td style="padding:7px 12px;">School name — one Excel file is created per school</td></tr>
          <tr><td style="padding:7px 12px;font-weight:600;color:#2E7D32;">Start Date</td><td style="padding:7px 12px;">Used to filter by year (can be blank for unplaced students)</td></tr>
          <tr style="background:#E8F5E9;"><td style="padding:7px 12px;font-weight:600;color:#2E7D32;">End Date</td><td style="padding:7px 12px;">Defines the placement date window</td></tr>
          <tr><td style="padding:7px 12px;font-weight:600;color:#2E7D32;">Status</td><td style="padding:7px 12px;">Only Confirmed students are included in the export</td></tr>
          <tr style="background:#E8F5E9;"><td style="padding:7px 12px;font-weight:600;color:#2E7D32;">Requirement Groups</td><td style="padding:7px 12px;">Contains the year — includes unplaced students with blank Start Date</td></tr>
          <tr><td style="padding:7px 12px;font-weight:600;color:#2E7D32;">Placement Allocation Groups</td><td style="padding:7px 12px;">Backup year filter</td></tr>
        </table>
        <p style="font-size:0.78rem;color:{MONASH_GREY};margin-top:10px;">
          Note: Student ID is sourced from the Google Form (Monash Student ID), not from InPlace —
          you do not need to add a Student Code column.
        </p>
        """, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"""
        <div class="step-header">
          <span class="step-badge" style="background:#2E7D32;">Steps</span>
          <p class="step-title">How to download</p>
        </div>
        """, unsafe_allow_html=True)

        steps_ip = [
            ("1", "Log in to InPlace", "Open your browser and go to the InPlace website. Sign in with your Monash staff credentials."),
            ("2", "Go to iReport", "From the InPlace menu, navigate to iReport. [Add your team's exact navigation path here.]"),
            ("3", "Select the correct report", "Choose the placement report for 1st and 2nd year students, non-method and non-HPE, for the current year."),
            ("4", "Download as Excel", "Click the download or export button. Choose Excel (.xlsx). Save it to your Desktop or Downloads folder."),
            ("5", "Do not edit the file", "Upload it to the app exactly as downloaded. Do not open it in Excel or rename any columns."),
        ]
        html = ""
        for num, title, detail in steps_ip:
            html += f"""
            <div style="display:flex;gap:12px;align-items:flex-start;margin-bottom:14px;">
              <div style="background:#2E7D32;color:white;font-size:0.75rem;font-weight:700;
                          width:26px;height:26px;border-radius:50%;display:flex;align-items:center;
                          justify-content:center;flex-shrink:0;margin-top:1px;">{num}</div>
              <div>
                <div style="font-size:0.88rem;font-weight:600;color:{MONASH_DARK};margin-bottom:3px;">{title}</div>
                <div style="font-size:0.82rem;color:{MONASH_GREY};">{detail}</div>
              </div>
            </div>"""
        st.markdown(html, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"""
        <div class="step-header">
          <span class="step-badge" style="background:#2E7D32;">Auto-filters</span>
          <p class="step-title">What the app removes automatically — nothing you need to do</p>
        </div>
        <ul style="font-size:0.83rem;color:{MONASH_DARK};margin:0;padding-left:20px;line-height:1.9;">
          <li>Schools starting with <strong>EDU -</strong> are excluded (these students don't need a placement at all)</li>
          <li>Only students with <strong>Status = Confirmed</strong> are included in school exports — Withdrawn, blank, and any other status are excluded</li>
          <li>Students with <strong>blank Status AND blank Agency</strong> (and not EDU-) are treated as <strong>awaiting placement</strong> — they go into the separate "No Placement" file instead of a school file</li>
          <li>Only students whose <strong>Requirement Groups</strong> or <strong>Placement Allocation Groups</strong> contain the selected year are included</li>
        </ul>
        <p style="margin-top:10px;font-size:0.8rem;color:{MONASH_GREY};">
          The Step 1 summary in the Generate Schedules tab shows how many students were Confirmed, Withdrawn, awaiting placement, or had another status, so you can see the full picture before exporting.
        </p>
        """, unsafe_allow_html=True)


with tab_form:
    st.markdown(f"""
    <div style="padding:0.5rem 0 0.2rem">
    <h3 style="color:#6A1B9A;font-size:1.1rem;font-weight:700;margin-bottom:4px;">Downloading the Google Form Responses</h3>
    <p style="color:{MONASH_GREY};font-size:0.85rem;margin-bottom:1.2rem;">
        Follow these steps each year to download the student availability responses.
    </p>
    </div>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"""
        <div class="step-header">
          <span class="step-badge" style="background:#6A1B9A;">Download link</span>
          <p class="step-title">Go to this Google Sheet to download responses</p>
        </div>
        <p style="font-size:0.88rem;color:{MONASH_DARK};margin-bottom:8px;">
          Click the link below to open the Google Sheet, then follow the steps underneath.
        </p>
        <a href="https://docs.google.com/spreadsheets/d/1n7dj0ikEaxphvFeGIFAMlUXpkcB9WEqrzq4b733dG04/edit"
           target="_blank"
           style="display:inline-block;background:#6A1B9A;color:white;padding:8px 18px;
                  border-radius:6px;font-size:0.85rem;font-weight:600;text-decoration:none;margin-bottom:8px;">
          Open Google Sheet (Form Responses)
        </a>
        """, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"""
        <div class="step-header">
          <span class="step-badge" style="background:#6A1B9A;">Steps</span>
          <p class="step-title">How to download</p>
        </div>
        """, unsafe_allow_html=True)

        steps_form = [
            ("1", "Open the Google Sheet", "Click the purple button above to open the sheet. Make sure you are logged in with your Monash Google account."),
            ("2", "Click File in the top menu", "In Google Sheets, click File in the top left menu bar."),
            ("3", "Click Download, then CSV", "Hover over Download, then click Comma Separated Values (.csv). The file downloads automatically."),
            ("4", "Save the file — do not open it", "Save it to your Desktop or Downloads folder. Do NOT open it in Excel before uploading — Excel changes the date format inside and causes the app to fail."),
        ]
        html = ""
        for num, title, detail in steps_form:
            html += f"""
            <div style="display:flex;gap:12px;align-items:flex-start;margin-bottom:14px;">
              <div style="background:#6A1B9A;color:white;font-size:0.75rem;font-weight:700;
                          width:26px;height:26px;border-radius:50%;display:flex;align-items:center;
                          justify-content:center;flex-shrink:0;margin-top:1px;">{num}</div>
              <div>
                <div style="font-size:0.88rem;font-weight:600;color:{MONASH_DARK};margin-bottom:3px;">{title}</div>
                <div style="font-size:0.82rem;color:{MONASH_GREY};">{detail}</div>
              </div>
            </div>"""
        st.markdown(html, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"""
        <div class="step-header">
          <span class="step-badge" style="background:#6A1B9A;">Updating dates each year</span>
          <p class="step-title">Only change the date options — nothing else</p>
        </div>
        <p style="font-size:0.85rem;color:{MONASH_DARK};margin-bottom:10px;">
          Each year, update the date options in the <strong>'What days will you be attending placement?'</strong> question.
          Everything else in the form stays the same.
        </p>
        <a href="https://docs.google.com/forms/d/1qITdAUy9iBiw_dLbwkgibNyctydnzqxlMrcf9QHXRMk/edit"
           target="_blank"
           style="display:inline-block;background:#6A1B9A;color:white;padding:8px 18px;
                  border-radius:6px;font-size:0.85rem;font-weight:600;text-decoration:none;margin-bottom:14px;">
          Open Google Form to Edit
        </a>
        <div style="background:#F3E5F5;border-left:4px solid #6A1B9A;border-radius:4px;
                    padding:10px 14px;font-size:0.83rem;color:{MONASH_DARK};margin-bottom:10px;">
          <strong style="color:#6A1B9A;">Date format — must be exact:</strong><br>
          Write each date as &nbsp;<strong>Weekday &nbsp; Month &nbsp; Day</strong><br>
          <span style="color:#6A1B9A;font-weight:600;">Monday July 21 &nbsp;&nbsp; Tuesday July 22 &nbsp;&nbsp; Friday August 8</span><br>
          <span style="color:{MONASH_GREY};font-size:0.8rem;">No year. No commas. No slashes. No ordinals (not 21st, not 22nd).</span>
        </div>
        <div style="background:#FFF3E0;border-left:4px solid #F57C00;border-radius:4px;
                    padding:10px 14px;font-size:0.83rem;color:{MONASH_DARK};">
          <strong style="color:#F57C00;">Column names must never change.</strong>
          The form has 8 columns (Timestamp, Email address, Name, etc.) — do not rename, delete, or reorder them.
          The app reads them by name.
        </div>
        """, unsafe_allow_html=True)


with tab_main:
    pass  # placeholder — actual content is indented below


# ════════════════════════════════════════════════════════════
# STEP 1 — Upload both files
# ════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("""
    <div class="step-header">
      <span class="step-badge">Step 1</span>
      <p class="step-title">Upload your files</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(
            f'<p style="font-size:0.83rem;font-weight:600;color:{MONASH_DARK};margin-bottom:6px;">'
            f'InPlace Export (CSV or Excel)</p>',
            unsafe_allow_html=True,
        )
        inplace_file = st.file_uploader(
            "InPlace", type=["csv", "xlsx", "xls"],
            label_visibility="collapsed", key="inplace"
        )

    with col2:
        st.markdown(
            f'<p style="font-size:0.83rem;font-weight:600;color:{MONASH_DARK};margin-bottom:6px;">'
            f'Google Form Response (CSV or Excel)</p>',
            unsafe_allow_html=True,
        )
        form_file = st.file_uploader(
            "Form", type=["csv", "xlsx", "xls"],
            label_visibility="collapsed", key="form"
        )

    if inplace_file and form_file:
        try:
            inplace_raw = load_inplace(inplace_file)
            form_df     = load_form(form_file)
        except ValueError as e:
            error_text = str(e)

            if "'Email'" in error_text and "iReport" in error_text:
                # Specific, friendly guidance for the most common mistake:
                # Email not ticked in iReport's Add Column option
                st.markdown(f"""
                <div style="background:#FFF3E0;border:1px solid #FFCC80;
                            border-left:5px solid #E65100;border-radius:8px;
                            padding:16px 20px;margin:8px 0;">
                  <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">
                    <span style="font-size:1.3rem;">⚠️</span>
                    <span style="font-size:1rem;font-weight:700;color:#E65100;">
                      Your InPlace file is missing <strong>Email</strong>
                    </span>
                  </div>
                  <p style="font-size:0.88rem;color:{MONASH_DARK};margin:0 0 10px;line-height:1.6;">
                    This column is <strong>not included by default</strong> when you export
                    from iReport — you must add it yourself before downloading.
                  </p>
                  <div style="background:white;border-radius:6px;padding:12px 16px;margin-bottom:10px;">
                    <p style="font-size:0.85rem;font-weight:600;color:{MONASH_DARK};margin:0 0 8px;">
                      How to fix this:
                    </p>
                    <ol style="font-size:0.85rem;color:{MONASH_DARK};margin:0;padding-left:20px;line-height:1.9;">
                      <li>Go back to <strong>iReport</strong> in InPlace</li>
                      <li>Find the <strong>"Add Column"</strong> option (near the report filters or column settings)</li>
                      <li>Tick the box next to <strong>Email</strong></li>
                      <li>Apply the changes, then download the report again</li>
                      <li>Upload the new file here</li>
                    </ol>
                  </div>
                  <p style="font-size:0.82rem;color:{MONASH_GREY};margin:0;">
                    See the <strong>"How to Download InPlace"</strong> tab above for a full walkthrough with more detail.
                  </p>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.error(error_text)
            st.stop()

        available_years = get_available_years(inplace_raw)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            f'<p style="font-size:0.83rem;font-weight:600;color:{MONASH_DARK};margin-bottom:6px;">'
            f'Select placement year</p>',
            unsafe_allow_html=True,
        )
        selected_year = st.selectbox(
            "Year", options=available_years,
            label_visibility="collapsed"
        )

        # Filter InPlace by year and other rules
        inplace_df, window_start, window_end = filter_inplace(inplace_raw, selected_year)

        # Status breakdown (Confirmed / Withdrawn / No Placement / Other)
        # for the selected year, calculated BEFORE the Confirmed-only filter
        # narrows things down — so staff can see the full picture.
        status_counts = get_status_breakdown(inplace_raw, selected_year)

        # No-placement students, sourced directly from InPlace
        # (blank Status AND blank Agency, excluding 'EDU -').
        no_placement_inplace = get_no_placement_students(inplace_raw, selected_year)

        # Detect placement dates from the form.
        # Form rows are filtered by timestamp year (students submit ~1 month
        # before placement, so timestamp year always matches placement year).
        try:
            ordered_dates, day_lookup = detect_placement_dates(form_df, selected_year)
        except Exception:
            # Work out what years ARE in the form so we can tell the user
            try:
                import pandas as _pd
                form_years = sorted(
                    _pd.to_datetime(form_df["timestamp"], dayfirst=True, errors="coerce")
                    .dt.year.dropna().astype(int).unique().tolist(),
                    reverse=True
                )
                form_years = [str(y) for y in form_years]
            except Exception:
                form_years = ["unknown"]

            st.error(
                f"**Year mismatch detected.**\n\n"
                f"You selected **{selected_year}** from the InPlace file, "
                f"but the Google Form responses are from **{', '.join(form_years)}**.\n\n"
                f"Please make sure both files are for the same placement year "
                f"and try again."
            )
            st.stop()

        short_labels, week_of, week_labels, range_label = build_schedule_labels(ordered_dates)

        if len(inplace_df) == 0:
            st.warning(
                f"No students found in the InPlace file for {selected_year}. "
                f"The available year(s) in this file are: {available_years}. "
                f"Please select a different year."
            )
            st.stop()

        # Build master join (Confirmed students with a real Agency only)
        master_df = build_master(inplace_df, form_df, day_lookup)

        # Build the No Placement list (sourced from InPlace, enriched with
        # Form data if the student happens to have submitted already)
        no_placement_df = build_no_placement_df(no_placement_inplace, form_df, day_lookup)

        all_schools  = sorted(master_df["agency"].dropna().unique())
        n_submitted  = master_df["submitted"].sum()
        n_no_sub     = (~master_df["submitted"]).sum()

        st.markdown(f"""
        <div class="metrics-row">
          <div class="metric-card">
            <div class="metric-label">InPlace students</div>
            <div class="metric-value">{len(inplace_df)}</div>
            <div class="metric-sub">after filters ({selected_year})</div>
          </div>
          <div class="metric-card">
            <div class="metric-label">Form submissions</div>
            <div class="metric-value">{n_submitted}</div>
            <div class="metric-sub">matched to InPlace</div>
          </div>
          <div class="metric-card">
            <div class="metric-label">No preference</div>
            <div class="metric-value">{n_no_sub}</div>
            <div class="metric-sub">in InPlace, no form</div>
          </div>
        </div>
        <div class="metrics-row">
          <div class="metric-card" style="border-left-color:#008A25;">
            <div class="metric-label">Confirmed</div>
            <div class="metric-value" style="color:#008A25;">{status_counts['confirmed']}</div>
            <div class="metric-sub">included in export</div>
          </div>
          <div class="metric-card" style="border-left-color:#C62828;">
            <div class="metric-label">Withdrawn</div>
            <div class="metric-value" style="color:#C62828;">{status_counts['withdrawn']}</div>
            <div class="metric-sub">excluded from export</div>
          </div>
          <div class="metric-card" style="border-left-color:#1565C0;">
            <div class="metric-label">No placement</div>
            <div class="metric-value" style="color:#1565C0;">{status_counts['no_placement']}</div>
            <div class="metric-sub">blank status &amp; agency</div>
          </div>
          <div class="metric-card" style="border-left-color:#F9A825;">
            <div class="metric-label">Other status</div>
            <div class="metric-value" style="color:#F9A825;">{status_counts['other']}</div>
            <div class="metric-sub">excluded from export</div>
          </div>
        </div>
        <div class="period-box">
          Detected placement period: <span>{range_label}</span>
        </div>
        """, unsafe_allow_html=True)

if not (inplace_file and form_file):
    st.markdown(
        f'<p style="text-align:center;color:#aaa;font-size:0.82rem;margin-top:0.4rem;">'
        f'Upload both files above to continue</p>',
        unsafe_allow_html=True,
    )
    st.stop()

st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# STEP 2 — Choose schools
# ════════════════════════════════════════════════════════════
with st.container(border=True):
    st.markdown("""
    <div class="step-header">
      <span class="step-badge">Step 2</span>
      <p class="step-title">Choose schools to export</p>
    </div>
    """, unsafe_allow_html=True)

    export_mode = st.radio(
        "Export mode",
        ["All schools", "Selected schools only"],
        horizontal=True,
        label_visibility="collapsed",
    )

    include_no_placement = st.checkbox(
        f"Include 'No Placement' file ({len(no_placement_df)} student(s) awaiting placement in InPlace)",
        value=True,
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

    if n_no_sub > 0:
        st.markdown(
            f'<div class="info-box">'
            f'{n_no_sub} student(s) in InPlace have not submitted the form. '
            f'They will appear at the bottom of their school sheet marked '
            f'"No preference submitted".</div>',
            unsafe_allow_html=True,
        )

st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# STEP 3 — Generate
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
        schools_data = []
        total_files  = n_schools + (1 if include_no_placement and len(no_placement_df) > 0 else 0)
        progress_bar = st.progress(0)
        status_text  = st.empty()

        for idx, school in enumerate(schools_to_export):
            progress_bar.progress((idx + 1) / total_files)
            status_text.markdown(
                f'<div class="progress-label">'
                f'Building {idx + 1} of {total_files}: {school}</div>',
                unsafe_allow_html=True,
            )
            mat      = build_matrix_from_master(school, master_df, ordered_dates)
            wb       = build_workbook(school, mat, ordered_dates,
                                      short_labels, week_of, week_labels, range_label)
            wb_bytes = workbook_to_bytes(wb)
            safe     = re.sub(r'[\\/:*?"<>|]', "_", school).strip()
            schools_data.append((f"{safe}.xlsx", wb_bytes))

        # No placement file
        if include_no_placement and len(no_placement_df) > 0:
            progress_bar.progress(1.0)
            status_text.markdown(
                f'<div class="progress-label">'
                f'Building: No Placement file...</div>',
                unsafe_allow_html=True,
            )
            wb_np    = build_no_placement_workbook(
                no_placement_df, ordered_dates, short_labels,
                week_of, week_labels, range_label
            )
            schools_data.append((
                f"No_Placement_{selected_year}.xlsx",
                workbook_to_bytes(wb_np)
            ))

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
                file_name=f"placement_schedules_{selected_year}.zip",
                mime="application/zip",
                use_container_width=True,
            )

st.markdown(
    '<div class="monash-footer">'
    'Monash University &nbsp;·&nbsp; Faculty of Education &nbsp;·&nbsp; Placement Schedule Generator'
    '</div>',
    unsafe_allow_html=True,
)
