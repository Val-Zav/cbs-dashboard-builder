"""
CBS Portfolio Tools -- Unified Streamlit App
Two tools, one interface:
  - Tab 1 -- CBS Portfolio Dashboard  (build_core.py)
  - Tab 2 -- CBS Oversight Tracker    (build_tracker_core.py)

Both tools accept the same four source Excel files.
"""

import io
import os
import sys
import tempfile
import datetime
import streamlit as st

# -- Page config (must be the first Streamlit call) -----------------------------
st.set_page_config(
    page_title="CBS Portfolio Tools",
    page_icon=":bar_chart:",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# -- SAP logo (inline SVG) ------------------------------------------------------
_SAP_SVG = """<svg version="1.1" xmlns="http://www.w3.org/2000/svg"
  xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 412.4 204"
  style="height:40px;display:inline-block;vertical-align:middle;">
<defs>
  <linearGradient id="sg" x1="206.19" y1="206" x2="206.19" y2="2"
    gradientUnits="userSpaceOnUse" gradientTransform="matrix(1 0 0 -1 0 206)">
    <stop offset="0"    stop-color="#00B8F1"/>
    <stop offset="0.02" stop-color="#01B6F0"/>
    <stop offset="0.31" stop-color="#0D90D9"/>
    <stop offset="0.58" stop-color="#1775C8"/>
    <stop offset="0.82" stop-color="#1C65BF"/>
    <stop offset="1"    stop-color="#1E5FBB"/>
  </linearGradient>
</defs>
<polyline style="fill-rule:evenodd;clip-rule:evenodd;fill:url(#sg)"
  points="0,204 208.4,204 412.4,0 0,0 0,204"/>
<path style="fill-rule:evenodd;clip-rule:evenodd;fill:#FFFFFF"
  d="M244.7,38.4h-40.6v96.5l-35.5-96.6h-35.2l-30.3,80.7C100,98.7,79,91.7,62.4,86.4
  C51.5,82.9,39.8,77.7,40,72c0.1-4.7,6.2-9,18.4-8.4c8.2,0.4,15.4,1.1,29.7,8l14.1-24.5
  c-13.1-6.6-31.2-10.9-46-10.9h-0.1c-17.3,0-31.7,5.6-40.6,14.8c-6.2,6.3-9.7,14.8-9.7,23.7
  C5.5,87.2,10.1,96,19.7,103c8.1,5.9,18.5,9.8,27.6,12.6c11.3,3.5,20.5,6.5,20.4,13
  c-0.1,2.4-1,4.7-2.7,6.4c-2.8,2.9-7.1,4-13.1,4.1c-11.5,0.2-20-1.6-33.6-9.6L5.8,154.4
  c14,8,29.9,12.2,46,12.2h2.1c14.2-0.2,25.7-4.3,34.9-11.7c0.5-0.4,1-0.8,1.5-1.3l-4.1,10.9
  H123l6.2-18.8c7,2.3,14.3,3.5,21.7,3.4c7.2,0,14.3-1.1,21.2-3.2l6,18.6h60.1v-39h13.1
  c31.7,0,50.5-16.2,50.5-43.2C301.7,52.2,283.5,38.4,244.7,38.4z
  M150.9,121c-4.4,0-8.8-0.7-13-2.3l12.9-40.6h0.2l12.6,40.7C159.6,120.3,155.2,121,150.9,121z
  M247.1,97.7h-8.9V64.9h8.9c11.9,0,21.4,4,21.4,16.1C268.5,93.7,259,97.6,247.1,97.7"/>
</svg>"""

# -- Global CSS -----------------------------------------------------------------
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');

  html, body, [class*="css"] {
    font-family: '72 Brand', '72', 'IBM Plex Sans', Arial, sans-serif;
  }

  /* Header */
  .sap-header {
    background: linear-gradient(135deg, #0057B8 0%, #0070F2 60%, #4CB1FF 100%);
    border-radius: 10px;
    padding: 22px 28px;
    display: flex;
    align-items: center;
    gap: 20px;
    margin-bottom: 8px;
  }
  .sap-header-text h1 {
    color: #fff;
    font-size: 1.35rem;
    font-weight: 700;
    margin: 0 0 4px 0;
  }
  .sap-header-text p {
    color: rgba(255,255,255,0.82);
    font-size: 0.82rem;
    margin: 0;
  }

  /* Upload card */
  .upload-card {
    background: #F5F6F8;
    border: 1px solid #E1E2E6;
    border-radius: 8px;
    padding: 14px 18px 4px;
    margin-bottom: 4px;
  }
  .upload-card h3 { font-size: 0.95rem; font-weight: 600; color: #1D2329; margin-bottom: 4px; }
  .upload-card p  { font-size: 0.8rem; color: #89919A; margin: 0; }

  /* Status badges */
  .badge-ok   { background: #E6F4EA; color: #107E3E; padding: 3px 10px;
                border-radius: 12px; font-size: 0.78rem; font-weight: 600; }
  .badge-miss { background: #F5F6F8; color: #89919A; padding: 3px 10px;
                border-radius: 12px; font-size: 0.78rem; font-weight: 600; }

  /* Build button */
  div[data-testid="stButton"] > button {
    background: #0070F2; color: white; border: none;
    border-radius: 6px; font-size: 1rem; font-weight: 600;
    padding: 0.6rem 2.4rem; width: 100%; cursor: pointer;
    transition: background 0.15s;
  }
  div[data-testid="stButton"] > button:hover   { background: #0040B0; }
  div[data-testid="stButton"] > button:disabled { background: #EAECEE; color: #89919A; cursor: not-allowed; }

  /* Download button */
  div[data-testid="stDownloadButton"] > button {
    background: #107E3E; color: white; border: none;
    border-radius: 6px; font-size: 1rem; font-weight: 600;
    padding: 0.6rem 2.4rem; width: 100%;
  }
  div[data-testid="stDownloadButton"] > button:hover { background: #0B5C2C; }

  /* Result boxes */
  .result-ok  { background: #F1FAF5; border: 1px solid #107E3E; border-radius: 8px;
                padding: 14px 20px; color: #0B5C2C; font-weight: 500; font-size: 0.9rem; }
  .result-err { background: #FFF3F3; border: 1px solid #BB0000; border-radius: 8px;
                padding: 14px 20px; color: #8B0000; font-weight: 500; font-size: 0.9rem; }

  /* Tabs */
  button[data-baseweb="tab"] { font-size: 0.9rem; font-weight: 600; }

  /* Footer */
  .sap-footer {
    margin-top: 32px; padding-top: 16px;
    border-top: 1px solid #EAECEE;
    text-align: center; font-size: 0.75rem; color: #89919A;
  }

  /* Hide Streamlit chrome */
  #MainMenu, footer { visibility: hidden; }
  header[data-testid="stHeader"] { display: none; }
</style>
""", unsafe_allow_html=True)

# -- Header ---------------------------------------------------------------------
st.markdown(f"""
<div class="sap-header">
  {_SAP_SVG}
  <div class="sap-header-text">
    <h1>CBS Portfolio Tools</h1>
    <p>Upload the four source Excel files to generate the interactive HTML reports.</p>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# -- Helper: file status row ----------------------------------------------------
def _status_row(files: dict):
    if not any(files.values()):
        return
    parts = []
    for label, f in files.items():
        if f:
            parts.append(f'<span class="badge-ok">&#10003; {label}</span>')
        else:
            parts.append(f'<span class="badge-miss">&#9711; {label}</span>')
    st.markdown(
        "<div style='margin:8px 0 12px 0; display:flex; flex-wrap:wrap; gap:10px;'>"
        + "  ".join(parts)
        + "</div>",
        unsafe_allow_html=True,
    )


# ===============================================================================
# SHARED FILE UPLOADERS (used by both Dashboard and Tracker)
# ===============================================================================
st.markdown('<div class="upload-card"><h3>Source Files</h3>'
            '<p>All four files are required. Upload once, use for both tools.</p></div>',
            unsafe_allow_html=True)

u_col1, u_col2 = st.columns(2)
with u_col1:
    u_si  = st.file_uploader("Services Integrated",   type=["xlsx"], key="u_si",  help="Services_Integrated_Project_Fi.xlsx")
    u_red = st.file_uploader("Red Project Data Base",  type=["xlsx"], key="u_red", help="Red Project Data Base.xlsx")
with u_col2:
    u_ma  = st.file_uploader("MANDI",                  type=["xlsx"], key="u_ma",  help="MANDI.xlsx")
    u_lk  = st.file_uploader("Leakage Report",         type=["xlsx"], key="u_lk",  help="Leakage Report.xlsx")

shared_files = {"Services Integrated": u_si, "MANDI": u_ma, "Red Project": u_red, "Leakage": u_lk}
_status_row(shared_files)

files_ready = all(shared_files.values())
if not files_ready:
    remaining = 4 - sum(1 for v in shared_files.values() if v)
    st.caption(f"{remaining} file{'s' if remaining != 1 else ''} still needed.")

st.markdown("<br>", unsafe_allow_html=True)

# ===============================================================================
# TABS
# ===============================================================================
tab_dash, tab_track = st.tabs(["CBS Portfolio Dashboard", "CBS Oversight Tracker"])


# +===========================================================================+
# |  TAB 1 -- CBS PORTFOLIO DASHBOARD                                         |
# +===========================================================================+
with tab_dash:

    with st.expander("How to use -- Portfolio Dashboard", expanded=False):
        st.markdown("""
**Three steps:**

1. Upload the four Excel files above.
2. Click **Build Dashboard**.
3. Download the full version and the public version (without personal data).

The dashboards are self-contained HTML files that anyone can open in a browser.

---
| File | Contents |
|---|---|
| `Services_Integrated_Project_Fi.xlsx` | Project financials, margins, EAC |
| `MANDI.xlsx` | Delivery status and LoS data |
| `Red Project Data Base.xlsx` | Weekly red project history |
| `Leakage Report.xlsx` | Leakage and billing type data |
""")

    st.markdown("<br>", unsafe_allow_html=True)

    d_build = st.button("Build Dashboard", disabled=not files_ready, use_container_width=True, key="d_build_btn")

    if d_build and files_ready:
        progress = st.progress(0, text="Starting build...")
        try:
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            import build_core  # noqa: E402

            with tempfile.TemporaryDirectory() as tmp:
                def _save(uploaded, name):
                    path = os.path.join(tmp, name)
                    with open(path, "wb") as fh:
                        fh.write(uploaded.getbuffer())
                    return path

                progress.progress(10, text="Saving uploaded files...")
                p_si  = _save(u_si,  "Services_Integrated.xlsx")
                p_ma  = _save(u_ma,  "MANDI.xlsx")
                p_red = _save(u_red, "Red_Project_Data_Base.xlsx")
                p_lk  = _save(u_lk,  "Leakage_Report.xlsx")

                progress.progress(20, text="Building full dashboard...")
                html_full = build_core.build(p_si, p_ma, p_red, p_lk)

                progress.progress(80, text="Building public dashboard (no personal data)...")
                html_public = build_core.build_public(p_si, p_ma, p_red, p_lk)

            progress.progress(100, text="Done.")
            st.session_state["d_html_full"]   = html_full
            st.session_state["d_html_public"] = html_public
            st.session_state["d_build_ok"]    = True

        except Exception as exc:
            progress.empty()
            st.markdown(
                f'<div class="result-err"><strong>Build failed.</strong><br>{exc}</div>',
                unsafe_allow_html=True,
            )
            st.session_state["d_build_ok"] = False

    if st.session_state.get("d_build_ok"):
        progress_elem = None
        try:
            progress_elem.empty()
        except Exception:
            pass
        st.markdown('<div class="result-ok">Dashboard built successfully. Download below.</div>',
                    unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                label="Download Full Dashboard",
                data=st.session_state["d_html_full"].encode("utf-8"),
                file_name="CBS_Portfolio_Dashboard.html",
                mime="text/html",
                use_container_width=True,
                key="d_dl_full",
            )
        with dl2:
            st.download_button(
                label="Download Public Dashboard",
                data=st.session_state["d_html_public"].encode("utf-8"),
                file_name="CBS_Portfolio_Dashboard_Public.html",
                mime="text/html",
                use_container_width=True,
                key="d_dl_public",
            )

    st.markdown('<div class="sap-footer">CBS Portfolio Operations &bull; SAP</div>',
                unsafe_allow_html=True)


# +===========================================================================+
# |  TAB 2 -- CBS OVERSIGHT TRACKER                                           |
# +===========================================================================+
with tab_track:

    with st.expander("How to use -- Oversight Tracker", expanded=False):
        st.markdown("""
**Steps:**

1. Upload the four current-week Excel files above.
2. *(Optional)* Expand **Previous Week Baseline** and upload last week's files to enable week-over-week delta indicators.
3. Set the report date (defaults to today).
4. Click **Build Tracker**.
5. Download the self-contained HTML report.

---
| File | Contents |
|---|---|
| `Services_Integrated_Project_Fi.xlsx` | Active projects, EAC, FELIPE fields |
| `MANDI.xlsx` | MANDI status, CBS Responsible, SO PM |
| `Red Project Data Base.xlsx` | Red/Yellow report history |
| `Leakage Report.xlsx` | Backlog leakage by item |

The tracker produces **4 sections**: Red Status, Positive Leakage, High Negative Leakage, Missing FELIPE Snapshot.  
All sections are filterable by Market Unit, Portfolio Segment, CBS Responsible, Contract Size, and Lifecycle Status.
""")

    # -- Previous week baseline (optional) ------------------------------------
    with st.expander("Previous Week Baseline (optional -- enables delta indicators)", expanded=False):
        st.markdown(
            "Upload the same four files from the **previous week** to display "
            "week-over-week change indicators (^ / v) on the KPI cards.",
            unsafe_allow_html=False,
        )
        b_col1, b_col2 = st.columns(2)
        with b_col1:
            b_si  = st.file_uploader("Services Integrated (prev week)",   type=["xlsx"], key="b_si")
            b_red = st.file_uploader("Red Project Data Base (prev week)", type=["xlsx"], key="b_red")
        with b_col2:
            b_ma  = st.file_uploader("MANDI (prev week)",                 type=["xlsx"], key="b_ma")
            b_lk  = st.file_uploader("Leakage Report (prev week)",        type=["xlsx"], key="b_lk")

        b_files = {"Services Integrated": b_si, "MANDI": b_ma, "Red Project": b_red, "Leakage": b_lk}
        _status_row(b_files)
        b_ready = all(b_files.values())

        if b_ready:
            b_date = st.date_input(
                "Previous week date",
                value=datetime.date.today() - datetime.timedelta(days=7),
                key="b_date_input",
                help="The date of the previous week's data snapshot (used for the delta label).",
            )
        else:
            b_date = None

    # -- Report date -----------------------------------------------------------
    today_val = st.date_input(
        "Report date",
        value=datetime.date.today(),
        key="t_date_input",
        help="The date displayed in the report header and used for snapshot-window calculations.",
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # -- Build button ----------------------------------------------------------
    t_build = st.button(
        "Build Tracker",
        disabled=not files_ready,
        use_container_width=True,
        key="t_build_btn",
    )

    if t_build and files_ready:
        t_progress = st.progress(0, text="Starting build...")
        try:
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            import build_tracker_core  # noqa: E402

            t_progress.progress(10, text="Reading current-week files...")
            si_bytes  = u_si.getvalue()
            ma_bytes  = u_ma.getvalue()
            red_bytes = u_red.getvalue()
            lk_bytes  = u_lk.getvalue()

            baseline_args = {}
            if b_ready:
                t_progress.progress(20, text="Reading baseline files...")
                baseline_args = dict(
                    baseline_si_src  = b_si.getvalue(),
                    baseline_m_src   = b_ma.getvalue(),
                    baseline_r_src   = b_red.getvalue(),
                    baseline_l_src   = b_lk.getvalue(),
                    baseline_date    = str(b_date),
                )

            t_progress.progress(40, text="Running data pipeline...")
            html_tracker = build_tracker_core.build_tracker(
                si_src      = si_bytes,
                m_src       = ma_bytes,
                r_src       = red_bytes,
                l_src       = lk_bytes,
                today_date  = str(today_val),
                **baseline_args,
            )

            t_progress.progress(100, text="Done.")
            st.session_state["t_html"]     = html_tracker
            st.session_state["t_build_ok"] = True
            st.session_state["t_filename"] = (
                f"CBS_Oversight_Tracker_{today_val.strftime('%Y%m%d')}.html"
            )

        except Exception as exc:
            t_progress.empty()
            st.markdown(
                f'<div class="result-err"><strong>Build failed.</strong><br>{exc}</div>',
                unsafe_allow_html=True,
            )
            import traceback
            st.code(traceback.format_exc(), language="python")
            st.session_state["t_build_ok"] = False

    if st.session_state.get("t_build_ok"):
        _html_bytes = st.session_state["t_html"].encode("utf-8")
        _fname      = st.session_state.get("t_filename", "CBS_Oversight_Tracker.html")
        _size_kb    = len(_html_bytes) / 1024

        st.markdown(
            f'<div class="result-ok">Tracker built successfully &mdash; '
            f'{_size_kb:,.0f} KB. Download below.</div>',
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="Download Oversight Tracker",
            data=_html_bytes,
            file_name=_fname,
            mime="text/html",
            use_container_width=True,
            key="t_dl_btn",
        )

        # Quick stats
        with st.expander("Build summary", expanded=False):
            st.caption(
                f"File: `{_fname}` &nbsp;|&nbsp; Size: {_size_kb:,.0f} KB"
            )

    st.markdown('<div class="sap-footer">CBS Portfolio Operations &bull; SAP</div>',
                unsafe_allow_html=True)
