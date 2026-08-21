"""
CBS Portfolio Dashboard Builder
Streamlit web app — upload 4 Excel files, download the HTML dashboard.
"""

import io
import base64
import tempfile
import os
import sys
import streamlit as st

# ── Page config (must be first Streamlit call) ────────────────────────────────
st.set_page_config(
    page_title="CBS Portfolio Dashboard Builder",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── SAP SVG logo (inline) ─────────────────────────────────────────────────────
SAP_SVG = """<svg version="1.1" xmlns="http://www.w3.org/2000/svg"
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
<polyline fill="url(#sg)" points="0,204 208.4,204 412.4,0 0,0 0,204"/>
<path fill="#FFFFFF" d="M244.7,38.4h-40.6v96.5l-35.5-96.6h-35.2l-30.3,80.7
  C100,98.7,79,91.7,62.4,86.4C51.5,82.9,39.8,77.7,40,72c0.1-4.7,6.2-9,18.4-8.4
  c8.2,0.4,15.4,1.1,29.7,8l14.1-24.5c-13.1-6.6-31.2-10.9-46-10.9h-0.1
  c-17.3,0-31.7,5.6-40.6,14.8C9,57.2,5.5,65.7,5.5,74.7C5.5,87.2,10.1,96,19.7,103
  c8.1,5.9,18.5,9.8,27.6,12.6c11.3,3.5,20.5,6.5,20.4,13c-0.1,2.4-1,4.7-2.7,6.4
  c-2.8,2.9-7.1,4-13.1,4.1c-11.5,0.2-20-1.6-33.6-9.6L5.8,154.4
  c14,8,29.9,12.2,46,12.2h2.1c14.2-0.2,25.7-4.3,34.9-11.7c0.5-0.4,1-0.8,1.5-1.3
  l-4.1,10.9H123l6.2-18.8c7,2.3,14.3,3.5,21.7,3.4c7.2,0,14.3-1.1,21.2-3.2
  l6,18.6h60.1v-39h13.1c31.7,0,50.5-16.2,50.5-43.2C301.7,52.2,283.5,38.4,244.7,38.4z
  M150.9,121c-4.4,0-8.8-0.7-13-2.3l12.9-40.6h0.2l12.6,40.7
  C159.6,120.3,155.2,121,150.9,121z M247.1,97.7h-8.9V64.9h8.9
  c11.9,0,21.4,4,21.4,16.1C268.5,93.7,259,97.6,247.1,97.7"/>
</svg>"""

# ── Global styles ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

  html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', '72', Arial, sans-serif;
  }

  /* Remove default Streamlit top padding */
  .block-container { padding-top: 0 !important; max-width: 780px; }

  /* Header bar */
  .sap-header {
    background: linear-gradient(135deg, #0070F2 0%, #0040B0 100%);
    padding: 24px 32px 20px 32px;
    margin: -1rem -1rem 0 -1rem;
    border-radius: 0;
    display: flex;
    align-items: center;
    gap: 18px;
  }
  .sap-header-text { color: #ffffff; }
  .sap-header-text h1 {
    margin: 0; font-size: 1.35rem; font-weight: 600; letter-spacing: -0.01em;
  }
  .sap-header-text p {
    margin: 2px 0 0 0; font-size: 0.82rem; opacity: 0.85; font-weight: 300;
  }

  /* Section card */
  .upload-card {
    background: #F5F6F7;
    border: 1px solid #EAECEE;
    border-radius: 10px;
    padding: 20px 24px;
    margin: 18px 0 8px 0;
  }
  .upload-card h3 {
    margin: 0 0 4px 0; font-size: 0.95rem; font-weight: 600; color: #32363A;
  }
  .upload-card p {
    margin: 0 0 14px 0; font-size: 0.8rem; color: #6A6D70;
  }

  /* File status badges */
  .badge-ok   { color: #107E3E; font-weight: 600; font-size: 0.82rem; }
  .badge-miss { color: #BB0000; font-size: 0.82rem; }

  /* Build button styling override */
  div[data-testid="stButton"] > button {
    background: #0070F2;
    color: white;
    border: none;
    border-radius: 6px;
    font-size: 1rem;
    font-weight: 600;
    padding: 0.6rem 2.4rem;
    width: 100%;
    cursor: pointer;
    transition: background 0.15s;
  }
  div[data-testid="stButton"] > button:hover {
    background: #0040B0;
  }
  div[data-testid="stButton"] > button:disabled {
    background: #EAECEE;
    color: #89919A;
    cursor: not-allowed;
  }

  /* Download button */
  div[data-testid="stDownloadButton"] > button {
    background: #107E3E;
    color: white;
    border: none;
    border-radius: 6px;
    font-size: 1rem;
    font-weight: 600;
    padding: 0.6rem 2.4rem;
    width: 100%;
  }
  div[data-testid="stDownloadButton"] > button:hover {
    background: #0B5C2C;
  }

  /* Success / error boxes */
  .result-ok {
    background: #F1FAF5; border: 1px solid #107E3E; border-radius: 8px;
    padding: 14px 20px; color: #0B5C2C; font-weight: 500; font-size: 0.9rem;
  }
  .result-err {
    background: #FFF3F3; border: 1px solid #BB0000; border-radius: 8px;
    padding: 14px 20px; color: #8B0000; font-weight: 500; font-size: 0.9rem;
  }

  /* Footer */
  .sap-footer {
    margin-top: 32px; padding-top: 16px; border-top: 1px solid #EAECEE;
    text-align: center; font-size: 0.75rem; color: #89919A;
  }

  /* Hide Streamlit branding */
  #MainMenu, footer { visibility: hidden; }
  header[data-testid="stHeader"] { display: none; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="sap-header">
  {SAP_SVG}
  <div class="sap-header-text">
    <h1>CBS Portfolio Dashboard Builder</h1>
    <p>Upload the four source Excel files to generate the interactive HTML dashboard.</p>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Instructions ──────────────────────────────────────────────────────────────
with st.expander("How to use this tool", expanded=False):
    st.markdown("""
**Three steps:**

1. Upload the four Excel files below — drag and drop or click to browse.
2. Click **Build Dashboard**.
3. Download both files — the full version with all data, and the public version without personal information (CBS Responsible, SO PM).

The dashboards are self-contained HTML files — anyone can open them in a browser with no login required.

---
**Required files**

| File | What it contains |
|---|---|
| `Services_Integrated_Project_Fi.xlsx` | Project financials, margins, EAC |
| `MANDI.xlsx` | Delivery status and LoS data |
| `Red Project Data Base.xlsx` | Weekly red project history |
| `Leakage Report.xlsx` | Leakage and billing type data |
""")

# ── File uploaders ─────────────────────────────────────────────────────────────
st.markdown('<div class="upload-card"><h3>Source Files</h3><p>All four files are required.</p></div>',
            unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    f_si = st.file_uploader(
        "Services Integrated",
        type=["xlsx"],
        key="si",
        help="Services_Integrated_Project_Fi.xlsx",
    )
    f_red = st.file_uploader(
        "Red Project Data Base",
        type=["xlsx"],
        key="red",
        help="Red Project Data Base.xlsx",
    )

with col2:
    f_mandi = st.file_uploader(
        "MANDI",
        type=["xlsx"],
        key="mandi",
        help="MANDI.xlsx",
    )
    f_leak = st.file_uploader(
        "Leakage Report",
        type=["xlsx"],
        key="leak",
        help="Leakage Report.xlsx",
    )

# ── Upload status summary ──────────────────────────────────────────────────────
files = {
    "Services Integrated":     f_si,
    "MANDI":                   f_mandi,
    "Red Project Data Base":   f_red,
    "Leakage Report":          f_leak,
}
n_ready = sum(1 for v in files.values() if v is not None)
all_ready = n_ready == 4

if n_ready > 0:
    status_parts = []
    for label, f in files.items():
        if f:
            status_parts.append(f'<span class="badge-ok">&#10003; {label}</span>')
        else:
            status_parts.append(f'<span class="badge-miss">&#9711; {label}</span>')
    st.markdown(
        "<div style='margin:8px 0 16px 0; display:flex; flex-wrap:wrap; gap:12px;'>"
        + "  ".join(status_parts)
        + "</div>",
        unsafe_allow_html=True,
    )

if not all_ready:
    remaining = 4 - n_ready
    st.caption(f"{remaining} file{'s' if remaining > 1 else ''} still needed.")

st.markdown("<br>", unsafe_allow_html=True)

# ── Build button ───────────────────────────────────────────────────────────────
build_clicked = st.button(
    "Build Dashboard",
    disabled=not all_ready,
    use_container_width=True,
)

# ── Build logic ────────────────────────────────────────────────────────────────
if build_clicked and all_ready:
    progress = st.progress(0, text="Starting build…")

    try:
        sys.path.insert(0, os.path.dirname(__file__))
        import build_core  # noqa: E402

        with tempfile.TemporaryDirectory() as tmp:
            def _save(uploaded, name):
                path = os.path.join(tmp, name)
                with open(path, "wb") as fh:
                    fh.write(uploaded.getbuffer())
                return path

            progress.progress(10, text="Saving uploaded files…")
            p_si   = _save(f_si,    "Services_Integrated.xlsx")
            p_ma   = _save(f_mandi, "MANDI.xlsx")
            p_red  = _save(f_red,   "Red_Project_Data_Base.xlsx")
            p_leak = _save(f_leak,  "Leakage_Report.xlsx")

            progress.progress(20, text="Building full dashboard…")
            html_full = build_core.build(p_si, p_ma, p_red, p_leak)

            progress.progress(80, text="Building public dashboard (wo personal data)…")
            html_public = build_core.build_public(p_si, p_ma, p_red, p_leak)

        progress.progress(100, text="Done.")
        st.session_state["html_full"]   = html_full
        st.session_state["html_public"] = html_public
        st.session_state["build_ok"]    = True

    except Exception as exc:
        progress.empty()
        st.markdown(
            f'<div class="result-err"><strong>Build failed.</strong><br>{exc}</div>',
            unsafe_allow_html=True,
        )
        st.session_state["build_ok"] = False

# ── Download section ───────────────────────────────────────────────────────────
if st.session_state.get("build_ok") and "html_full" in st.session_state:
    b_full   = st.session_state["html_full"].encode("utf-8")
    b_public = st.session_state["html_public"].encode("utf-8")

    st.markdown(
        f'<div class="result-ok">&#10003; Both dashboards built successfully &nbsp;|&nbsp; '
        f'{len(b_full)/1024:.0f} KB full &nbsp;/&nbsp; {len(b_public)/1024:.0f} KB public</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.download_button(
            label="Download — Full Dashboard",
            data=b_full,
            file_name="CBS_Portfolio_Dashboard.html",
            mime="text/html",
            use_container_width=True,
        )
    with col_b:
        st.download_button(
            label="Download — wo Personal Data",
            data=b_public,
            file_name="CBS_Portfolio_Dashboard_wo_Personal_Data.html",
            mime="text/html",
            use_container_width=True,
        )

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="sap-footer">CBS Portfolio Dashboard &nbsp;|&nbsp; SAP Delivery Operations</div>',
    unsafe_allow_html=True,
)
