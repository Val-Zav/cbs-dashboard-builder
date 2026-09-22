"""
CBS Oversight Tracker -- build_tracker_core.py
Callable module for Streamlit / programmatic use.

Exposes:
    build_tracker(si_src, m_src, r_src, l_src, ...) -> str (HTML)

Sources (si_src / m_src / r_src / l_src) can be:
  - bytes / bytearray      (from st.file_uploader().read())
  - file-like (BytesIO)    (from io.BytesIO(...))
  - str / Path             (local file path)

Baseline sources (baseline_*_src) are optional; when supplied the tracker
embeds week-over-week delta pills.
"""
import io
import json
import warnings
import pandas as pd
import numpy as np

warnings.filterwarnings('ignore')

# -- INLINE SAP LOGO (unmodified brand SVG) ------------------------------------
_SAP_SVG = (
    '<svg version="1.1" xmlns="http://www.w3.org/2000/svg" '
    'xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 412.4 204" '
    'style="height:40px;display:inline-block;vertical-align:middle;">'
    '<defs><linearGradient id="sgt" x1="206.19" y1="206" x2="206.19" y2="2" '
    'gradientUnits="userSpaceOnUse" gradientTransform="matrix(1 0 0 -1 0 206)">'
    '<stop offset="0" stop-color="#00B8F1"/>'
    '<stop offset="0.02" stop-color="#01B6F0"/>'
    '<stop offset="0.31" stop-color="#0D90D9"/>'
    '<stop offset="0.58" stop-color="#1775C8"/>'
    '<stop offset="0.82" stop-color="#1C65BF"/>'
    '<stop offset="1" stop-color="#1E5FBB"/>'
    '</linearGradient></defs>'
    '<polyline style="fill-rule:evenodd;clip-rule:evenodd;fill:url(#sgt)" '
    'points="0,204 208.4,204 412.4,0 0,0 0,204"/>'
    '<path style="fill-rule:evenodd;clip-rule:evenodd;fill:#FFFFFF" '
    'd="M244.7,38.4h-40.6v96.5l-35.5-96.6h-35.2l-30.3,80.7C100,98.7,79,91.7,62.4,86.4'
    'C51.5,82.9,39.8,77.7,40,72c0.1-4.7,6.2-9,18.4-8.4c8.2,0.4,15.4,1.1,29.7,8l14.1-24.5'
    'c-13.1-6.6-31.2-10.9-46-10.9h-0.1c-17.3,0-31.7,5.6-40.6,14.8c-6.2,6.3-9.7,14.8-9.7,23.7'
    'C5.5,87.2,10.1,96,19.7,103c8.1,5.9,18.5,9.8,27.6,12.6c11.3,3.5,20.5,6.5,20.4,13'
    'c-0.1,2.4-1,4.7-2.7,6.4c-2.8,2.9-7.1,4-13.1,4.1c-11.5,0.2-20-1.6-33.6-9.6L5.8,154.4'
    'c14,8,29.9,12.2,46,12.2h2.1c14.2-0.2,25.7-4.3,34.9-11.7c0.5-0.4,1-0.8,1.5-1.3l-4.1,10.9'
    'H123l6.2-18.8c7,2.3,14.3,3.5,21.7,3.4c7.2,0,14.3-1.1,21.2-3.2l6,18.6h60.1v-39h13.1'
    'c31.7,0,50.5-16.2,50.5-43.2C301.7,52.2,283.5,38.4,244.7,38.4z '
    'M150.9,121c-4.4,0-8.8-0.7-13-2.3l12.9-40.6h0.2l12.6,40.7C159.6,120.3,155.2,121,150.9,121z '
    'M247.1,97.7h-8.9V64.9h8.9c11.9,0,21.4,4,21.4,16.1C268.5,93.7,259,97.6,247.1,97.7"/>'
    '</svg>'
)

# -- COLUMN CONSTANTS ----------------------------------------------------------
_CBR = 'Contract Baseline Total Revenue before Re-allocation'
_REGULATED = {'A&D', 'Federal', 'State&Local'}

# CBS Team list (first-two-word keys, lowered, for fuzzy match after normalization)
_CBS_TEAM_NAMES = {
    'Fernanda Aviles', 'Tanya Cisneros', 'Santiago Ferro', 'Sofia Galvan',
    'Rebecka Jimenez', 'Salvador Lopez', 'Franco Ortelli', 'Jorge Padilla',
    'Rodrigo Reyes', 'Gerardo Ruiz', 'Almudena Sanz', 'Estefania Villanueva',
    'Valia Zavala',
}
# Build a set of (first, last) lower-cased keys for matching longer normalized names
_CBS_TEAM_KEYS = {
    tuple(n.lower().split()[:2]) for n in _CBS_TEAM_NAMES
}


def _is_cbs_team(name):
    """Return True if *name* belongs to the CBS analyst team."""
    if not name or str(name).strip() in ('--', '', 'nan', 'None', '-- not in MANDI --', '-unassigned-'):
        return False
    parts = str(name).strip().lower().split()
    if len(parts) < 2:
        return False
    return (parts[0], parts[1]) in _CBS_TEAM_KEYS


# =============================================================================
# SOURCE LOADER
# =============================================================================
def _load_src(src, sheet=0):
    """Load a DataFrame from bytes, BytesIO, file-like object, or path string."""
    if src is None:
        raise ValueError("Source file must not be None.")
    if isinstance(src, (bytes, bytearray)):
        return pd.read_excel(io.BytesIO(src), sheet_name=sheet)
    if hasattr(src, 'read'):
        # Rewind if seekable (handles multiple reads)
        if hasattr(src, 'seek'):
            src.seek(0)
        return pd.read_excel(src, sheet_name=sheet)
    return pd.read_excel(str(src), sheet_name=sheet)


# =============================================================================
# BUSINESS-DAY HELPERS
# =============================================================================
def _bd_range(year, month):
    start = pd.Timestamp(year, month, 1)
    end   = start + pd.offsets.MonthEnd(0)
    return pd.bdate_range(start, end)

def _bd_nth(year, month, n):
    days = _bd_range(year, month)
    return days[n - 1] if n <= len(days) else days[-1]

def _eom(year, month):
    return pd.Timestamp(year, month, 1) + pd.offsets.MonthEnd(0)

def _prev_month(ts):
    first = pd.Timestamp(ts.year, ts.month, 1)
    prev  = first - pd.DateOffset(months=1)
    return prev.year, prev.month

def _snapshot_status(snap_date, change_date, snap_type, cbr_val, today):
    """Return 'Updated' or 'Outdated' per BD-window rules."""
    if 'Project Baseline Update' in str(snap_type):
        return 'Outdated'
    if pd.notna(change_date) and pd.notna(snap_date) and change_date > snap_date:
        return 'Outdated'
    if pd.isna(snap_date):
        return 'Outdated'
    is_small = cbr_val < 1_500_000
    bd13_cur = _bd_nth(today.year, today.month, 13)
    if today >= bd13_cur:
        if is_small:
            ws = bd13_cur
            we = _eom(today.year, today.month)
        else:
            ws = _bd_nth(today.year, today.month, 1)
            we = _bd_nth(today.year, today.month, 9)
    else:
        if is_small:
            py, pm = _prev_month(today)
            ws = _bd_nth(py, pm, 13)
            we = _eom(py, pm)
        else:
            ws = _bd_nth(today.year, today.month, 1)
            we = _bd_nth(today.year, today.month, 9)
    snap_ts = pd.Timestamp(snap_date).normalize()
    return 'Updated' if ws.normalize() <= snap_ts <= we.normalize() else 'Outdated'


# =============================================================================
# FORMAT HELPERS
# =============================================================================
def _ss(v, n=60):
    if pd.isna(v) or str(v).strip() in ('nan', 'None', ''):
        return '--'
    s = str(v).strip()
    return s[:n] + '...' if len(s) > n else s

def _sf(v):
    try:
        f = float(v)
        return 0.0 if (np.isnan(f) or np.isinf(f)) else f
    except Exception:
        return 0.0

def _fmt_usd(v):
    v = _sf(v)
    if abs(v) >= 1_000_000:
        return f'{"-" if v < 0 else ""}${abs(v) / 1_000_000:.2f}M'
    if abs(v) >= 1_000:
        return f'{"-" if v < 0 else ""}${abs(v) / 1_000:.0f}K'
    return f'${v:,.0f}'

def _fmt_date(v):
    if pd.isna(v):
        return '--'
    try:
        return pd.Timestamp(v).strftime('%b %d, %Y')
    except Exception:
        return '--'


# =============================================================================
# NAME NORMALISATION
# =============================================================================
def _normalize_responsible(series):
    """Unify CBS Responsible name variants.

    Groups names by their first two words (case-insensitive) and maps every
    short variant to the longest form found in the data, e.g.:
        'Almudena Sanz' + 'Almudena Sanz Sakar'  ->  'Almudena Sanz Sakar'
    """
    _skip = {'--', '', '-- not in MANDI --', '-unassigned-', 'nan', 'None'}
    names = [
        str(n).strip()
        for n in series.dropna().unique()
        if str(n).strip() not in _skip
    ]
    # Group by first two words (lowered)
    groups = {}
    for name in names:
        parts = name.split()
        if len(parts) >= 2:
            key = (parts[0].lower(), parts[1].lower())
        else:
            continue
        groups.setdefault(key, []).append(name)
    # Canonical = longest variant per group
    alias = {}
    for _key, variants in groups.items():
        canonical = max(variants, key=len)
        for v in variants:
            if v != canonical:
                alias[v] = canonical
    if not alias:
        return series
    return series.map(lambda x: alias.get(str(x).strip(), x) if pd.notna(x) else x)


# =============================================================================
# DATA PIPELINE
# =============================================================================
def _run_pipeline(df_si, df_m, df_r, df_l, TODAY):
    """
    Full data pipeline.
    Returns (df, sec1, sec2, sec3, sec4, ALL_MUS, LATEST_RR_DATE).
    """
    # Active projects
    df = df_si.copy()
    df = df[df['Project Lifecycle Status'] != 'Delivery Completed'].copy()
    _end    = pd.to_datetime(df['Planned Finish Date'], errors='coerce')
    _cutoff = TODAY - pd.Timedelta(days=90)
    df = df[_end.isna() | (_end >= _cutoff)].copy()
    df['Bucket'] = df['Bucket'].apply(
        lambda x: 'Regulated Industries' if x in _REGULATED else x
    )
    ALL_MUS = sorted(df['Bucket'].dropna().unique().tolist())

    # MANDI join
    _mandi_cols = [c for c in ['ID', 'Responsible', 'SO PM', 'Overall Status', 'LoS']
                   if c in df_m.columns]
    mandi = df_m[_mandi_cols].copy()
    _col_map = {'ID': 'Project', 'Responsible': 'CBS_Responsible',
                'SO PM': 'SO_PM', 'Overall Status': 'MANDI_Status', 'LoS': 'LoS'}
    mandi.columns = [_col_map.get(c, c) for c in mandi.columns]
    mandi = mandi.drop_duplicates(subset='Project', keep='first')
    df = df.merge(mandi, on='Project', how='left')
    df['CBS_Responsible'] = df['CBS_Responsible'].fillna('-- not in MANDI --')
    df['CBS_Responsible'] = _normalize_responsible(df['CBS_Responsible'])
    df['MANDI_Status']    = df['MANDI_Status'].fillna('--')
    if 'SO_PM' not in df.columns:
        df['SO_PM'] = '--'
    df['Portfolio_Segment'] = df['Managed Portfolio'].map(
        {'Yes': 'Managed by CBS', 'No': 'SPOT Projects'}
    ).fillna('Other')
    df['_mp_val'] = df['Managed Portfolio'].map(
        {'Yes': 'Yes', 'No': 'No'}
    ).fillna('--')
    df['_team'] = df['CBS_Responsible'].apply(
        lambda x: 'CBS' if _is_cbs_team(x) else 'APM'
    )

    # Leakage join
    dl = df_l.copy()
    dl['BL']     = pd.to_numeric(dl.get('Backlog Leakage',   0), errors='coerce').fillna(0)
    dl['BL_pct'] = pd.to_numeric(dl.get('Backlog Leakage %', 0), errors='coerce').fillna(0)
    dl['CNV']    = pd.to_numeric(dl.get('Contract Net Value', 0), errors='coerce').fillna(0)
    dl['ANV']    = pd.to_numeric(dl.get('Adjusted Net Value', 0), errors='coerce').fillna(0)

    def _item_color(r):
        lk, cnv = r['BL'], r['CNV']
        if lk > 0:
            return 'Red'
        if lk < 0:
            pct = abs(lk) / abs(cnv) if cnv != 0 else 0
            return 'Orange' if pct > 0.05 else 'Green'
        return 'Green'

    def _worst(c):
        cl = list(c)
        return 'Red' if 'Red' in cl else ('Orange' if 'Orange' in cl else 'Green')

    dl['Item_Color'] = dl.apply(_item_color, axis=1)

    _agg_spec = {
        'Leak_Status':  ('Item_Color',             _worst),
        'Total_BL':     ('BL',                     'sum'),
        'Total_CNV':    ('CNV',                    'sum'),
        'Total_ANV':    ('ANV',                    'sum'),
        'Pos_BL':       ('BL',                     lambda x: x[x > 0].sum()),
        'Neg_BL':       ('BL',                     lambda x: x[x < 0].sum()),
        'Red_Items':    ('Item_Color',             lambda x: (x == 'Red').sum()),
        'Orange_Items': ('Item_Color',             lambda x: (x == 'Orange').sum()),
    }
    if 'Initial Contract Type' in dl.columns:
        _agg_spec['CT_mode'] = ('Initial Contract Type',
                                lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else '--')
    if 'Leakage reason text' in dl.columns:
        _agg_spec['Leak_Reason'] = ('Leakage reason text',
                                    lambda x: '; '.join(x.dropna().unique()[:2]))
    if 'Comment' in dl.columns:
        _agg_spec['Leak_Comment'] = ('Comment', lambda x: ('; '.join(
            v for v in x.dropna().unique()
            if str(v).strip() not in ('(No Value)', '', 'nan', 'None')
        ))[:120] or '--')
    if 'Backlog Leakage %' in dl.columns:
        _agg_spec['BL_pct_max'] = ('BL_pct', 'max')

    so_agg = dl.groupby('Sales Document Number').agg(**_agg_spec).reset_index()
    if 'CT_mode'      not in so_agg.columns: so_agg['CT_mode']      = '--'
    if 'Leak_Comment' not in so_agg.columns: so_agg['Leak_Comment'] = '--'

    so_agg['SO_num'] = pd.to_numeric(so_agg['Sales Document Number'], errors='coerce')
    df['SO_num']     = pd.to_numeric(df['Sales Order'], errors='coerce')
    df = df.merge(so_agg, on='SO_num', how='left')
    for col in ('Leak_Status', 'Total_BL', 'Total_CNV', 'Pos_BL', 'Neg_BL', 'Leak_Comment'):
        if col not in df.columns:
            df[col] = 0 if col not in ('Leak_Status', 'Leak_Comment') else ('N/A' if col == 'Leak_Status' else '--')
    df['Leak_Status'] = df['Leak_Status'].fillna('N/A')
    df['Total_BL']    = df['Total_BL'].fillna(0)
    df['Total_CNV']   = df['Total_CNV'].fillna(0)
    df['Pos_BL']      = df['Pos_BL'].fillna(0)
    df['Neg_BL']      = df['Neg_BL'].fillna(0)

    # Red Report join -- latest entry per project, current year
    df_r2 = df_r.copy()
    df_r2['_rev'] = pd.to_datetime(df_r2['Reviewed'], errors='coerce')
    df_r_yr = df_r2[df_r2['_rev'].dt.year == TODAY.year].copy()
    if df_r_yr.empty:
        df_r_yr = df_r2.copy()
    rr_latest = df_r_yr.sort_values('_rev').groupby('Project').last().reset_index()

    _rr_want = ['Project', '_rev', 'Intern Comments', 'Standard Comment', 'Fixable',
                'Analysis', 'FELIPE Last Changed Date', 'FELIPE Last Snapshot At Date',
                'FELIPE Last Snapshot Type', 'Overall Project Status', 'SAP Margin Status']
    rr_sub = rr_latest[[c for c in _rr_want if c in rr_latest.columns]].copy()
    rr_sub.columns = [
        'RR_' + c.replace(' ', '_') if c != 'Project' else c
        for c in rr_sub.columns
    ]
    df = df.merge(rr_sub, on='Project', how='left')

    LATEST_RR_DATE    = df_r2['_rev'].max()
    LATEST_RR_PROJECTS = set(df_r2[df_r2['_rev'] == LATEST_RR_DATE]['Project'].tolist())

    def _consec_weeks(project_id):
        dates = sorted(
            df_r2[df_r2['Project'] == project_id]['_rev'].dropna().dt.normalize().unique(),
            reverse=True,
        )
        count, expected = 0, LATEST_RR_DATE.normalize()
        for d in dates:
            d = pd.Timestamp(d)
            if abs((d - expected).days) <= 2:
                count  += 1
                expected -= pd.Timedelta(days=7)
            else:
                break
        return count

    _resolved = (
        df_r2[df_r2['Intern Comments'].str.strip().str.lower() == 'issue solved']
        .groupby('Project').size()
        .to_dict()
    )
    _consec_dict = {p: _consec_weeks(p) for p in LATEST_RR_PROJECTS}

    # FELIPE snapshot fields
    df['_felipe_change'] = pd.to_datetime(
        df.get('FELIPE Last Change Date',    pd.Series(dtype='datetime64[ns]')),
        errors='coerce',
    )
    df['_felipe_snap'] = pd.to_datetime(
        df.get('FELIPE Last Snapshot At Date', pd.Series(dtype='datetime64[ns]')),
        errors='coerce',
    )
    df['_snap_stale'] = (df['_felipe_change'] > df['_felipe_snap']) | df['_felipe_snap'].isna()

    # Section 1: Red status projects in most recent RR week
    _s1_base = df[(df['MANDI_Status'] == 'R') & df['Project'].isin(LATEST_RR_PROJECTS)].copy()
    _s1_base['_consec_weeks']   = _s1_base['Project'].map(_consec_dict).fillna(0).astype(int)
    _s1_base['_resolved_count'] = _s1_base['Project'].map(_resolved).fillna(0).astype(int)
    sec1 = _s1_base.sort_values('_consec_weeks', ascending=False)

    # Section 2: Positive leakage
    sec2 = df[df['Pos_BL'] > 0].sort_values('Pos_BL', ascending=False)

    # Section 3: High negative leakage (> 15 % of CNV)
    def _neg_leak_pct(row):
        cnv = abs(row['Total_CNV'])
        return abs(row['Total_BL']) / cnv if cnv > 0 else 0

    df['_neg_leak_pct'] = df.apply(_neg_leak_pct, axis=1)
    sec3 = df[(df['Total_BL'] < 0) & (df['_neg_leak_pct'] > 0.15)].sort_values(
        '_neg_leak_pct', ascending=False
    )

    # Section 4: Stale/missing FELIPE snapshot -- DIP only
    dip = df[df['Project Lifecycle Status'] == 'Delivery in Process'].copy()
    dip['_snap_status'] = dip.apply(
        lambda r: _snapshot_status(
            r['_felipe_snap'],
            r['_felipe_change'],
            r.get('FELIPE Last Snapshot Type', ''),
            _sf(r.get(_CBR, 0)),
            TODAY,
        ),
        axis=1,
    )
    sec4 = dip[dip['_snap_status'] == 'Outdated'].copy()
    sec4['_gap_days']  = (sec4['_felipe_change'] - sec4['_felipe_snap']).dt.days
    sec4['_snap_flag'] = sec4.apply(
        lambda r: 'No Snapshot' if pd.isna(r['_felipe_snap']) else 'Outdated', axis=1
    )
    sec4 = sec4.sort_values('_gap_days', ascending=False, na_position='first')

    return df, sec1, sec2, sec3, sec4, ALL_MUS, LATEST_RR_DATE


def _run_baseline_pipeline(bsi, bm, br, bl, ref_today):
    """Lightweight pipeline for baseline snapshot (delta pill computation)."""
    b = bsi.copy()
    b = b[b['Project Lifecycle Status'] != 'Delivery Completed'].copy()
    _end = pd.to_datetime(b['Planned Finish Date'], errors='coerce')
    b = b[_end.isna() | (_end >= ref_today - pd.Timedelta(days=90))].copy()
    b['Bucket'] = b['Bucket'].apply(
        lambda x: 'Regulated Industries' if x in _REGULATED else x
    )

    _mc = [c for c in ['ID', 'Responsible', 'Overall Status'] if c in bm.columns]
    mandi_b = bm[_mc].copy()
    _cm = {'ID': 'Project', 'Responsible': 'CBS_Responsible', 'Overall Status': 'MANDI_Status'}
    mandi_b.columns = [_cm.get(c, c) for c in mandi_b.columns]
    mandi_b = mandi_b.drop_duplicates(subset='Project', keep='first')
    b = b.merge(mandi_b, on='Project', how='left')
    b['MANDI_Status'] = b['MANDI_Status'].fillna('--')
    if 'CBS_Responsible' in b.columns:
        b['CBS_Responsible'] = _normalize_responsible(b['CBS_Responsible'])

    bl2 = bl.copy()
    bl2['BL']  = pd.to_numeric(bl2.get('Backlog Leakage',   0), errors='coerce').fillna(0)
    bl2['CNV'] = pd.to_numeric(bl2.get('Contract Net Value', 0), errors='coerce').fillna(0)
    so_agg_b = bl2.groupby('Sales Document Number').agg(
        Total_BL  = ('BL', 'sum'),
        Total_CNV = ('CNV', 'sum'),
        Pos_BL    = ('BL', lambda x: x[x > 0].sum()),
    ).reset_index()
    so_agg_b['SO_num'] = pd.to_numeric(so_agg_b['Sales Document Number'], errors='coerce')
    b['SO_num'] = pd.to_numeric(b['Sales Order'], errors='coerce')
    b = b.merge(so_agg_b, on='SO_num', how='left')
    b['Total_BL']  = b['Total_BL'].fillna(0)
    b['Total_CNV'] = b['Total_CNV'].fillna(0)
    b['Pos_BL']    = b['Pos_BL'].fillna(0)

    br2 = br.copy()
    br2['_rev'] = pd.to_datetime(br2['Reviewed'], errors='coerce')
    LATEST_RR_B   = br2['_rev'].max()
    LATEST_PROJ_B = set(br2[br2['_rev'] == LATEST_RR_B]['Project'].tolist())

    b['_felipe_change'] = pd.to_datetime(
        b.get('FELIPE Last Change Date',    pd.Series(dtype='datetime64[ns]')), errors='coerce'
    )
    b['_felipe_snap'] = pd.to_datetime(
        b.get('FELIPE Last Snapshot At Date', pd.Series(dtype='datetime64[ns]')), errors='coerce'
    )

    bs1 = b[(b['MANDI_Status'] == 'R') & b['Project'].isin(LATEST_PROJ_B)].copy()
    bs2 = b[b['Pos_BL'] > 0].copy()

    b['_nlp'] = b.apply(
        lambda r: abs(r['Total_BL']) / abs(r['Total_CNV']) if r['Total_CNV'] != 0 else 0, axis=1
    )
    bs3 = b[(b['Total_BL'] < 0) & (b['_nlp'] > 0.15)].copy()

    dip_b = b[b['Project Lifecycle Status'] == 'Delivery in Process'].copy()
    dip_b['_snap_status'] = dip_b.apply(
        lambda r: _snapshot_status(
            r['_felipe_snap'],
            r['_felipe_change'],
            r.get('FELIPE Last Snapshot Type', ''),
            _sf(r.get(_CBR, 0)),
            ref_today,
        ),
        axis=1,
    )
    bs4 = dip_b[dip_b['_snap_status'] == 'Outdated'].copy()

    return _compute_mu_counts(b, bs1, bs2, bs3, bs4)


def _compute_mu_counts(df_all, s1, s2, s3, s4):
    """Return {mu: {s1:{c,v}, s2:{c,v}, s3:{c,v}, s4:{c,v}, proj:int}} for JS."""
    result = {}
    for mu_key in ['all'] + sorted(df_all['Bucket'].dropna().unique().tolist()):
        d = {}
        for sec_id, sec_df, val_col in [
            ('s1', s1, None), ('s2', s2, 'Pos_BL'),
            ('s3', s3, 'Total_BL'), ('s4', s4, None),
        ]:
            rows = sec_df if mu_key == 'all' else sec_df[sec_df['Bucket'] == mu_key]
            c = int(len(rows))
            v = float(rows[val_col].sum()) if val_col and val_col in rows.columns else 0.0
            d[sec_id] = {'c': c, 'v': v}
        d['proj'] = int(
            len(df_all) if mu_key == 'all' else int((df_all['Bucket'] == mu_key).sum())
        )
        result[mu_key] = d
    return result



# =============================================================================
# SNAPSHOT HELPERS  (multi-week trend history)
# =============================================================================
_SNAP_VER  = 2
_MAX_WEEKS = 12


def _build_week_snap(date_str, df, sec1, sec2, sec3, sec4, mu_counts):
    """Create a compact snapshot dict for a single week."""
    s1_ids = set(sec1['Project'].tolist()) if len(sec1) else set()
    s2_ids = set(sec2['Project'].tolist()) if len(sec2) else set()
    s3_ids = set(sec3['Project'].tolist()) if len(sec3) else set()
    s4_ids = set(sec4['Project'].tolist()) if len(sec4) else set()
    projects = {}
    for _, row in df.iterrows():
        pid = str(row.get('Project', ''))
        if not pid or pid in ('nan', 'None', ''):
            continue
        secs = []
        if pid in s1_ids: secs.append(1)
        if pid in s2_ids: secs.append(2)
        if pid in s3_ids: secs.append(3)
        if pid in s4_ids: secs.append(4)
        projects[pid] = {'s': str(row.get('MANDI_Status', '--'))[:2], 'sec': secs}
    return {'date': date_str, 'mu_counts': mu_counts, 'projects': projects}


def _merge_snapshots(new_week, existing_src=None):
    """Merge *new_week* into existing snapshot history.
    Returns (merged_dict, merged_json_str).
    """
    if existing_src is not None:
        raw = existing_src
        if isinstance(raw, (bytes, bytearray)):
            raw = raw.decode('utf-8')
        if hasattr(raw, 'read'):
            if hasattr(raw, 'seek'):
                raw.seek(0)
            raw = raw.read()
            if isinstance(raw, (bytes, bytearray)):
                raw = raw.decode('utf-8')
        snap = json.loads(raw) if isinstance(raw, str) else raw
    else:
        snap = {'version': _SNAP_VER, 'weeks': []}

    weeks = [w for w in snap.get('weeks', []) if w.get('date') != new_week['date']]
    weeks.append(new_week)
    weeks.sort(key=lambda w: w['date'])
    if len(weeks) > _MAX_WEEKS:
        weeks = weeks[-_MAX_WEEKS:]
    merged = {'version': _SNAP_VER, 'weeks': weeks}
    return merged, json.dumps(merged, separators=(',', ':'))


def _trend_dots_kpi(weeks, key, sub_key='c', lower_is_better=True):
    """Return HTML trend-dot strip for a KPI across snapshot weeks."""
    if len(weeks) < 2:
        return ''
    vals = []
    for w in weeks:
        mc = w.get('mu_counts', {}).get('all', {})
        if sub_key:
            vals.append(mc.get(key, {}).get(sub_key, 0) if isinstance(mc.get(key), dict) else 0)
        else:
            vals.append(mc.get(key, 0))
    dots = []
    for i in range(1, len(vals)):
        diff = vals[i] - vals[i - 1]
        if abs(diff) < 0.005:
            cls = 'td-same'
        elif lower_is_better:
            cls = 'td-worse' if diff > 0 else 'td-better'
        else:
            cls = 'td-better' if diff > 0 else 'td-worse'
        dots.append(f'<span class="trend-dot {cls}" title="w{i}"></span>')
    return '<span class="trend-dots">' + ''.join(dots) + '</span>' if dots else ''


def _trend_dots_project(history_weeks, project_id, section_num):
    """Return HTML dots showing per-project section membership over historical weeks.
    *history_weeks* = [(date_str, {pid: {s, sec}}), ...] -- excludes current week.
    """
    if not history_weeks:
        return '<span class="trend-na">--</span>'
    pid = str(project_id)
    dots = []
    for date, projs in history_weeks:
        p = projs.get(pid)
        in_sec = p is not None and section_num in p.get('sec', [])
        cls = 'td-in' if in_sec else 'td-out'
        dots.append(f'<span class="trend-dot {cls}" title="{date}"></span>')
    return '<span class="trend-dots">' + ''.join(dots) + '</span>' if dots else '<span class="trend-na">--</span>'

# =============================================================================
# MAIN PUBLIC FUNCTION
# =============================================================================
def build_tracker(
    si_src,
    m_src,
    r_src,
    l_src,
    baseline_si_src=None,
    baseline_m_src=None,
    baseline_r_src=None,
    baseline_l_src=None,
    today_date=None,
    baseline_date=None,
    snapshot_src=None,
):
    """
    Build the CBS Oversight Tracker HTML report.

    Parameters
    ----------
    si_src, m_src, r_src, l_src
        Current-week source files (bytes, BytesIO, or path).
    baseline_si_src ... baseline_l_src
        Previous-week source files (optional; enables delta pills).
    today_date
        datetime.date, str 'YYYY-MM-DD', or None (defaults to today).
    baseline_date
        str 'YYYY-MM-DD' for the baseline week label (optional).

    snapshot_src
        Snapshot JSON from a previous build (bytes, file-like, or str).
        When supplied, provides multi-week trend data and replaces the
        need for baseline files.

    Returns
    -------
    (str, str)
        Tuple of (HTML document, snapshot JSON for next week).
    """
    from datetime import date as _date

    # -- 1. TODAY --------------------------------------------------------------
    if today_date is None:
        today_date = _date.today()
    TODAY     = pd.Timestamp(today_date)
    TODAY_STR = TODAY.strftime('%B %-d, %Y')
    CBR       = _CBR

    # -- 2. Load current files -------------------------------------------------
    df_si = _load_src(si_src)
    df_m  = _load_src(m_src)
    df_r  = _load_src(r_src, sheet=0)
    df_l  = _load_src(l_src)

    # -- 3. Pipeline -----------------------------------------------------------
    df, sec1, sec2, sec3, sec4, ALL_MUS, LATEST_RR_DATE = _run_pipeline(
        df_si, df_m, df_r, df_l, TODAY
    )

    # -- 3b. Current MU counts (for snapshot) --------------------------------
    _cur_mu_counts = _compute_mu_counts(df, sec1, sec2, sec3, sec4)

    # -- 4. KPIs ---------------------------------------------------------------
    total_projects     = len(df)
    red_count          = int((df['MANDI_Status'] == 'R').sum())
    yellow_count       = int((df['MANDI_Status'] == 'Y').sum())
    pos_leak_count     = len(sec2)
    neg_leak_count     = len(sec3)
    stale_snap_count   = len(sec4)
    total_pos_leak_val = float(sec2['Pos_BL'].sum())
    total_neg_leak_val = float(sec3['Total_BL'].sum())

    # -- 5. Snapshot + baseline ------------------------------------------------
    _this_snap = _build_week_snap(
        TODAY.strftime('%Y-%m-%d'), df, sec1, sec2, sec3, sec4, _cur_mu_counts,
    )
    _snap_dict, _snap_json_out = _merge_snapshots(_this_snap, snapshot_src)
    _snap_weeks = _snap_dict.get('weeks', [])

    # History = all weeks except the current one (for project trend dots)
    _history_weeks = [
        (w['date'], w.get('projects', {}))
        for w in _snap_weeks if w['date'] != TODAY.strftime('%Y-%m-%d')
    ]

    # Delta source: prefer snapshot (2nd-to-last week), fall back to legacy baseline
    _has_baseline = (baseline_si_src is not None and
                     all(x is not None for x in [baseline_m_src, baseline_r_src, baseline_l_src]))
    if len(_snap_weeks) >= 2:
        _prev_snap_week = _snap_weeks[-2]
        _PREV_MU_RAW   = _prev_snap_week.get('mu_counts', {})
        _prev_d = pd.Timestamp(_prev_snap_week['date'])
        PREV_DATE_STR = _prev_d.strftime('%b %-d')
    elif _has_baseline:
        b_today = pd.Timestamp(baseline_date) if baseline_date else TODAY - pd.Timedelta(days=7)
        bsi = _load_src(baseline_si_src)
        bm  = _load_src(baseline_m_src)
        br  = _load_src(baseline_r_src, sheet=0)
        bl  = _load_src(baseline_l_src)
        _PREV_MU_RAW = _run_baseline_pipeline(bsi, bm, br, bl, b_today)
        PREV_DATE_STR = b_today.strftime('%b %-d')
    else:
        _PREV_MU_RAW  = {}
        PREV_DATE_STR = ''

    PREV_MU_JSON = json.dumps(_PREV_MU_RAW, separators=(',', ':'))

    # KPI-level trend dots (from full snapshot history)
    _td_proj  = _trend_dots_kpi(_snap_weeks, 'proj', sub_key=None, lower_is_better=False)
    _td_red   = _trend_dots_kpi(_snap_weeks, 's1', 'c', True)
    _td_pos_n = _trend_dots_kpi(_snap_weeks, 's2', 'c', True)
    _td_pos_v = _trend_dots_kpi(_snap_weeks, 's2', 'v', True)
    _td_neg_n = _trend_dots_kpi(_snap_weeks, 's3', 'c', True)
    _td_neg_v = _trend_dots_kpi(_snap_weeks, 's3', 'v', True)
    _td_stale = _trend_dots_kpi(_snap_weeks, 's4', 'c', True)

    # -- 6. Delta pills --------------------------------------------------------
    def _dpill(curr, prev, lower_is_better=True, fmt='int', use_abs=False, pid=''):
        if not PREV_DATE_STR or prev is None:
            return ''
        c    = abs(curr) if use_abs else curr
        diff = c - prev
        if abs(diff) < 0.005:
            return (f'<span class="delta-kpi delta-nc" id="dp-{pid}">'
                    f'&#8212;&thinsp;vs&nbsp;{PREV_DATE_STR}</span>')
        if lower_is_better is None:
            cls = 'delta-nc'
        else:
            cls = 'delta-worse' if (diff > 0 if lower_is_better else diff < 0) else 'delta-better'
        arrow = '&#9650;' if diff > 0 else '&#9660;'
        if fmt == 'usd':
            ad   = abs(diff)
            sign = '+' if diff > 0 else '&#8722;'
            ds   = (f'${ad / 1e6:.2f}M' if ad >= 1e6
                    else f'${ad / 1e3:.0f}K' if ad >= 1e3
                    else f'${ad:.0f}')
            ds = sign + ds
        else:
            sign = '+' if diff > 0 else ''
            ds   = f'{sign}{abs(int(round(diff)))}'
        return (f'<span class="delta-kpi {cls}" id="dp-{pid}">'
                f'{arrow}&thinsp;{ds}&nbsp;vs&nbsp;{PREV_DATE_STR}</span>')

    _all_prev = _PREV_MU_RAW.get('all', {}) if _PREV_MU_RAW else {}
    _pp = lambda sec, k: (
        (_all_prev.get(sec) or {}).get(k, None)
        if isinstance(_all_prev.get(sec), dict) else None
    )
    _proj_prev = (_all_prev.get('proj')
                  if isinstance(_all_prev.get('proj'), (int, float)) else None)

    d_proj  = _dpill(total_projects,     _proj_prev,       None, 'int', False, 'proj')
    d_red   = _dpill(len(sec1),          _pp('s1', 'c'),   True, 'int', False, 'red')
    d_yel   = _dpill(yellow_count,       None,             True, 'int', False, 'yel')
    d_pos_n = _dpill(pos_leak_count,     _pp('s2', 'c'),   True, 'int', False, 'pos-n')
    d_pos_v = _dpill(total_pos_leak_val, _pp('s2', 'v'),   True, 'usd', False, 'pos-v')
    d_neg_n = _dpill(neg_leak_count,     _pp('s3', 'c'),   True, 'int', False, 'neg-n')
    d_neg_v = _dpill(total_neg_leak_val, _pp('s3', 'v'),   True, 'usd', True,  'neg-v')
    d_stale = _dpill(stale_snap_count,   _pp('s4', 'c'),   True, 'int', False, 'stale')

    # -- 7. HTML builders (closures over TODAY, CBR) ---------------------------
    def badge(text, color):
        _colors = {
            'R': '#BB0000', 'Y': '#E8B400', 'G': '#107F3E',
            'Orange': '#E56000', 'N/A': '#8396A8', '--': '#8396A8', 'blue': '#0070F2',
        }
        bg = _colors.get(color, '#8396A8')
        return f'<span class="badge" style="background:{bg}">{text}</span>'

    def status_badge(s):
        lbl = {'R': 'Red', 'Y': 'Yellow', 'G': 'Green'}.get(s, s)
        return badge(lbl, s)

    def fixable_badge(s):
        if pd.isna(s) or s == '--':
            return badge('Unknown', 'N/A')
        s = str(s).strip()
        if s == 'Fixed':               return badge('Fixed', 'G')
        if 'Working with' in s:        return badge(s, 'Y')
        if 'Not Fixable' in s:         return badge('Not Fixable', 'R')
        if 'Not Found' in s:           return badge('Not Found', 'N/A')
        if 'Rebaseline' in s:          return badge('Rebaseline in Process', 'blue')
        if 'ICO' in s or 'CR in' in s: return badge(s, 'blue')
        return f'<span class="badge badge-gray">{s}</span>'

    def table_row(*cells, cls='', mu='', seg='', resp='', size='', lc='', numval='', team=''):
        da = (f' data-mu="{mu}" data-seg="{seg}" data-resp="{resp}"'
              f' data-size="{size}" data-lc="{lc}" data-team="{team}"') if mu else ''
        if numval != '':
            da += f' data-numval="{numval}"'
        tds = ''.join(f'<td>{c}</td>' for c in cells)
        return f'<tr class="{cls}"{da}>{tds}</tr>'

    def section_header(num, title, count, color, description, questions):
        qs = ''.join(f'<li>{q}</li>' for q in questions)
        _cc = {'R': '#BB0000', 'Y': '#E8B400', 'G': '#107F3E',
               'orange': '#E56000', 'blue': '#0070F2'}.get(color, '#0070F2')
        return f'''
<div class="section" id="s{num}">
  <div class="section-hdr" style="border-left:4px solid {_cc}">
    <div class="section-title-row">
      <span class="section-num">{num}</span>
      <h2 class="section-title">{title}</h2>
      <span class="section-count" style="background:{_cc}">
        <span class="vis-count" id="vc{num}">{count}</span> / {count}
      </span>
      <div style="flex:1"></div>
    </div>
    <p class="section-desc">{description}</p>
    <div class="questions">
      <span class="q-label">Preguntas clave para la llamada</span>
      <ul>{qs}</ul>
    </div>
  </div>'''

    def s1_row(row):
        fix_val  = row.get('RR_Fixable', '--')
        sc_val   = row.get('RR_Standard_Comment', '--')
        sms      = _ss(row.get('SAP Margin Status', '--'), 20)
        eac_m    = _sf(row.get('EAC Margin', 0))
        so_val   = _ss(str(int(row['Sales Order'])) if pd.notna(row.get('Sales Order')) else '--')
        sopm_val = _ss(row.get('SO_PM', '--'), 25)
        mu_val   = _ss(row.get('Bucket', ''))
        seg_val  = _ss(row.get('Portfolio_Segment', ''))
        resp_val = _ss(row.get('CBS_Responsible', '--'))
        mp_val   = _ss(row.get('_mp_val', '--'))
        team_val = row.get('_team', 'APM')
        size_val = 'lt1.5m' if _sf(row.get(CBR, 0)) < 1_500_000 else 'ge1.5m'
        lc_val   = _ss(row.get('Project Lifecycle Status', ''))
        cls      = ('row-red' if row.get('MANDI_Status') == 'R'
                    else 'row-yellow' if row.get('MANDI_Status') == 'Y' else '')
        consec   = int(row.get('_consec_weeks', 0))
        resolved = int(row.get('_resolved_count', 0))
        cb = f'<span class="badge" style="background:#BB0000;font-size:13px">{consec}w</span>'
        rc = (f'<span style="font-weight:700;color:'
              f'{"#107F3E" if resolved > 0 else "#8396A8"}">{resolved}</span>')
        return table_row(
            f'<code class="proj-id">{_ss(row["Project"], 20)}</code>',
            so_val, sopm_val,
            f'<span class="proj-desc" title="{_ss(row.get("Project Description", ""), 100)}">'
            f'{_ss(row.get("Project Description", ""), 35)}</span>',
            _ss(row.get('Customer', '--'), 30),
            mu_val, mp_val, resp_val,
            status_badge(row.get('MANDI_Status', '--')),
            f'<span class="sms-{sms.lower()}">{sms}</span>',
            _fmt_usd(row.get(CBR, 0)),
            f'{eac_m * 100:.1f}%',
            fixable_badge(fix_val),
            _ss(sc_val, 35),
            cb, rc,
            _trend_dots_project(_history_weeks, row['Project'], 1),
            cls=cls, mu=mu_val, seg=seg_val, resp=resp_val, size=size_val, lc=lc_val, team=team_val,
        )

    def s2_row(row):
        ct       = str(row.get('CT_mode', '--')).replace('Mix', 'Mixed')
        pct      = row['Pos_BL'] / abs(row['Total_CNV']) * 100 if row['Total_CNV'] != 0 else 0
        so_val   = _ss(str(int(row['Sales Order'])) if pd.notna(row.get('Sales Order')) else '--')
        sopm_val = _ss(row.get('SO_PM', '--'), 25)
        mu_val   = _ss(row.get('Bucket', ''))
        seg_val  = _ss(row.get('Portfolio_Segment', ''))
        resp_val = _ss(row.get('CBS_Responsible', '--'))
        mp_val   = _ss(row.get('_mp_val', '--'))
        team_val = row.get('_team', 'APM')
        size_val = 'lt1.5m' if _sf(row.get(CBR, 0)) < 1_500_000 else 'ge1.5m'
        lc_val   = _ss(row.get('Project Lifecycle Status', ''))
        return table_row(
            f'<code class="proj-id">{_ss(row["Project"], 20)}</code>',
            so_val, sopm_val,
            f'<span class="proj-desc" title="{_ss(row.get("Project Description", ""), 100)}">'
            f'{_ss(row.get("Project Description", ""), 35)}</span>',
            _ss(row.get('Customer', '--'), 30),
            mu_val, mp_val, resp_val,
            badge(ct, 'blue'),
            _fmt_usd(row['Total_CNV']),
            f'<span class="leak-pos">{_fmt_usd(row["Pos_BL"])}</span>',
            f'<span class="leak-pos">{pct:.1f}%</span>',
            _ss(row.get('Leak_Comment', '--'), 50),
            _fmt_date(row.get('_felipe_snap', None)),
            _trend_dots_project(_history_weeks, row['Project'], 2),
            mu=mu_val, seg=seg_val, resp=resp_val, size=size_val, lc=lc_val, team=team_val,
            numval=round(float(row['Pos_BL']), 2),
        )

    def s3_row(row):
        ct       = str(row.get('CT_mode', '--')).replace('Mix', 'Mixed')
        pct      = row['_neg_leak_pct'] * 100
        so_val   = _ss(str(int(row['Sales Order'])) if pd.notna(row.get('Sales Order')) else '--')
        sopm_val = _ss(row.get('SO_PM', '--'), 25)
        mu_val   = _ss(row.get('Bucket', ''))
        seg_val  = _ss(row.get('Portfolio_Segment', ''))
        resp_val = _ss(row.get('CBS_Responsible', '--'))
        mp_val   = _ss(row.get('_mp_val', '--'))
        team_val = row.get('_team', 'APM')
        size_val = 'lt1.5m' if _sf(row.get(CBR, 0)) < 1_500_000 else 'ge1.5m'
        lc_val   = _ss(row.get('Project Lifecycle Status', ''))
        return table_row(
            f'<code class="proj-id">{_ss(row["Project"], 20)}</code>',
            so_val, sopm_val,
            f'<span class="proj-desc" title="{_ss(row.get("Project Description", ""), 100)}">'
            f'{_ss(row.get("Project Description", ""), 35)}</span>',
            _ss(row.get('Customer', '--'), 30),
            mu_val, mp_val, resp_val,
            badge(ct, 'blue'),
            _fmt_usd(row['Total_CNV']),
            f'<span class="leak-neg">{_fmt_usd(row["Total_BL"])}</span>',
            f'<span class="leak-neg">-{pct:.1f}%</span>',
            _ss(row.get('Leak_Comment', '--'), 50),
            _fmt_date(row.get('_felipe_snap', None)),
            _trend_dots_project(_history_weeks, row['Project'], 3),
            mu=mu_val, seg=seg_val, resp=resp_val, size=size_val, lc=lc_val, team=team_val,
            numval=round(float(row['Total_BL']), 2),
        )

    def s4_row(row):
        chg       = row.get('_felipe_change', None)
        snap      = row.get('_felipe_snap', None)
        snap_type = _ss(row.get('FELIPE Last Snapshot Type', '--'), 30)
        gap_days  = row.get('_gap_days', None)
        gap_str   = f'{int(gap_days)} days' if pd.notna(gap_days) else '-- no snapshot'
        gap_cls   = 'gap-high' if (pd.notna(gap_days) and gap_days > 14) or pd.isna(gap_days) else 'gap-med'
        st_val    = row.get('_snap_status', 'Outdated')
        st_badge  = (
            '<span class="badge" style="background:#107F3E">Updated</span>'
            if st_val == 'Updated'
            else '<span class="badge" style="background:#BB0000">Outdated</span>'
        )
        end_date = row.get('Planned Finish Date', None)
        end_ts   = pd.to_datetime(end_date, errors='coerce')
        past_end = pd.notna(end_ts) and end_ts < TODAY
        so_val   = _ss(str(int(row['Sales Order'])) if pd.notna(row.get('Sales Order')) else '--')
        sopm_val = _ss(row.get('SO_PM', '--'), 25)
        mu_val   = _ss(row.get('Bucket', ''))
        seg_val  = _ss(row.get('Portfolio_Segment', ''))
        resp_val = _ss(row.get('CBS_Responsible', '--'))
        mp_val   = _ss(row.get('_mp_val', '--'))
        team_val = row.get('_team', 'APM')
        size_val = 'lt1.5m' if _sf(row.get(CBR, 0)) < 1_500_000 else 'ge1.5m'
        lc_val   = _ss(row.get('Project Lifecycle Status', ''))
        cls      = 'row-past-end' if past_end else ''
        return table_row(
            f'<code class="proj-id">{_ss(row["Project"], 20)}</code>',
            so_val, sopm_val,
            _ss(row.get('Customer', '--'), 30),
            mu_val, mp_val, lc_val, resp_val,
            _fmt_date(end_date),
            st_badge,
            _fmt_date(chg),
            _fmt_date(snap),
            f'<span class="{gap_cls}">{gap_str}</span>',
            snap_type,
            status_badge(row.get('MANDI_Status', '--')),
            _trend_dots_project(_history_weeks, row['Project'], 4),
            cls=cls, mu=mu_val, seg=seg_val, resp=resp_val, size=size_val, lc=lc_val, team=team_val,
        )

    # -- 8. UI fragments -------------------------------------------------------
    def _mu_id(mu):
        return mu.replace(' ', '-').replace('&', '').replace('/', '-')

    mu_buttons = '\n'.join(
        f'    <button class="filter-btn" id="mu-{_mu_id(mu)}" '
        f'onclick="setFilter(\'mu\',\'{mu}\',this)">{mu}</button>'
        for mu in ALL_MUS
    )

    _raw_resp = df['CBS_Responsible'].dropna().unique().tolist()
    _skip_r   = {'--', '', '-unassigned-', '-- not in MANDI --', 'nan', 'None'}
    all_resp  = sorted({_ss(r) for r in _raw_resp if _ss(r) not in _skip_r})
    resp_options = '\n'.join(f'<option value="{r}">{r}</option>' for r in all_resp)
    if df['CBS_Responsible'].str.strip().isin(['-- not in MANDI --']).any():
        resp_options += '\n<option value="-- not in MANDI --">-- Not in MANDI</option>'

    # -- 9. HTML ---------------------------------------------------------------
    html = rf'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CBS Oversight Tracker -- All Market Units | {TODAY_STR}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"72 Brand","72",Arial,sans-serif;font-size:13px;color:#1D2329;background:#F5F6F7;line-height:1.5}}
.header{{background:#fff;border-bottom:2px solid #0070F2;padding:14px 32px;display:flex;align-items:center;justify-content:space-between;gap:20px;position:sticky;top:0;z-index:100;box-shadow:0 2px 10px rgba(0,0,0,.08)}}
.header-brand{{display:flex;align-items:center;gap:16px;flex:1;min-width:0}}
.header-logo{{height:40px;flex-shrink:0}}
.header-sep{{width:1px;height:40px;background:#D8DCE0;flex-shrink:0}}
.header-info h1{{font-size:19px;font-weight:700;color:#1D2329;white-space:nowrap}}
.header-info p{{font-size:11px;color:#8396A8;margin-top:3px}}
.header-actions{{display:flex;align-items:center;gap:16px;flex-shrink:0}}
.header-date{{text-align:right;white-space:nowrap}}
.header-date-label{{display:block;font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:#8396A8;margin-bottom:2px}}
.header-date strong{{display:block;font-size:15px;font-weight:700;color:#1D2329}}
.kpi-bar{{background:#F5F6F8;border-bottom:1px solid #E1E2E6;padding:20px 32px;display:flex;gap:10px;flex-wrap:wrap;align-items:stretch;justify-content:center}}
.kpi-card{{background:#fff;border-radius:10px;box-shadow:0 1px 5px rgba(0,0,0,.07);padding:14px 18px 12px;min-width:110px;display:flex;flex-direction:column;border-top:3px solid #E1E2E6;transition:box-shadow .15s,transform .15s;cursor:default}}
.kpi-card:hover{{box-shadow:0 4px 14px rgba(0,0,0,.11);transform:translateY(-2px)}}
.kpi-card.kpi-blue{{border-top-color:#0070F2}}
.kpi-card.kpi-red{{border-top-color:#BB0000}}
.kpi-card.kpi-amber{{border-top-color:#C87400}}
.kpi-card.kpi-orange{{border-top-color:#E56000}}
.kpi-card.kpi-green{{border-top-color:#107F3E}}
.kpi-group-lbl{{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:#A9B4BC;margin-bottom:5px}}
.kpi-val{{font-size:26px;font-weight:700;line-height:1;color:#1D2329;transition:color .2s}}
.kpi-val-md{{font-size:20px;font-weight:700;line-height:1.1;color:#1D2329;transition:color .2s}}
.kpi-card.kpi-blue .kpi-val,.kpi-card.kpi-blue .kpi-val-md{{color:#0070F2}}
.kpi-card.kpi-red .kpi-val,.kpi-card.kpi-red .kpi-val-md{{color:#BB0000}}
.kpi-card.kpi-amber .kpi-val,.kpi-card.kpi-amber .kpi-val-md{{color:#C87400}}
.kpi-card.kpi-orange .kpi-val,.kpi-card.kpi-orange .kpi-val-md{{color:#E56000}}
.kpi-card.kpi-green .kpi-val,.kpi-card.kpi-green .kpi-val-md{{color:#107F3E}}
.kpi-lbl{{font-size:10.5px;color:#5D6A73;margin-top:5px;line-height:1.35}}
.delta-kpi{{display:inline-block;font-size:10px;font-weight:700;margin-top:4px;padding:2px 6px;border-radius:3px;white-space:nowrap;letter-spacing:.01em}}
.delta-better{{background:#E6F4EA;color:#107F3E}}
.delta-worse{{background:#FDECEA;color:#BB0000}}
.delta-nc{{background:#F0F1F2;color:#8396A8}}
.kpi-sep{{width:1px;background:#E1E2E6;align-self:stretch;margin:0 4px;flex-shrink:0}}
.nav{{background:#fff;border-bottom:2px solid #E1E2E6;padding:0 32px;display:flex;overflow-x:auto;flex-shrink:0}}
.nav a{{display:inline-flex;align-items:center;gap:8px;padding:14px 24px;font-size:12.5px;font-weight:600;color:#5D6A73;text-decoration:none;border-bottom:3px solid transparent;margin-bottom:-2px;white-space:nowrap;transition:color .15s,background .15s,border-color .15s}}
.nav a:hover{{background:#F5F6F8}}
.nav a.s1{{color:#BB0000}}.nav a.s1:hover,.nav a.s1.spy-active{{background:#FFF5F5;border-bottom-color:#BB0000}}
.nav a.s2{{color:#E56000}}.nav a.s2:hover,.nav a.s2.spy-active{{background:#FFF8F0;border-bottom-color:#E56000}}
.nav a.s3{{color:#C87400}}.nav a.s3:hover,.nav a.s3.spy-active{{background:#FFFBF0;border-bottom-color:#C87400}}
.nav a.s4{{color:#0070F2}}.nav a.s4:hover,.nav a.s4.spy-active{{background:#EEF6FF;border-bottom-color:#0070F2}}
.nav-badge{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;color:inherit;transition:background .15s}}
.nav a.s1 .nav-badge{{background:rgba(187,0,0,.1)}}
.nav a.s2 .nav-badge{{background:rgba(229,96,0,.1)}}
.nav a.s3 .nav-badge{{background:rgba(200,116,0,.1)}}
.nav a.s4 .nav-badge{{background:rgba(0,112,242,.1)}}
.content{{width:100%;padding:24px 32px;box-sizing:border-box}}
.section{{background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.06),0 0 0 1px rgba(0,0,0,.05);margin-bottom:28px;overflow:visible;width:100%;transition:box-shadow .2s}}
.section:hover{{box-shadow:0 4px 16px rgba(0,0,0,.09),0 0 0 1px rgba(0,0,0,.05)}}
.section-hdr{{padding:18px 22px;background:#FAFBFC;border-bottom:1px solid #E1E2E6;border-radius:10px 10px 0 0}}
.section-title-row{{display:flex;align-items:center;gap:12px;margin-bottom:8px;flex-wrap:wrap}}
.section-num{{width:28px;height:28px;border-radius:50%;background:#0070F2;color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;box-shadow:0 2px 6px rgba(0,112,242,.35)}}
.section-title{{font-size:15px;font-weight:700;flex:1;min-width:0;letter-spacing:-.01em}}
.section-count{{padding:3px 11px;border-radius:12px;color:#fff;font-size:12px;font-weight:700;white-space:nowrap;letter-spacing:.01em}}
.section-desc{{font-size:12px;color:#5D6A73;margin-bottom:10px;line-height:1.5}}
.questions{{background:linear-gradient(135deg,#EEF6FF,#F5F9FF);border:1px solid #CBE0FF;border-radius:8px;padding:11px 16px}}
.q-label{{font-size:9.5px;font-weight:700;color:#0070F2;text-transform:uppercase;letter-spacing:.6px;display:block;margin-bottom:6px}}
.questions ul{{padding-left:18px}}
.questions li{{font-size:12px;color:#1D2329;margin-bottom:4px;line-height:1.5}}
.table-wrap{{overflow-x:auto;-webkit-overflow-scrolling:touch;border-radius:0 0 10px 10px;border-top:1px solid #E1E2E6}}
table{{width:100%;border-collapse:collapse;min-width:700px}}
th{{background:#F3F4F6;font-size:10.5px;font-weight:700;color:#6B7A86;text-transform:uppercase;letter-spacing:.4px;padding:9px 11px;text-align:left;white-space:nowrap;border-bottom:2px solid #D8DCE0;position:sticky;top:0;z-index:1}}
td{{padding:8px 11px;border-bottom:1px solid #F0F2F5;vertical-align:middle;font-size:12px;line-height:1.4}}
tr:hover td{{background:#F4F7FC}}
tr:last-child td{{border-bottom:none}}
.row-red td{{background:#FFF5F5}}
.row-red:hover td{{background:#FFECEC}}
.row-past-end td{{background:#FEF0EF}}
.row-past-end:hover td{{background:#FAD7D3}}
.row-yellow td{{background:#FFFCF0}}
.row-yellow:hover td{{background:#FFF7E0}}
.badge{{display:inline-block;padding:2px 8px;border-radius:10px;color:#fff;font-size:10px;font-weight:700;white-space:nowrap}}
.badge-gray{{background:#8396A8;color:#fff;display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700}}
.proj-id{{font-size:11px;font-family:monospace;color:#0034A0}}
.proj-desc{{color:#1D2329}}
code.proj-id{{background:#EEF2FF;padding:2px 5px;border-radius:3px;font-size:11px;color:#0034A0}}
.leak-pos{{color:#BB0000;font-weight:700}}
.leak-neg{{color:#C87400;font-weight:700}}
.sms-red{{color:#BB0000;font-weight:700}}
.sms-yellow{{color:#C87400;font-weight:700}}
.sms-green{{color:#107F3E;font-weight:700}}
.gap-high{{color:#BB0000;font-weight:700}}
.gap-med{{color:#C87400;font-weight:700}}
.empty-state{{padding:32px;text-align:center;color:#8396A8;font-size:13px}}
.footer{{padding:16px 32px;text-align:center;color:#8396A8;font-size:11px;border-top:1px solid #E1E2E6;margin-top:8px}}
.filter-bar{{background:#fff;border-bottom:1px solid #E1E2E6;padding:9px 32px;display:flex;align-items:center;gap:0;flex-wrap:wrap;position:sticky;top:68px;z-index:90;box-shadow:0 2px 6px rgba(0,0,0,.05)}}
.filter-bar-2{{background:#FAFBFC;border-bottom:1px solid #E1E2E6;padding:7px 32px;display:flex;align-items:center;gap:0;flex-wrap:wrap}}
.filter-group{{display:flex;align-items:center;gap:6px;padding:3px 16px 3px 0;margin-right:8px;border-right:1px solid #E1E2E6}}
.filter-group:last-of-type{{border-right:none;margin-right:0}}
.filter-label{{font-size:9.5px;font-weight:700;color:#A9B4BC;text-transform:uppercase;letter-spacing:.6px;white-space:nowrap;margin-right:2px}}
.filter-btn{{padding:4px 13px;border:1.5px solid #E1E2E6;border-radius:20px;background:#fff;font-size:11px;font-weight:600;color:#5D6A73;cursor:pointer;transition:all .15s;white-space:nowrap}}
.filter-btn:hover{{border-color:#0070F2;color:#0070F2;background:#EEF6FF}}
.filter-btn.active{{background:#0070F2;border-color:#0070F2;color:#fff;box-shadow:0 1px 4px rgba(0,112,242,.25)}}
.filter-select{{padding:4px 10px;border:1.5px solid #E1E2E6;border-radius:20px;font-size:11px;color:#5D6A73;background:#fff;cursor:pointer;max-width:200px;transition:border-color .15s}}
.filter-select:hover{{border-color:#0070F2}}
.filter-select:focus{{outline:none;border-color:#0070F2}}
.filter-reset{{display:inline-flex;align-items:center;gap:5px;margin-left:12px;padding:5px 13px;border:1.5px solid #E1E2E6;border-radius:20px;background:#fff;font-size:11px;font-weight:600;color:#5D6A73;cursor:pointer;transition:all .15s;white-space:nowrap}}
.filter-reset:hover{{border-color:#BB0000;color:#BB0000;background:#FFF5F5}}
.filter-active-badge{{display:none;margin-left:8px;background:#0070F2;color:#fff;font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px}}
.filter-active-badge.visible{{display:inline-block}}
.filter-result{{font-size:11px;color:#8396A8;margin-left:auto;white-space:nowrap}}
.global-export-btn{{display:inline-flex;align-items:center;gap:8px;padding:10px 20px;border:none;border-radius:8px;background:#107F3E;font-size:13px;font-weight:700;color:#fff;cursor:pointer;transition:all .15s;white-space:nowrap;flex-shrink:0;font-family:inherit;box-shadow:0 2px 8px rgba(16,127,62,.3)}}
.global-export-btn:hover{{background:#0D6A32;box-shadow:0 4px 14px rgba(16,127,62,.4);transform:translateY(-1px)}}
.global-export-btn:active{{transform:translateY(0);box-shadow:0 1px 4px rgba(16,127,62,.3)}}
th.sortable{{cursor:pointer;user-select:none;white-space:nowrap}}
th.sortable:hover{{background:#E8F4FF;color:#0070F2}}
th.sortable::after{{content:' \2195';font-size:10px;opacity:.35;margin-left:3px}}
th.sort-asc::after{{content:' \2191';opacity:1;color:#0070F2}}
th.sort-desc::after{{content:' \2193';opacity:1;color:#0070F2}}
.lc-pill{{display:inline-flex;align-items:center;gap:5px;cursor:pointer;user-select:none}}
.lc-pill input[type=checkbox]{{display:none}}
.lc-pill span{{display:inline-block;padding:3px 11px;border:1.5px solid #D8DCE0;border-radius:20px;font-size:11px;font-weight:600;color:#5D6A73;background:#fff;transition:all .15s;white-space:nowrap}}
.lc-pill input:checked + span{{background:#0070F2;border-color:#0070F2;color:#fff}}
.lc-pill:hover span{{border-color:#0070F2;color:#0070F2}}
.trend-dots{{display:inline-flex;gap:3px;align-items:center;margin-top:3px}}
.trend-dot{{width:7px;height:7px;border-radius:50%;flex-shrink:0}}
.td-better{{background:#107F3E}}
.td-worse{{background:#BB0000}}
.td-same{{background:#D8DCE0}}
.td-in{{background:#0070F2}}
.td-out{{background:#E1E2E6}}
.trend-na{{color:#D8DCE0;font-size:10px}}
.trend-hdr{{font-size:9px;color:#8396A8;white-space:nowrap}}
@media print{{
  .header{{position:relative}}
  .filter-bar,.filter-bar-2,.global-export-btn{{display:none}}
  .section{{break-inside:avoid}}
  body{{background:#fff}}
}}
</style>
</head>
<body>

<!-- HEADER -->
<div class="header">
  <div class="header-brand">
    <div class="header-logo">{_SAP_SVG}</div>
    <div class="header-sep"></div>
    <div class="header-info">
      <h1>CBS Oversight Tracker &mdash; All Market Units</h1>
      <p>Market Unit Supervisor Briefing &bull; Projects active as of {TODAY_STR} &bull; Source: FELIPE Monitor / MANDI / Leakage Report</p>
    </div>
  </div>
  <div class="header-actions">
    <div class="header-date">
      <span class="header-date-label">Oversight Call Prep</span>
      <strong>{TODAY_STR}</strong>
    </div>
    <button class="global-export-btn" id="export-btn" onclick="exportAllXLSX()" title="Export all 4 sections to Excel">
      <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Export Excel
    </button>
  </div>
</div>

<!-- KPI BAR -->
<div class="kpi-bar">
  <div class="kpi-card kpi-blue">
    <span class="kpi-group-lbl">Portfolio</span>
    <span class="kpi-val" id="kv-proj">{total_projects}</span>
    {d_proj}
    {_td_proj}
    <span class="kpi-lbl">Active Projects<br>All Market Units</span>
  </div>
  <div class="kpi-sep"></div>
  <div class="kpi-card kpi-red">
    <span class="kpi-group-lbl">MANDI Status</span>
    <span class="kpi-val" id="kv-red">{red_count}</span>
    {d_red}
    {_td_red}
    <span class="kpi-lbl">Red Status</span>
  </div>
  <div class="kpi-card kpi-amber">
    <span class="kpi-group-lbl">MANDI Status</span>
    <span class="kpi-val" id="kv-yel">{yellow_count}</span>
    {d_yel}
    <span class="kpi-lbl">Yellow Status</span>
  </div>
  <div class="kpi-sep"></div>
  <div class="kpi-card kpi-orange">
    <span class="kpi-group-lbl">Positive Leakage</span>
    <span class="kpi-val" id="kv-pos-n">{pos_leak_count}</span>
    {d_pos_n}
    {_td_pos_n}
    <span class="kpi-lbl">Projects with<br>Positive Leakage</span>
  </div>
  <div class="kpi-card kpi-orange">
    <span class="kpi-group-lbl">Positive Leakage</span>
    <span class="kpi-val-md" id="kv-pos-v">{_fmt_usd(total_pos_leak_val)}</span>
    {d_pos_v}
    {_td_pos_v}
    <span class="kpi-lbl">Total Exposure</span>
  </div>
  <div class="kpi-sep"></div>
  <div class="kpi-card kpi-amber">
    <span class="kpi-group-lbl">Negative Leakage</span>
    <span class="kpi-val" id="kv-neg-n">{neg_leak_count}</span>
    {d_neg_n}
    {_td_neg_n}
    <span class="kpi-lbl">High Neg. Leakage<br>(&gt;15%)</span>
  </div>
  <div class="kpi-card kpi-amber">
    <span class="kpi-group-lbl">Negative Leakage</span>
    <span class="kpi-val-md" id="kv-neg-v">{_fmt_usd(total_neg_leak_val)}</span>
    {d_neg_v}
    {_td_neg_v}
    <span class="kpi-lbl">Total Value</span>
  </div>
  <div class="kpi-sep"></div>
  <div class="kpi-card kpi-blue">
    <span class="kpi-group-lbl">Snapshot</span>
    <span class="kpi-val" id="kv-stale">{stale_snap_count}</span>
    {d_stale}
    {_td_stale}
    <span class="kpi-lbl">Stale / Missing<br>Snapshot</span>
  </div>
</div>

<!-- NAV -->
<nav class="nav">
  <a href="#s1" class="s1" data-sec="s1">
    <svg width="10" height="10" viewBox="0 0 10 10" fill="currentColor"><rect width="10" height="10" rx="2"/></svg>
    Red Status <span class="nav-badge" id="nb1">{red_count}</span>
  </a>
  <a href="#s2" class="s2" data-sec="s2">
    <svg width="10" height="10" viewBox="0 0 10 10" fill="currentColor"><polygon points="5,1 9,9 1,9"/></svg>
    Positive Leakage <span class="nav-badge" id="nb2">{pos_leak_count}</span>
  </a>
  <a href="#s3" class="s3" data-sec="s3">
    <svg width="10" height="10" viewBox="0 0 10 10" fill="currentColor"><polygon points="5,9 9,1 1,1"/></svg>
    High Negative Leakage <span class="nav-badge" id="nb3">{neg_leak_count}</span>
  </a>
  <a href="#s4" class="s4" data-sec="s4">
    <svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="6" cy="6" r="5"/><line x1="6" y1="3" x2="6" y2="6"/><line x1="6" y1="8.5" x2="6" y2="9"/></svg>
    Missing Snapshot <span class="nav-badge" id="nb4">{stale_snap_count}</span>
  </a>
</nav>

<!-- FILTER BAR -->
<div class="filter-bar">
  <div class="filter-group">
    <span class="filter-label">Market Unit</span>
    <button class="filter-btn active" id="mu-all" onclick="setFilter('mu','all',this)">All MUs</button>
{mu_buttons}
  </div>
  <div class="filter-group">
    <span class="filter-label">Portfolio</span>
    <button class="filter-btn active" id="seg-all" onclick="setFilter('seg','all',this)">All</button>
    <button class="filter-btn" id="seg-cbs"  onclick="setFilter('seg','Managed by CBS',this)">CBS Managed</button>
    <button class="filter-btn" id="seg-spot" onclick="setFilter('seg','SPOT Projects',this)">SPOT</button>
  </div>
  <div class="filter-group">
    <span class="filter-label">Team</span>
    <button class="filter-btn active" id="team-all" onclick="setFilter('team','all',this)">All</button>
    <button class="filter-btn" id="team-cbs"  onclick="setFilter('team','CBS',this)">CBS Analysts</button>
    <button class="filter-btn" id="team-apm"  onclick="setFilter('team','APM',this)">APMs</button>
  </div>
  <div class="filter-group">
    <span class="filter-label">CBS Responsible</span>
    <select class="filter-select" id="resp-select" onchange="setFilter('resp',this.value,null)">
      <option value="all">All Responsibles</option>
      {resp_options}
    </select>
  </div>
  <button class="filter-reset" onclick="resetFilters()">
    <svg xmlns="http://www.w3.org/2000/svg" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    Reset filters
  </button>
  <span class="filter-active-badge" id="active-badge"></span>
  <span class="filter-result" id="filter-status"></span>
</div>

<!-- SIZE + LIFECYCLE FILTER BAR -->
<div class="filter-bar-2">
  <div class="filter-group">
    <span class="filter-label">Contract Size (CBR)</span>
    <button class="filter-btn active" id="size-all" onclick="setFilter('size','all',this)">All</button>
    <button class="filter-btn" id="size-lt" onclick="setFilter('size','lt1.5m',this)">&lt; $1.5M</button>
    <button class="filter-btn" id="size-ge" onclick="setFilter('size','ge1.5m',this)">&#8805; $1.5M</button>
  </div>
  <div class="filter-group">
    <span class="filter-label">Lifecycle Status</span>
    <div style="display:flex;align-items:center;flex-wrap:wrap;gap:6px">
      <label class="lc-pill"><input type="checkbox" value="Not Started" onchange="toggleLC(this)"><span>Not Started</span></label>
      <label class="lc-pill"><input type="checkbox" value="Delivery in Process" onchange="toggleLC(this)"><span>In Process</span></label>
      <label class="lc-pill"><input type="checkbox" value="On Hold" onchange="toggleLC(this)"><span>On Hold</span></label>
      <label class="lc-pill"><input type="checkbox" value="Delivery Completed" onchange="toggleLC(this)"><span>Completed</span></label>
    </div>
  </div>
</div>

<div class="content">

<!-- SECTION 1: RED STATUS -->
{section_header(1, "Red Status Projects", red_count, 'R',
    f"Projects with MANDI Overall Status = Red. Sorted by consecutive weeks in red. "
    f"Includes SAP Margin Status, Fixable classification, and last Red Report entry.",
    [
        "&#191;Cu&#225;l es la situaci&#243;n actual de este proyecto y qu&#233; acci&#243;n concreta se est&#225; tomando esta semana?",
        "&#191;Ya hubo contacto formal con el PM / PCo / O2i? &#191;Hay evidencia de esa comunicaci&#243;n en el sistema?",
        "&#191;El clasificador 'Fixable' refleja la realidad actual o necesita actualizarse?",
        "&#191;Hay un plan de recuperaci&#243;n documentado o un CR / ICO en proceso?",
        "&#191;Cu&#225;ndo se espera el siguiente hito que cambie el status a Yellow o Green?",
    ]
)}
<div class="table-wrap">
<table>
<thead><tr>
  <th>Project ID</th><th>Sales Order</th><th>SO PM</th><th>Description</th><th>Customer</th><th>Bucket</th>
  <th>MP</th><th>CBS Responsible</th><th>MANDI</th><th>SAP Margin</th>
  <th>CBR</th><th>EAC Margin</th>
  <th>Fixable</th><th>Standard Comment</th><th>Weeks Red</th><th>Resolved</th><th class="trend-hdr">Trend</th>
</tr></thead>
<tbody>
{"".join(s1_row(row) for _, row in sec1.iterrows()) if len(sec1) > 0
 else '<tr><td colspan="17" class="empty-state">No red status projects found.</td></tr>'}
</tbody>
</table>
</div>
</div>

<!-- SECTION 2: POSITIVE LEAKAGE -->
{section_header(2, "Positive Leakage Projects", pos_leak_count, 'orange',
    f"Projects where Backlog Leakage &gt; 0 at the Sales Order level. "
    f"Total exposure: {_fmt_usd(total_pos_leak_val)}. Sorted by leakage amount (descending).",
    [
        "&#191;Se tiene un Change Request (CR) o amendment en proceso para formalizar el exceso de consumo?",
        "&#191;El PM ya coordin&#243; con O2i/Revenue Recognition para el ajuste del planning?",
        "&#191;Cu&#225;l es la causa ra&#237;z del overrun: scope creep, underestimation o problemas de ejecuci&#243;n?",
        "&#191;El cliente ya fue informado y hay acuerdo sobre el exceso?",
        "&#191;Hay un snapshot actualizado de FELIPE que refleje esta situaci&#243;n?",
    ]
)}
<div class="table-wrap">
<table>
<thead><tr>
  <th>Project ID</th><th>Sales Order</th><th>SO PM</th><th>Description</th><th>Customer</th><th>Bucket</th>
  <th>MP</th><th>CBS Responsible</th><th>Contract Type</th><th>Contract Net Value</th>
  <th>Positive Leakage</th><th>Leakage %</th><th>Comment</th><th>Last Snapshot</th><th class="trend-hdr">Trend</th>
</tr></thead>
<tbody>
{"".join(s2_row(row) for _, row in sec2.iterrows()) if len(sec2) > 0
 else '<tr><td colspan="15" class="empty-state">No positive leakage projects found.</td></tr>'}
</tbody>
</table>
</div>
</div>

<!-- SECTION 3: HIGH NEGATIVE LEAKAGE -->
{section_header(3, "High Negative Leakage (&gt;15% of Contract)", neg_leak_count, 'Y',
    f"Projects where projected backlog delivery is &gt;15% below contracted value. "
    f"Total negative exposure: {_fmt_usd(total_neg_leak_val)}. Sorted by leakage % (highest first).",
    [
        "&#191;Existe una justificaci&#243;n documentada para el underrun de backlog (scope reduction formal, billing block, etc.)?",
        "&#191;El planning en FELIPE refleja correctamente lo que se va a facturar al cliente?",
        "&#191;Hay un Billing Block activo que explique el gap? &#191;Cu&#225;ndo se resuelve?",
        "&#191;El cliente firm&#243; alg&#250;n Amendment que reduzca el alcance contratado?",
        "&#191;Se tiene visible este riesgo en el Revenue Forecast del &#225;rea?",
    ]
)}
<div class="table-wrap">
<table>
<thead><tr>
  <th>Project ID</th><th>Sales Order</th><th>SO PM</th><th>Description</th><th>Customer</th><th>Bucket</th>
  <th>MP</th><th>CBS Responsible</th><th>Contract Type</th><th>Contract Net Value</th>
  <th>Backlog Leakage</th><th>Leakage %</th><th>Comment</th><th>Last Snapshot</th><th class="trend-hdr">Trend</th>
</tr></thead>
<tbody>
{"".join(s3_row(row) for _, row in sec3.iterrows()) if len(sec3) > 0
 else '<tr><td colspan="15" class="empty-state">No high negative leakage projects found.</td></tr>'}
</tbody>
</table>
</div>
</div>

<!-- SECTION 4: MISSING SNAPSHOT -->
{section_header(4, "Missing / Stale FELIPE Snapshot", stale_snap_count, 'blue',
    "Active projects (Delivery in Process) where the last FELIPE change date is more recent "
    "than the last snapshot date, or where no snapshot has been taken this period.",
    [
        "&#191;Qu&#233; impidi&#243; que se hiciera el snapshot esta semana/mes?",
        "&#191;El PM tiene acceso a FELIPE y est&#225; al tanto de la obligaci&#243;n de mantener el snapshot actualizado?",
        "&#191;Hay un ETC pendiente de actualizar que bloquea el snapshot?",
        "Para proyectos 'Not Started': &#191;ya se inici&#243; la ejecuci&#243;n? &#191;El lifecycle status est&#225; correcto?",
        "&#191;Este proyecto aparece en el Red Report por falta de actualizaci&#243;n?",
    ]
)}
<div class="table-wrap">
<table>
<thead><tr>
  <th>Project ID</th><th>Sales Order</th><th>SO PM</th><th>Customer</th><th>Bucket</th><th>MP</th><th>Lifecycle</th>
  <th>CBS Responsible</th><th>End Date</th><th>Status</th>
  <th>Last Change Date</th><th>Last Snapshot Date</th><th>Gap</th>
  <th>Snapshot Type</th><th>MANDI</th><th class="trend-hdr">Trend</th>
</tr></thead>
<tbody>
{"".join(s4_row(row) for _, row in sec4.iterrows()) if len(sec4) > 0
 else '<tr><td colspan="16" class="empty-state">All active projects have up-to-date FELIPE snapshots.</td></tr>'}
</tbody>
</table>
</div>
</div>

</div><!-- /content -->

<div class="footer">
  CBS Portfolio Operations &bull; All Market Units &bull;
  Data sources: FELIPE Monitor, MANDI, Leakage Report &bull;
  Generated {TODAY_STR}
</div>

<script>
const PREV_MU   = {PREV_MU_JSON};
const PREV_DATE = '{PREV_DATE_STR}';
const FILTERS   = {{mu:'all', seg:'all', resp:'all', size:'all', team:'all', lc: new Set()}};

function setFilter(key, val, btn) {{
  FILTERS[key] = val;
  if (btn) {{
    const grp = btn.parentElement;
    grp.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
  }}
  applyFilters();
}}
function toggleLC(cb) {{
  if (cb.checked) FILTERS.lc.add(cb.value);
  else FILTERS.lc.delete(cb.value);
  applyFilters();
}}
function resetFilters() {{
  FILTERS.mu = FILTERS.seg = FILTERS.resp = FILTERS.size = FILTERS.team = 'all';
  FILTERS.lc.clear();
  document.querySelectorAll('.filter-btn').forEach(b => {{
    b.classList.toggle('active', b.id.endsWith('-all'));
  }});
  const sel = document.getElementById('resp-select');
  if (sel) sel.value = 'all';
  document.querySelectorAll('.lc-pill input[type=checkbox]').forEach(cb => cb.checked = false);
  applyFilters();
}}
function fmtKPI(val) {{
  const neg = val < 0, abs = Math.abs(val);
  let s;
  if (abs >= 1e9) s = '$' + (abs/1e9).toFixed(2) + 'B';
  else if (abs >= 1e6) s = '$' + (abs/1e6).toFixed(2) + 'M';
  else if (abs >= 1e3) s = '$' + (abs/1e3).toFixed(1) + 'K';
  else s = '$' + abs.toFixed(0);
  return neg ? '\u2212' + s : s;
}}
function applyFilters() {{
  const {{mu, seg, resp, size, team, lc}} = FILTERS;
  const counts = {{s1:0, s2:0, s3:0, s4:0}};
  let totalVis = 0, totalAll = 0, posLeakVal = 0, negLeakVal = 0;
  document.querySelectorAll('tbody tr[data-mu]').forEach(row => {{
    const rmu=row.dataset.mu||'', rseg=row.dataset.seg||'', rrsp=row.dataset.resp||'',
          rsz=row.dataset.size||'', rlc=row.dataset.lc||'', rtm=row.dataset.team||'';
    const sec = row.closest('.section')?.id;
    if (!sec) return;
    totalAll++;
    const show =
      (mu==='all'||rmu===mu) &&
      (seg==='all'||rseg===seg) &&
      (resp==='all'||rrsp===resp) &&
      (size==='all'||rsz===size) &&
      (team==='all'||rtm===team) &&
      (lc.size===0||lc.has(rlc));
    row.style.display = show ? '' : 'none';
    if (show) {{
      counts[sec] = (counts[sec]||0) + 1;
      totalVis++;
      const nv = parseFloat(row.dataset.numval);
      if (!isNaN(nv)) {{
        if (sec==='s2') posLeakVal += nv;
        if (sec==='s3') negLeakVal += nv;
      }}
    }}
  }});
  ['s1','s2','s3','s4'].forEach((sid, i) => {{
    const vc = document.getElementById('vc'+(i+1));
    if (vc) vc.textContent = counts[sid]||0;
    const nb = document.getElementById('nb'+(i+1));
    if (nb) nb.textContent = counts[sid]||0;
  }});
  const upd = (id, val) => {{ const el=document.getElementById(id); if(el) el.textContent=val; }};
  upd('kv-proj',  totalVis);
  upd('kv-red',   counts.s1||0);
  upd('kv-pos-n', counts.s2||0);
  upd('kv-pos-v', fmtKPI(posLeakVal));
  upd('kv-neg-n', counts.s3||0);
  upd('kv-neg-v', fmtKPI(negLeakVal));
  upd('kv-stale', counts.s4||0);
  if (PREV_DATE) updateDeltaPills(mu, posLeakVal, negLeakVal);
  const activeCount = (mu!=='all'?1:0)+(seg!=='all'?1:0)+(resp!=='all'?1:0)+(size!=='all'?1:0)+(team!=='all'?1:0)+(lc.size>0?1:0);
  const badge = document.getElementById('active-badge');
  if (badge) {{
    badge.textContent = activeCount > 0 ? `${{activeCount}} filter${{activeCount>1?'s':''}} active` : '';
    badge.classList.toggle('visible', activeCount > 0);
  }}
  const st = document.getElementById('filter-status');
  if (st) st.textContent = activeCount > 0 ? `${{totalVis}} of ${{totalAll}} projects` : '';
}}
function updateDeltaPills(mu, posV, negV) {{
  if (!PREV_DATE || !Object.keys(PREV_MU).length) return;
  const pmu = PREV_MU[mu] || PREV_MU['all'];
  if (!pmu) return;
  function setDP(id, curr, prev, lib, fmt) {{
    const el = document.getElementById(id);
    if (!el) return;
    const diff = curr - prev;
    if (Math.abs(diff) < 0.005) {{
      el.className = 'delta-kpi delta-nc';
      el.innerHTML = '\u2014\u2009vs\u00a0' + PREV_DATE;
      return;
    }}
    const better = lib ? diff < 0 : diff > 0;
    el.className  = 'delta-kpi ' + (better ? 'delta-better' : 'delta-worse');
    const arrow   = diff > 0 ? '\u25b2' : '\u25bc';
    let ds;
    if (fmt === 'usd') {{
      const ad = Math.abs(diff), sign = diff > 0 ? '+' : '\u2212';
      ds = ad>=1e6 ? sign+'$'+(ad/1e6).toFixed(2)+'M' : ad>=1e3 ? sign+'$'+(ad/1e3).toFixed(0)+'K' : sign+'$'+ad.toFixed(0);
    }} else {{
      ds = (diff > 0 ? '+' : '') + Math.abs(Math.round(diff));
    }}
    el.innerHTML = arrow + '\u2009' + ds + '\u00a0vs\u00a0' + PREV_DATE;
  }}
  const ps1=(pmu.s1||{{c:0}}).c, ps2=(pmu.s2||{{c:0,v:0}}), ps3=(pmu.s3||{{c:0,v:0}}), ps4=(pmu.s4||{{c:0}}).c;
  const cRed  =parseInt(document.getElementById('kv-red')  ?.textContent||'0');
  const cPosN =parseInt(document.getElementById('kv-pos-n')?.textContent||'0');
  const cNegN =parseInt(document.getElementById('kv-neg-n')?.textContent||'0');
  const cStale=parseInt(document.getElementById('kv-stale')?.textContent||'0');
  const cProj =parseInt(document.getElementById('kv-proj') ?.textContent||'0');
  setDP('dp-proj',  cProj,          pmu.proj||0,   false, 'int');
  setDP('dp-red',   cRed,           ps1,           true,  'int');
  setDP('dp-pos-n', cPosN,          ps2.c,         true,  'int');
  setDP('dp-pos-v', posV,           ps2.v,         true,  'usd');
  setDP('dp-neg-n', cNegN,          ps3.c,         true,  'int');
  setDP('dp-neg-v', Math.abs(negV), Math.abs(ps3.v),true, 'usd');
  setDP('dp-stale', cStale,         ps4,           true,  'int');
}}
function exportAllXLSX() {{
  const btn = document.getElementById('export-btn');
  if (typeof XLSX === 'undefined') {{
    if (btn) {{ btn.textContent = 'Loading...'; btn.disabled = true; }}
    const s = document.createElement('script');
    s.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
    s.onload  = () => {{ if (btn) btn.disabled = false; exportAllXLSX(); }};
    s.onerror = () => {{ if (btn) {{ btn.innerHTML = '&#9888; Error'; btn.disabled = false; }} }};
    document.head.appendChild(s);
    return;
  }}
  const origHTML = btn ? btn.innerHTML : '';
  if (btn) {{ btn.textContent = 'Building...'; btn.disabled = true; }}
  setTimeout(() => {{
    try {{
      const sections = [
        {{id:'s1',name:'Red Status'}},{{id:'s2',name:'Positive Leakage'}},
        {{id:'s3',name:'Neg Leakage'}},{{id:'s4',name:'Missing Snapshot'}},
      ];
      const wb = XLSX.utils.book_new();
      sections.forEach(sec => {{
        const el = document.getElementById(sec.id), table = el?.querySelector('table');
        if (!table) return;
        const aoa = [];
        const headers = Array.from(table.querySelectorAll('thead th')).map(th=>th.innerText.trim());
        aoa.push(headers);
        table.querySelectorAll('tbody tr').forEach(tr => {{
          if (tr.style.display==='none') return;
          const tds = tr.querySelectorAll('td');
          if (tds.length===1 && tds[0].colSpan>1) return;
          aoa.push(Array.from(tds).map(td=>td.innerText.trim()));
        }});
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws['!cols'] = headers.map((h,i)=>{{
          const maxLen = Math.max(h.length,...aoa.slice(1).map(r=>(r[i]||'').length));
          return {{wch:Math.min(maxLen+2,45)}};
        }});
        XLSX.utils.book_append_sheet(wb, ws, sec.name);
      }});
      const ds = new Date().toISOString().slice(0,10).replace(/-/g,'');
      XLSX.writeFile(wb, `CBS_Oversight_Tracker_${{ds}}.xlsx`);
    }} finally {{
      if (btn) {{ btn.innerHTML = origHTML; btn.disabled = false; }}
    }}
  }}, 50);
}}
function extractSortVal(cell) {{
  const text = (cell.innerText||'').trim();
  if (text==='--'||text==='') return '\uFFFF';
  const mM=text.match(/^[\u2212-]?\$?([\d.]+)M$/);
  if (mM) return parseFloat(mM[1])*(text[0]==='\u2212'||text[0]==='-'?-1:1)*1e6;
  const mK=text.match(/^[\u2212-]?\$?([\d.]+)K$/);
  if (mK) return parseFloat(mK[1])*(text[0]==='\u2212'||text[0]==='-'?-1:1)*1e3;
  const mP=text.match(/^[\u2212-]?([\d.]+)%$/);
  if (mP) return parseFloat(mP[1])*(text[0]==='\u2212'||text[0]==='-'?-1:1);
  const mW=text.match(/^(\d+)w$/); if (mW) return parseInt(mW[1]);
  const mD=text.match(/^(\d+) days$/); if (mD) return parseInt(mD[1]);
  if (text==='\u2014 no snapshot') return Infinity;
  const mN=parseFloat(text.replace(/[^0-9.\-]/g,''));
  if (!isNaN(mN)&&/^[\d.\-$]/.test(text)) return mN;
  const d=new Date(text); if (!isNaN(d.getTime())) return d.getTime();
  return text.toLowerCase();
}}
function sortTable(th) {{
  const table=th.closest('table'), tbody=table.querySelector('tbody');
  const colIdx=Array.from(th.parentElement.children).indexOf(th);
  const asc=th.dataset.sortDir!=='asc';
  th.closest('thead').querySelectorAll('th').forEach(h=>{{h.dataset.sortDir='';h.classList.remove('sort-asc','sort-desc');}});
  th.dataset.sortDir=asc?'asc':'desc';
  th.classList.add(asc?'sort-asc':'sort-desc');
  const rows=Array.from(tbody.querySelectorAll('tr[data-mu]'));
  rows.sort((a,b)=>{{
    const av=extractSortVal(a.children[colIdx]), bv=extractSortVal(b.children[colIdx]);
    if (typeof av==='number'&&typeof bv==='number') return asc?av-bv:bv-av;
    return asc?String(av).localeCompare(String(bv),undefined,{{sensitivity:'base'}}):String(bv).localeCompare(String(av),undefined,{{sensitivity:'base'}});
  }});
  rows.forEach(r=>tbody.appendChild(r));
}}
document.addEventListener('DOMContentLoaded', () => {{
  document.querySelectorAll('table thead th').forEach(th => {{
    th.classList.add('sortable');
    th.addEventListener('click', () => sortTable(th));
  }});
  const navLinks=Array.from(document.querySelectorAll('.nav a[data-sec]'));
  const sections=navLinks.map(a=>document.getElementById(a.dataset.sec));
  if ('IntersectionObserver' in window) {{
    const spy=new IntersectionObserver(entries=>{{
      entries.forEach(entry=>{{
        if (entry.isIntersecting) {{
          navLinks.forEach(a=>a.classList.remove('spy-active'));
          const link=document.querySelector(`.nav a[data-sec="${{entry.target.id}}"]`);
          if (link) link.classList.add('spy-active');
        }}
      }});
    }}, {{rootMargin:'-15% 0px -70% 0px',threshold:0}});
    sections.forEach(s=>{{ if (s) spy.observe(s); }});
  }}
}});
</script>
</body>
</html>'''

    return html, _snap_json_out
