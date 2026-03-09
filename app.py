"""
Alpine Macro — Readership Analytics Web App v6
Run with: streamlit run app.py

Metric changes from v5:
- Email: initial_open (reach) + open (re-opens) tracked separately. Clicks removed.
- Portal: OpenLandingPage (views) + OpenChapter (deep reads) + DownloadProduct (downloads)
- Most Active Readers: initial_open + open + OpenLandingPage + OpenChapter combined
- Top Reports: email initial_open only
- Top Authors: email initial_open only
- Donut chart: email initial_open by Leaf product (unchanged)
"""

import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import io, os, base64
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime

st.set_page_config(page_title="Alpine Macro — Readership Analytics", page_icon="📊", layout="wide")

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "alpine_logo.png")

def get_logo_b64():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None

logo_b64  = get_logo_b64()
logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height:64px;">' if logo_b64 else "<strong>ALPINE MACRO</strong>"

C_BLUE  = "#0077C8"
C_DARK  = "#2B2B2B"
C_LBLUE = "#F0F7FD"
C_LGREY = "#F4F8FC"
C_RULE  = "#D0DCE8"
CHART_COLORS = ["#0077C8", "#005A96", "#33A1E0", "#80C4ED", "#B3DDF5", "#E6F4FC"]

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@400;600&family=DM+Sans:wght@300;400;500;600&display=swap');
html, body, [class*="css"] {{ font-family: 'DM Sans', sans-serif; background: {C_LGREY}; }}
.am-nav {{ background:#fff; border-bottom:3px solid {C_BLUE}; padding:16px 28px; display:flex; align-items:center; justify-content:space-between; margin-bottom:28px; }}
.am-nav-right {{ font-size:0.72rem; color:#888; text-transform:uppercase; letter-spacing:0.1em; font-weight:500; }}
.kpi-card {{ background:#fff; border-radius:10px; padding:16px 14px 12px; border-left:4px solid var(--ac); box-shadow:0 1px 4px rgba(0,119,200,0.08); }}
.kpi-val {{ font-family:'Source Serif 4',serif; font-size:1.9rem; font-weight:600; color:{C_DARK}; line-height:1; margin-bottom:4px; }}
.kpi-lbl {{ font-size:0.68rem; color:#6B6B6B; text-transform:uppercase; letter-spacing:0.07em; font-weight:500; }}
.kpi-sub {{ font-size:0.65rem; color:#999; margin-top:2px; }}
.sec-title {{ font-family:'Source Serif 4',serif; font-size:1.05rem; font-weight:600; color:{C_BLUE}; border-bottom:2px solid {C_BLUE}; padding-bottom:5px; margin:20px 0 12px; display:inline-block; }}
.tbl-wrap {{ background:#fff; border-radius:10px; padding:14px 16px; box-shadow:0 1px 4px rgba(0,119,200,0.06); }}
.tbl-hdr {{ display:flex; gap:4px; font-size:0.68rem; font-weight:600; color:#aaa; text-transform:uppercase; letter-spacing:0.06em; padding-bottom:6px; border-bottom:1px solid {C_RULE}; margin-bottom:2px; }}
.tbl-row {{ display:flex; gap:4px; align-items:center; padding:5px 0; font-size:0.78rem; border-bottom:1px solid #F4F8FC; }}
.tbl-row:last-child {{ border-bottom:none; }}
.tbl-row:nth-child(odd) {{ background:#FAFCFE; border-radius:4px; }}
.t-rank {{ font-weight:700; color:{C_BLUE}; min-width:18px; font-size:0.7rem; }}
.t-name {{ flex:2; color:{C_DARK}; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
.t-num  {{ flex:1; text-align:right; font-weight:600; color:#444; font-size:0.74rem; }}
.t-sub  {{ flex:1; text-align:right; color:#888; font-size:0.72rem; }}
.quality-box {{ background:#FFFBF0; border:1px solid #F0D080; border-radius:8px; padding:10px 14px; font-size:0.78rem; color:#7A5C00; margin-bottom:14px; }}
div[data-testid="stFileUploader"] {{ border:2px dashed {C_BLUE} !important; border-radius:10px !important; background:{C_LBLUE} !important; }}
.dl-btn > button {{ background:{C_DARK} !important; color:#fff !important; border:none !important; border-radius:8px !important; font-weight:600 !important; font-size:0.9rem !important; padding:0.55rem 2rem !important; width:100%; transition:background 0.2s; }}
.dl-btn > button:hover {{ background:{C_BLUE} !important; }}
.chip {{ display:inline-block; background:{C_LBLUE}; color:{C_BLUE}; border-radius:20px; padding:2px 10px; font-size:0.72rem; font-weight:500; margin:2px; }}
.section-divider {{ border:none; border-top:2px solid {C_RULE}; margin:24px 0 8px; }}
.channel-label {{ font-size:0.7rem; font-weight:600; text-transform:uppercase; letter-spacing:0.08em; color:#888; margin-bottom:6px; }}
</style>
""", unsafe_allow_html=True)

st.markdown(f'<div class="am-nav">{logo_html}<span class="am-nav-right">Readership Analytics</span></div>', unsafe_allow_html=True)

# ── Column aliases ─────────────────────────────────────────────────────────────
REQUIRED = ['Contact Name','Contact Email','EventSource','EventAction','EventDate','Report Title','Authors','Leaf product']
ALIASES  = {
    'contact name':'Contact Name', 'name':'Contact Name',
    'contact email':'Contact Email', 'email':'Contact Email',
    'eventsource':'EventSource', 'source':'EventSource', 'event source':'EventSource',
    'eventaction':'EventAction', 'action':'EventAction', 'event action':'EventAction',
    'eventdate':'EventDate', 'date':'EventDate', 'event date':'EventDate',
    'report title':'Report Title', 'title':'Report Title', 'report':'Report Title',
    'authors':'Authors', 'author':'Authors',
    'leaf product':'Leaf product', 'product':'Leaf product', 'category':'Leaf product',
}

# ── Event action constants ─────────────────────────────────────────────────────
# Email
EA_EMAIL_REACH   = 'initial_open'       # first open — reach metric
EA_EMAIL_REOPEN  = 'open'               # subsequent opens — re-engagement

# Portal (My Oxford)
EA_PORTAL_VIEW   = 'openlandingpage'    # viewed licensed product page — portal reach
EA_PORTAL_READ   = 'openchapter'        # read content inside report — deep engagement
EA_PORTAL_DL     = 'downloadproduct'    # downloaded report

# Sources
SRC_EMAIL  = 'email'
SRC_PORTAL = ['my oxford', 'myoxford', 'portal']

# ── Data loading & cleaning ────────────────────────────────────────────────────
def load_clean(file_bytes, filename):
    issues = []
    try:
        df = pd.read_csv(io.BytesIO(file_bytes)) if filename.endswith('.csv') else pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        st.error(f"Could not read file: {e}"); st.stop()

    # Normalise column names
    df.columns = df.columns.str.strip()
    rmap = {c: ALIASES[c.lower().strip()] for c in df.columns
            if c.lower().strip() in ALIASES and ALIASES[c.lower().strip()] not in df.columns}
    if rmap:
        df = df.rename(columns=rmap)
        issues.append(f"Auto-renamed columns: {', '.join(rmap.keys())}")

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {', '.join(missing)}"); st.stop()

    # Drop fully blank rows
    before = len(df)
    df = df.dropna(how='all')
    if before - len(df):
        issues.append(f"Removed {before - len(df)} blank rows")

    # Clean text fields — strip whitespace and null-like strings
    for col in ['Contact Name','Contact Email','EventAction','EventSource','Report Title','Authors','Leaf product']:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({'nan':'', 'NaN':'', 'NULL':'', 'none':'', 'None':''})
        df[col] = df[col].replace('', None)

    # Drop rows missing critical fields
    before = len(df)
    df = df.dropna(subset=['Contact Name','Contact Email','Report Title','EventAction','EventDate'])
    if before - len(df):
        issues.append(f"Removed {before - len(df)} rows with missing critical fields")

    # Normalise event fields to lowercase
    df['EventAction'] = df['EventAction'].str.lower().str.strip()
    df['EventSource']  = df['EventSource'].str.lower().str.strip()

    # Parse dates
    df['EventDate'] = pd.to_datetime(df['EventDate'], errors='coerce', infer_datetime_format=True)
    bad = df['EventDate'].isna().sum()
    if bad:
        issues.append(f"Removed {bad} rows with unparseable dates")
        df = df.dropna(subset=['EventDate'])

    # Filter out login and auth events
    login_actions  = ['login', 'loginfederatedidentity', 'loginipaddress', 'logout', 'ssogeorestriction']
    df = df[~df['EventAction'].isin(login_actions)]
    df = df[~df['EventSource'].isin(['login.oxfordeconomics.com'])]

    # Deduplicate: same reader + action + report + date = one event
    before = len(df)
    df = df.drop_duplicates(subset=['Contact Email','EventAction','Report Title','EventDate'])
    if before - len(df):
        issues.append(f"Removed {before - len(df)} duplicate events")

    df['Month'] = df['EventDate'].dt.to_period('M')
    return df, issues

# ── Analysis ───────────────────────────────────────────────────────────────────
def analyse(df):
    # ── Email segments ─────────────────────────────────────────────────────────
    email_df     = df[df['EventSource'] == SRC_EMAIL]
    email_reach  = email_df[email_df['EventAction'] == EA_EMAIL_REACH]   # initial_open
    email_reopen = email_df[email_df['EventAction'] == EA_EMAIL_REOPEN]  # open

    # ── Portal segments ────────────────────────────────────────────────────────
    portal_df    = df[df['EventSource'].isin(SRC_PORTAL)]
    portal_views = portal_df[portal_df['EventAction'] == EA_PORTAL_VIEW]  # OpenLandingPage
    portal_reads = portal_df[portal_df['EventAction'] == EA_PORTAL_READ]  # OpenChapter
    portal_dls   = portal_df[portal_df['EventAction'] == EA_PORTAL_DL]    # DownloadProduct

    # ── Top products — email reach (initial_open), drop blank ──────────────────
    top_products = (
        email_reach[email_reach['Leaf product'].notna()]
        ['Leaf product'].str.strip().value_counts().head(8)
    )

    # ── Top reports — email initial_open + portal OpenLandingPage combined ───
    reports_combined = pd.concat([
        email_reach[email_reach['Report Title'].notna()][['Report Title']],
        portal_views[portal_views['Report Title'].notna()][['Report Title']],
    ])
    top_reports = reports_combined['Report Title'].str.strip().value_counts().head(5)

    # ── Top authors — email initial_open only ─────────────────────────────────
    top_authors = (
        email_reach[email_reach['Authors'].notna()]
        ['Authors'].str.strip().value_counts().head(5)
    )

    # ── Most Active Readers ────────────────────────────────────────────────────
    # Columns: Email Reach (initial_open) | Email Re-opens (open) |
    #          Portal Views (OpenLandingPage) | Portal Reads (OpenChapter) | Total
    e_reach_cnt  = email_reach.groupby('Contact Name').size().rename('Email Reach')
    e_reopen_cnt = email_reopen.groupby('Contact Name').size().rename('Email Re-opens')
    p_view_cnt   = portal_views.groupby('Contact Name').size().rename('Portal Views')
    p_read_cnt   = portal_reads.groupby('Contact Name').size().rename('Portal Deep Reads')

    readers = pd.concat([e_reach_cnt, e_reopen_cnt, p_view_cnt, p_read_cnt], axis=1).fillna(0).astype(int)
    readers['Total'] = readers['Email Reach'] + readers['Email Re-opens'] + readers['Portal Views'] + readers['Portal Deep Reads']
    readers_all = readers.sort_values('Total', ascending=False).reset_index()
    readers_top = readers_all.head(10).reset_index(drop=True)

    # ── Unique readers — anyone with at least one reach/view/read event ────────
    all_engaged = pd.concat([email_reach, email_reopen, portal_views, portal_reads])
    unique_readers = all_engaged['Contact Name'].nunique()

    return {
        'date_min':           df['EventDate'].min().strftime('%d %b %Y'),
        'date_max':           df['EventDate'].max().strftime('%d %b %Y'),
        # Email KPIs
        'email_reach':        len(email_reach),          # initial_open count
        'email_reach_unique': email_reach['Contact Name'].nunique(),
        'email_reopens':      len(email_reopen),         # open count
        # Portal KPIs
        'portal_views':       len(portal_views),         # OpenLandingPage
        'portal_views_unique':portal_views['Contact Name'].nunique(),
        'portal_reads':       len(portal_reads),         # OpenChapter
        'portal_downloads':   len(portal_dls),           # DownloadProduct
        # Summary
        'unique_readers':     unique_readers,
        # Rankings
        'top_products':       top_products,
        'top_reports':        top_reports,
        'top_authors':        top_authors,
        'readers_top':        readers_top,
        'readers_all':        readers_all,
    }

# ── Charts (PDF only) ──────────────────────────────────────────────────────────
def make_channel_chart(data):
    fig, ax = plt.subplots(figsize=(5.5, 2.6))

    categories  = ['Email\nReach', 'Email\nRe-opens', 'Portal\nViews', 'Portal\nDeep Reads', 'Portal\nDownloads']
    vals        = [data['email_reach'], data['email_reopens'],
                   data['portal_views'], data['portal_reads'], data['portal_downloads']]
    bar_colors  = [CHART_COLORS[0], CHART_COLORS[2], CHART_COLORS[1], CHART_COLORS[3], CHART_COLORS[4]]

    x = np.arange(len(categories))
    bars = ax.bar(x, vals, color=bar_colors, zorder=3, width=0.55)

    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    ax.yaxis.grid(True, color='#CBD2DA', linewidth=0.5, zorder=0); ax.set_axisbelow(True)
    for s in ['top','right','left']: ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color('#E0E8F0')
    ax.tick_params(axis='y', labelsize=7, colors='#555')
    ax.set_xticks(x); ax.set_xticklabels(categories, fontsize=7.5, color='#333')

    mx = max(vals) or 1
    for xi, v in zip(x, vals):
        if v > 0:
            ax.text(xi, v + mx * 0.02, f'{v:,}', ha='center', va='bottom',
                    fontsize=7, color='#333', fontweight='bold')

    # Channel separator line
    ax.axvline(x=1.5, color='#D0DCE8', linewidth=1, linestyle='--', zorder=2)
    ax.text(0.5, ax.get_ylim()[1]*0.95, 'Email', ha='center', fontsize=7, color='#0077C8', fontweight='bold')
    ax.text(3.0, ax.get_ylim()[1]*0.95, 'Portal', ha='center', fontsize=7, color='#005A96', fontweight='bold')

    ax.set_title('Engagement by Channel & Event Type', fontsize=9, color='#2B2B2B',
                 pad=8, fontweight='bold', loc='left')
    fig.tight_layout(pad=0.5)
    return fig

def make_donut(top_products):
    if top_products is None or len(top_products) == 0:
        fig, ax = plt.subplots(figsize=(5.5, 4.2))
        ax.text(0.5, 0.5, 'No product data', ha='center', va='center', transform=ax.transAxes)
        ax.axis('off'); return fig

    n      = min(len(top_products), len(CHART_COLORS))
    labels = list(top_products.index)[:n]
    vals   = list(top_products.values)[:n]
    cols   = CHART_COLORS[:n]

    fig, ax = plt.subplots(figsize=(5.5, 4.2))
    ax.pie(vals, labels=None, colors=cols, startangle=90,
           wedgeprops=dict(width=0.52, edgecolor='white', linewidth=1.8))
    ax.set_facecolor('none'); fig.patch.set_alpha(0)

    patches = [mpatches.Patch(color=cols[i], label=f'{labels[i]}   ({vals[i]:,})')
               for i in range(n)]
    ax.legend(handles=patches, loc='lower center', bbox_to_anchor=(0.5, -0.48),
              fontsize=7.2, frameon=False, ncol=1, labelspacing=0.55,
              handlelength=1.2, handleheight=0.9)
    ax.set_title('Email Reach by Product Category', fontsize=9, color='#2B2B2B',
                 pad=8, fontweight='bold', loc='left')
    fig.tight_layout(pad=0.8)
    return fig

# ── Web table renderers ────────────────────────────────────────────────────────
def render_simple(rows, cols):
    hdr = '<div class="tbl-hdr">' + ''.join([
        '<span class="t-rank">#</span>' if i == 0 else
        f'<span class="t-name">{c}</span>' if i == 1 else
        f'<span class="t-num">{c}</span>'
        for i, c in enumerate(cols)]) + '</div>'
    body = ''
    for i, row in enumerate(rows):
        v = list(row)
        cells = f'<span class="t-rank">{i+1}</span><span class="t-name">{v[0]}</span>'
        for x in v[1:]: cells += f'<span class="t-num">{x}</span>'
        body += f'<div class="tbl-row">{cells}</div>'
    return f'<div class="tbl-wrap">{hdr}{body}</div>'

def render_readers(readers_df):
    hdr = ('<div class="tbl-hdr">'
           '<span class="t-rank">#</span>'
           '<span class="t-name">Reader</span>'
           '<span class="t-sub">📧 Reach</span>'
           '<span class="t-sub">📧 Re-opens</span>'
           '<span class="t-sub">🌐 Views</span>'
           '<span class="t-sub">🌐 Reads</span>'
           '<span class="t-num">Total</span>'
           '</div>')
    body = ''
    for rank, (_, row) in enumerate(readers_df.iterrows(), start=1):
        body += (f'<div class="tbl-row">'
                 f'<span class="t-rank">{rank}</span>'
                 f'<span class="t-name">{row["Contact Name"]}</span>'
                 f'<span class="t-sub">{int(row["Email Reach"]):,}</span>'
                 f'<span class="t-sub">{int(row["Email Re-opens"]):,}</span>'
                 f'<span class="t-sub">{int(row["Portal Views"]):,}</span>'
                 f'<span class="t-sub">{int(row["Portal Deep Reads"]):,}</span>'
                 f'<span class="t-num">{int(row["Total"]):,}</span>'
                 f'</div>')
    return f'<div class="tbl-wrap">{hdr}{body}</div>'

def kpi_card(label, value, ac, sub=None):
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ''
    return (f'<div class="kpi-card" style="--ac:{ac}">'
            f'<div class="kpi-val">{value}</div>'
            f'<div class="kpi-lbl">{label}</div>'
            f'{sub_html}'
            f'</div>')

# ── PDF constants ──────────────────────────────────────────────────────────────
PBLUE  = colors.HexColor("#0077C8")
PDARK  = colors.HexColor("#2B2B2B")
PLBLU  = colors.HexColor("#F0F7FD")
PRULE  = colors.HexColor("#D0DCE8")
PWHITE = colors.white
PMGREY = colors.HexColor("#999999")
PNBLUE = colors.HexColor("#005A96")

def fig_to_ir(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=160, bbox_inches='tight', transparent=True)
    buf.seek(0); plt.close(fig)
    return ImageReader(buf)

# ── PDF generation ─────────────────────────────────────────────────────────────
def generate_pdf(data, account_name):
    buf = io.BytesIO()
    W, H = A4
    c = canvas.Canvas(buf, pagesize=A4)
    M  = 12 * mm
    CW = (W - 2*M - 4*mm) / 2

    all_readers    = data['readers_all']
    ROWS_PER_PAGE  = 38
    reader_pages   = max(1, -(-len(all_readers) // ROWS_PER_PAGE))
    total_pages    = 1 + reader_pages

    def draw_header():
        c.setFillColor(PWHITE); c.rect(0, H-38*mm, W, 38*mm, fill=1, stroke=0)
        c.setFillColor(PBLUE);  c.rect(0, H-1.5*mm, W, 1.5*mm, fill=1, stroke=0)
        if os.path.exists(LOGO_PATH):
            c.drawImage(LOGO_PATH, M, H-34*mm, width=54*mm, height=24*mm,
                        preserveAspectRatio=True, mask='auto')
        c.setFont("Helvetica-Bold", 15); c.setFillColor(PDARK)
        c.drawRightString(W-M, H-17*mm, "Readership Analytics Report")
        c.setFont("Helvetica", 8.5); c.setFillColor(PBLUE)
        c.drawRightString(W-M, H-24*mm, f"{account_name}  |  {data['date_min']} – {data['date_max']}")
        c.setFont("Helvetica", 7.5); c.setFillColor(PMGREY)
        c.drawRightString(W-M, H-30*mm, f"Generated {datetime.now().strftime('%d %b %Y')}")
        c.setFillColor(PRULE); c.rect(0, H-39*mm, W, 0.5*mm, fill=1, stroke=0)

    def draw_footer(page_num):
        c.setFillColor(PDARK); c.rect(0, 0, W, 9*mm, fill=1, stroke=0)
        c.setFillColor(PBLUE); c.rect(0, 9*mm, W, 0.8*mm, fill=1, stroke=0)
        c.setFont("Helvetica", 6.5); c.setFillColor(colors.HexColor("#AAAAAA"))
        c.drawString(M, 3.2*mm, f"{account_name}  |  Readership Analytics  |  Confidential")
        c.drawRightString(W-M, 3.2*mm,
            f"Alpine Macro — An Oxford Economics Company  |  Page {page_num} of {total_pages}")

    def draw_table_row(x, y, w, rank, label, val, shade, rh, max_chars=40):
        if shade:
            c.setFillColor(PLBLU); c.rect(x, y, w, rh-0.5*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7); c.setFillColor(PBLUE)
        c.drawString(x+1.5*mm, y+1.8*mm, str(rank))
        c.setFont("Helvetica", 7); c.setFillColor(PDARK)
        short = label[:max_chars] + '...' if len(label) > max_chars else label
        c.drawString(x+6*mm, y+1.8*mm, short)
        c.setFont("Helvetica-Bold", 7); c.setFillColor(PNBLUE)
        c.drawRightString(x+w-1*mm, y+1.8*mm, str(val))
        return y - rh

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 1 — Summary
    # ══════════════════════════════════════════════════════════════════════════
    draw_header()

    # ── KPI tiles: 3 Email + 3 Portal ─────────────────────────────────────────
    kpis = [
        ("Email Reach",       f"{data['email_reach']:,}",        PBLUE,  "initial_open events"),
        ("Email Re-opens",    f"{data['email_reopens']:,}",       colors.HexColor("#33A1E0"), "open events"),
        ("Unique Readers",    f"{data['unique_readers']:,}",      PDARK,  "all channels"),
        ("Portal Views",      f"{data['portal_views']:,}",        PNBLUE, "OpenLandingPage"),
        ("Portal Deep Reads", f"{data['portal_reads']:,}",        colors.HexColor("#80C4ED"), "OpenChapter"),
        ("Portal Downloads",  f"{data['portal_downloads']:,}",    PMGREY, "DownloadProduct"),
    ]
    bw = (W - 2*M) / len(kpis); by = H-59*mm; bh = 17*mm
    for i, (lbl, val, col, hint) in enumerate(kpis):
        bx = M + i * bw
        c.setFillColor(PLBLU); c.roundRect(bx, by, bw-1.5*mm, bh, 3, fill=1, stroke=0)
        c.setFillColor(col);   c.rect(bx, by, 2*mm, bh, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 12); c.setFillColor(PDARK)
        c.drawCentredString(bx+2*mm+(bw-3.5*mm)/2, by+8*mm, val)
        c.setFont("Helvetica", 6); c.setFillColor(colors.HexColor("#666666"))
        c.drawCentredString(bx+2*mm+(bw-3.5*mm)/2, by+4*mm, lbl.upper())
        c.setFont("Helvetica-Oblique", 5.5); c.setFillColor(PMGREY)
        c.drawCentredString(bx+2*mm+(bw-3.5*mm)/2, by+1.5*mm, hint)

    # ── Channel separator label ────────────────────────────────────────────────
    sep_y = by - 4*mm
    # Email bracket
    c.setFillColor(PBLUE); c.rect(M, sep_y, bw*2-1.5*mm, 0.8*mm, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 6.5); c.setFillColor(PBLUE)
    c.drawString(M+1*mm, sep_y-3.5*mm, "📧  EMAIL CHANNEL")
    # Portal bracket
    c.setFillColor(PNBLUE); c.rect(M+bw*3, sep_y, bw*3-1.5*mm, 0.8*mm, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 6.5); c.setFillColor(PNBLUE)
    c.drawString(M+bw*3+1*mm, sep_y-3.5*mm, "🌐  PORTAL CHANNEL")

    # ── Charts ────────────────────────────────────────────────────────────────
    cy = H - 126*mm
    c.drawImage(fig_to_ir(make_channel_chart(data)), M, cy,
                width=108*mm, height=52*mm, preserveAspectRatio=True, mask='auto')
    c.drawImage(fig_to_ir(make_donut(data['top_products'])), M+112*mm, cy-14*mm,
                width=74*mm, height=68*mm, preserveAspectRatio=True, mask='auto')

    # ── Divider ───────────────────────────────────────────────────────────────
    dy = cy - 8*mm
    c.setStrokeColor(PRULE); c.setLineWidth(0.5)
    c.line(M, dy, W-M, dy)

    rh  = 6*mm
    ty  = dy - 4*mm
    y1  = ty

    # ── Left column — Top Reports + Top Authors ───────────────────────────────
    c.setFillColor(PBLUE); c.rect(M, y1, CW, 5.5*mm, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 7.5); c.setFillColor(PWHITE)
    c.drawString(M+3*mm, y1+1.5*mm, "TOP REPORTS  (Email Reach + Portal Views)")
    y1 -= rh
    for i, (t, v) in enumerate(data['top_reports'].items()):
        y1 = draw_table_row(M, y1, CW, i+1, t, f"{v:,}", i%2==1, rh)

    y1 -= 2.5*mm
    c.setFillColor(PBLUE); c.rect(M, y1, CW, 5.5*mm, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 7.5); c.setFillColor(PWHITE)
    c.drawString(M+3*mm, y1+1.5*mm, "TOP AUTHORS  (Email Reach)")
    y1 -= rh
    for i, (a, v) in enumerate(data['top_authors'].items()):
        y1 = draw_table_row(M, y1, CW, i+1, a, f"{v:,}", i%2==1, rh)

    # ── Right column — Top 10 Readers ────────────────────────────────────────
    rx = M + CW + 4*mm; rw = CW; y2 = ty

    c.setFillColor(PBLUE); c.rect(rx, y2, rw, 5.5*mm, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 7.5); c.setFillColor(PWHITE)
    c.drawString(rx+3*mm, y2+1.5*mm, f"MOST ACTIVE READERS  (Top 10 of {len(all_readers)})")
    y2 -= rh

    # Sub-header
    c.setFont("Helvetica", 5.5); c.setFillColor(PMGREY)
    c.drawString(rx+6*mm,          y2+1.8*mm, "Reader")
    c.drawRightString(rx+rw-34*mm, y2+1.8*mm, "📧Reach")
    c.drawRightString(rx+rw-24*mm, y2+1.8*mm, "📧Re-op")
    c.drawRightString(rx+rw-14*mm, y2+1.8*mm, "🌐Views")
    c.drawRightString(rx+rw-4*mm,  y2+1.8*mm, "Total")
    y2 -= rh

    for rank, (_, row) in enumerate(data['readers_top'].iterrows(), start=1):
        shade = rank % 2 == 0
        if shade:
            c.setFillColor(PLBLU); c.rect(rx, y2, rw, rh-0.5*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7); c.setFillColor(PBLUE)
        c.drawString(rx+1.5*mm, y2+1.8*mm, str(rank))
        c.setFont("Helvetica", 6.5); c.setFillColor(PDARK)
        name = row['Contact Name']
        c.drawString(rx+6*mm, y2+1.8*mm, name[:24]+'...' if len(name)>24 else name)
        c.setFont("Helvetica", 6); c.setFillColor(colors.HexColor("#555555"))
        c.drawRightString(rx+rw-34*mm, y2+1.8*mm, f"{int(row['Email Reach']):,}")
        c.drawRightString(rx+rw-24*mm, y2+1.8*mm, f"{int(row['Email Re-opens']):,}")
        c.drawRightString(rx+rw-14*mm, y2+1.8*mm, f"{int(row['Portal Views']):,}")
        c.setFont("Helvetica-Bold", 7); c.setFillColor(PNBLUE)
        c.drawRightString(rx+rw-4*mm,  y2+1.8*mm, f"{int(row['Total']):,}")
        y2 -= rh

    c.setFont("Helvetica-Oblique", 6.5); c.setFillColor(PMGREY)
    c.drawString(rx+1.5*mm, y2+1*mm, f"→ See pages 2–{total_pages} for complete reader list")

    draw_footer(1)

    # ══════════════════════════════════════════════════════════════════════════
    # PAGES 2+ — Complete Reader List
    # ══════════════════════════════════════════════════════════════════════════
    full_rw = W - 2*M
    rh2     = 6.2*mm

    for page_idx in range(reader_pages):
        c.showPage()
        page_num = 2 + page_idx
        chunk    = all_readers.iloc[page_idx*ROWS_PER_PAGE : (page_idx+1)*ROWS_PER_PAGE]

        # Mini header
        c.setFillColor(PWHITE); c.rect(0, H-22*mm, W, 22*mm, fill=1, stroke=0)
        c.setFillColor(PBLUE);  c.rect(0, H-1.5*mm, W, 1.5*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 11); c.setFillColor(PDARK)
        c.drawString(M, H-13*mm, f"Complete Reader List  ({len(all_readers)} readers)")
        c.setFont("Helvetica", 8); c.setFillColor(PBLUE)
        c.drawRightString(W-M, H-13*mm, f"{account_name}  |  {data['date_min']} – {data['date_max']}")
        c.setFillColor(PRULE); c.rect(0, H-23*mm, W, 0.5*mm, fill=1, stroke=0)

        # Table header row
        ty2 = H - 30*mm
        c.setFillColor(PBLUE); c.rect(M, ty2, full_rw, 6*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7); c.setFillColor(PWHITE)
        c.drawString(M+10*mm,              ty2+1.5*mm, "READER")
        c.drawRightString(M+full_rw-52*mm, ty2+1.5*mm, "EMAIL REACH")
        c.drawRightString(M+full_rw-36*mm, ty2+1.5*mm, "RE-OPENS")
        c.drawRightString(M+full_rw-20*mm, ty2+1.5*mm, "PORTAL VIEWS")
        c.drawRightString(M+full_rw-1*mm,  ty2+1.5*mm, "TOTAL")

        ry = ty2 - rh2
        for local_i, (_, row) in enumerate(chunk.iterrows()):
            global_rank = page_idx * ROWS_PER_PAGE + local_i + 1
            shade = global_rank % 2 == 0
            if shade:
                c.setFillColor(PLBLU); c.rect(M, ry, full_rw, rh2-0.5*mm, fill=1, stroke=0)
            c.setFont("Helvetica-Bold", 7.5); c.setFillColor(PBLUE)
            c.drawString(M+1.5*mm, ry+1.8*mm, str(global_rank))
            c.setFont("Helvetica", 7.5); c.setFillColor(PDARK)
            name = row['Contact Name']
            c.drawString(M+10*mm, ry+1.8*mm, name[:50]+'...' if len(name)>50 else name)
            c.setFont("Helvetica", 7); c.setFillColor(colors.HexColor("#555555"))
            c.drawRightString(M+full_rw-52*mm, ry+1.8*mm, f"{int(row['Email Reach']):,}")
            c.drawRightString(M+full_rw-36*mm, ry+1.8*mm, f"{int(row['Email Re-opens']):,}")
            c.drawRightString(M+full_rw-20*mm, ry+1.8*mm, f"{int(row['Portal Views']):,}")
            c.setFont("Helvetica-Bold", 7.5); c.setFillColor(PNBLUE)
            c.drawRightString(M+full_rw-1*mm,  ry+1.8*mm, f"{int(row['Total']):,}")
            ry -= rh2

        draw_footer(page_num)

    c.save()
    buf.seek(0)
    return buf.read()

# ── UI ─────────────────────────────────────────────────────────────────────────
cu, cn = st.columns([2, 1])
with cu:
    uploaded = st.file_uploader("Upload analytics file (.xlsx, .xls, .csv)", type=["xlsx","xls","csv"])
with cn:
    account_name = st.text_input("Account / Client name", placeholder="e.g. Richardson Wealth")

if uploaded:
    if not account_name:
        account_name = os.path.splitext(uploaded.name)[0].split('_')[-1]

    with st.spinner("Cleaning and analysing..."):
        df, issues = load_clean(uploaded.read(), uploaded.name)
        data = analyse(df)

    if issues:
        st.markdown(
            f'<div class="quality-box"><strong>⚠️ Data Quality Notes</strong>'
            + ''.join(f'<div>• {i}</div>' for i in issues)
            + '</div>',
            unsafe_allow_html=True)

    # ── Summary line ──────────────────────────────────────────────────────────
    st.markdown(
        f'<div style="background:#fff;border-radius:10px;padding:14px 18px;'
        f'border-left:4px solid {C_BLUE};box-shadow:0 1px 4px rgba(0,119,200,0.08);'
        f'margin-bottom:20px;font-size:0.85rem;color:{C_DARK};">'
        f'<strong style="font-size:1rem;color:{C_BLUE};">{account_name}</strong>'
        f'&ensp;·&ensp;{data["date_min"]} — {data["date_max"]}'
        f'&ensp;·&ensp;{len(data["readers_all"])} readers'
        f'&ensp;·&ensp;{data["email_reach"]:,} email opens'
        f'&ensp;·&ensp;{data["portal_views"]:,} portal views'
        f'&ensp;·&ensp;{data["portal_downloads"]:,} downloads'
        f'</div>',
        unsafe_allow_html=True)

    # ── PDF download ──────────────────────────────────────────────────────────
    with st.spinner("Building PDF..."):
        pdf_bytes = generate_pdf(data, account_name)
    st.markdown('<div class="dl-btn">', unsafe_allow_html=True)
    st.download_button(
        "⬇️  Download PDF Report",
        data=pdf_bytes,
        file_name=f"{account_name.replace(' ', '_')}_Readership_Report.pdf",
        mime="application/pdf")
    st.markdown('</div>', unsafe_allow_html=True)

else:
    st.info("Upload an analytics export above to get started.")
    st.markdown(
        '**Accepted formats:** <span class="chip">.xlsx</span> <span class="chip">.xls</span> <span class="chip">.csv</span><br><br>'
        '**Required columns:** <span class="chip">Contact Name</span> <span class="chip">Contact Email</span> '
        '<span class="chip">EventSource</span> <span class="chip">EventAction</span> <span class="chip">EventDate</span> '
        '<span class="chip">Report Title</span> <span class="chip">Authors</span> <span class="chip">Leaf product</span>',
        unsafe_allow_html=True)
