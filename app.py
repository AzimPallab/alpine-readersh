"""
Alpine Macro — Readership Analytics Web App v7
Run with: streamlit run app.py

Metrics:
- Email Reach:       initial_open (EventSource = email)
- Portal Views:      OpenLandingPage (EventSource = my oxford / myoxford / portal)
- Portal Deep Reads: OpenChapter
- Portal Downloads:  DownloadProduct
- Most Active Readers: Email Reach + Portal Views + Portal Deep Reads combined
- Top Reports:       Email Reach + Portal Views combined
- Top Authors:       Email Reach only
- Donut chart:       Email Reach by Leaf product
"""

import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import io, os, base64, re
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
logo_html = (f'<img src="data:image/png;base64,{logo_b64}" style="height:64px;">'
             if logo_b64 else "<strong>ALPINE MACRO</strong>")

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
.am-nav {{ background:#fff; border-bottom:3px solid {C_BLUE}; padding:16px 28px;
           display:flex; align-items:center; justify-content:space-between; margin-bottom:28px; }}
.am-nav-right {{ font-size:0.72rem; color:#888; text-transform:uppercase;
                 letter-spacing:0.1em; font-weight:500; }}
.summary-card {{ background:#fff; border-radius:10px; padding:16px 20px;
                 border-left:4px solid {C_BLUE};
                 box-shadow:0 1px 4px rgba(0,119,200,0.08); margin-bottom:22px; }}
.summary-account {{ font-family:'Source Serif 4',serif; font-size:1.15rem;
                    font-weight:600; color:{C_BLUE}; margin-bottom:5px; }}
.summary-meta {{ font-size:0.8rem; color:#666; }}
.summary-meta span {{ color:{C_DARK}; font-weight:500; }}
.dl-btn > button {{ background:{C_DARK} !important; color:#fff !important;
                    border:none !important; border-radius:8px !important;
                    font-weight:600 !important; font-size:0.95rem !important;
                    padding:0.6rem 2.4rem !important; transition:background 0.2s; }}
.dl-btn > button:hover {{ background:{C_BLUE} !important; }}
.chip {{ display:inline-block; background:{C_LBLUE}; color:{C_BLUE}; border-radius:20px;
         padding:2px 10px; font-size:0.72rem; font-weight:500; margin:2px; }}
div[data-testid="stFileUploader"] {{ border:2px dashed {C_BLUE} !important;
                                      border-radius:10px !important;
                                      background:{C_LBLUE} !important; }}
</style>
""", unsafe_allow_html=True)

st.markdown(
    f'<div class="am-nav">{logo_html}'
    f'<span class="am-nav-right">Readership Analytics</span></div>',
    unsafe_allow_html=True)

# ── Column aliases ─────────────────────────────────────────────────────────────
REQUIRED = ['Contact Name','Contact Email','EventSource','EventAction',
            'EventDate','Report Title','Authors','Leaf product']
ALIASES  = {
    'contact name':'Contact Name',   'name':'Contact Name',
    'contact email':'Contact Email', 'email':'Contact Email',
    'eventsource':'EventSource',     'source':'EventSource',   'event source':'EventSource',
    'eventaction':'EventAction',     'action':'EventAction',   'event action':'EventAction',
    'eventdate':'EventDate',         'date':'EventDate',       'event date':'EventDate',
    'report title':'Report Title',   'title':'Report Title',   'report':'Report Title',
    'authors':'Authors',             'author':'Authors',
    'leaf product':'Leaf product',   'product':'Leaf product', 'category':'Leaf product',
}

# ── Event / source constants ───────────────────────────────────────────────────
EA_EMAIL_REACH = 'initial_open'
EA_PORTAL_VIEW = 'openlandingpage'
EA_PORTAL_READ = 'openchapter'
EA_PORTAL_DL   = 'downloadproduct'
SRC_EMAIL      = 'email'
SRC_PORTAL     = ['my oxford', 'myoxford', 'portal']
LOGIN_ACTIONS  = ['login', 'loginfederatedidentity', 'loginipaddress',
                  'logout', 'ssogeorestriction']

# ── Filename → account name ────────────────────────────────────────────────────
def account_from_filename(filename):
    stem = os.path.splitext(filename)[0]
    stem = re.sub(r'[-_]+\d{4}[-_]\d{2}[-_]\d{2}([-_]+\d{2}[-_]\d{2}[-_]\d{2})?', '', stem)
    stem = re.sub(r'[-_]+', ' ', stem).strip()
    return stem if stem else filename

# ── Data loading & cleaning ────────────────────────────────────────────────────
def load_clean(file_bytes, filename):
    try:
        df = (pd.read_csv(io.BytesIO(file_bytes)) if filename.endswith('.csv')
              else pd.read_excel(io.BytesIO(file_bytes)))
    except Exception as e:
        st.error(f"Could not read file: {e}"); st.stop()

    # Normalise column names
    df.columns = df.columns.str.strip()
    rmap = {c: ALIASES[c.lower().strip()] for c in df.columns
            if c.lower().strip() in ALIASES and ALIASES[c.lower().strip()] not in df.columns}
    if rmap:
        df = df.rename(columns=rmap)

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {', '.join(missing)}"); st.stop()

    df = df.dropna(how='all')

    for col in ['Contact Name','Contact Email','EventAction','EventSource',
                'Report Title','Authors','Leaf product']:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({'nan':'', 'NaN':'', 'NULL':'', 'none':'', 'None':''})
        df[col] = df[col].replace('', None)

    df = df.dropna(subset=['Contact Name','Contact Email','Report Title',
                            'EventAction','EventDate'])

    df['EventAction'] = df['EventAction'].str.lower().str.strip()
    df['EventSource']  = df['EventSource'].str.lower().str.strip()

    df['EventDate'] = pd.to_datetime(df['EventDate'], errors='coerce',
                                      infer_datetime_format=True)
    df = df.dropna(subset=['EventDate'])

    df = df[~df['EventAction'].isin(LOGIN_ACTIONS)]
    df = df[~df['EventSource'].isin(['login.oxfordeconomics.com'])]

    df = df.drop_duplicates(subset=['Contact Email','EventAction','Report Title','EventDate'])
    df['Month'] = df['EventDate'].dt.to_period('M')
    return df

# ── Analysis ───────────────────────────────────────────────────────────────────
def analyse(df):
    email_df     = df[df['EventSource'] == SRC_EMAIL]
    email_reach  = email_df[email_df['EventAction'] == EA_EMAIL_REACH]

    portal_df    = df[df['EventSource'].isin(SRC_PORTAL)]
    portal_views = portal_df[portal_df['EventAction'] == EA_PORTAL_VIEW]
    portal_reads = portal_df[portal_df['EventAction'] == EA_PORTAL_READ]
    portal_dls   = portal_df[portal_df['EventAction'] == EA_PORTAL_DL]

    # Top products — email reach by Leaf product
    top_products = (
        email_reach[email_reach['Leaf product'].notna()]
        ['Leaf product'].str.strip().value_counts().head(8)
    )

    # Top reports — email reach + portal views
    top_reports = pd.concat([
        email_reach[email_reach['Report Title'].notna()][['Report Title']],
        portal_views[portal_views['Report Title'].notna()][['Report Title']],
    ])['Report Title'].str.strip().value_counts().head(5)

    # Top authors — email reach only
    top_authors = (
        email_reach[email_reach['Authors'].notna()]
        ['Authors'].str.strip().value_counts().head(5)
    )

    # Most Active Readers — Email Reach + Portal Views + Portal Deep Reads
    e_reach_cnt = email_reach.groupby('Contact Name').size().rename('Email Reach')
    p_view_cnt  = portal_views.groupby('Contact Name').size().rename('Portal Views')
    p_read_cnt  = portal_reads.groupby('Contact Name').size().rename('Portal Deep Reads')

    readers = pd.concat([e_reach_cnt, p_view_cnt, p_read_cnt], axis=1).fillna(0).astype(int)
    readers['Total'] = (readers['Email Reach'] + readers['Portal Views']
                        + readers['Portal Deep Reads'])
    readers_all = readers.sort_values('Total', ascending=False).reset_index()
    readers_top = readers_all.head(10).reset_index(drop=True)

    unique_readers = pd.concat([email_reach, portal_views, portal_reads]
                                )['Contact Name'].nunique()

    return {
        'date_min':            df['EventDate'].min().strftime('%d %b %Y'),
        'date_max':            df['EventDate'].max().strftime('%d %b %Y'),
        'email_reach':         len(email_reach),
        'email_reach_unique':  email_reach['Contact Name'].nunique(),
        'portal_views':        len(portal_views),
        'portal_views_unique': portal_views['Contact Name'].nunique(),
        'portal_reads':        len(portal_reads),
        'portal_downloads':    len(portal_dls),
        'unique_readers':      unique_readers,
        'top_products':        top_products,
        'top_reports':         top_reports,
        'top_authors':         top_authors,
        'readers_top':         readers_top,
        'readers_all':         readers_all,
    }

# ── Charts (PDF) ───────────────────────────────────────────────────────────────
def make_channel_chart(data):
    fig, ax = plt.subplots(figsize=(5.5, 2.6))
    categories = ['Email\nReach', 'Portal\nViews', 'Portal\nDeep Reads', 'Portal\nDownloads']
    vals       = [data['email_reach'], data['portal_views'],
                  data['portal_reads'], data['portal_downloads']]
    bar_colors = [CHART_COLORS[0], CHART_COLORS[1], CHART_COLORS[3], CHART_COLORS[4]]

    x = np.arange(len(categories))
    ax.bar(x, vals, color=bar_colors, zorder=3, width=0.52)
    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    ax.yaxis.grid(True, color='#CBD2DA', linewidth=0.5, zorder=0); ax.set_axisbelow(True)
    for s in ['top','right','left']: ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color('#E0E8F0')
    ax.tick_params(axis='y', labelsize=7, colors='#555')
    ax.set_xticks(x); ax.set_xticklabels(categories, fontsize=7.5, color='#333')

    mx = max(vals) or 1
    for xi, v in zip(x, vals):
        if v > 0:
            ax.text(xi, v + mx*0.02, f'{v:,}', ha='center', va='bottom',
                    fontsize=7, color='#333', fontweight='bold')

    ax.axvline(x=0.5, color='#D0DCE8', linewidth=1, linestyle='--', zorder=2)
    ax.text(0,   ax.get_ylim()[1]*0.96, 'Email',  ha='center',
            fontsize=7, color='#0077C8', fontweight='bold')
    ax.text(2.0, ax.get_ylim()[1]*0.96, 'Portal', ha='center',
            fontsize=7, color='#005A96', fontweight='bold')

    ax.set_title('Engagement by Channel & Event Type', fontsize=9,
                 color='#2B2B2B', pad=8, fontweight='bold', loc='left')
    fig.tight_layout(pad=0.5)
    return fig

def make_donut(top_products):
    if top_products is None or len(top_products) == 0:
        fig, ax = plt.subplots(figsize=(5.5, 4.2))
        ax.text(0.5, 0.5, 'No product data', ha='center', va='center',
                transform=ax.transAxes)
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
    ax.set_title('Email Reach by Product Category', fontsize=9,
                 color='#2B2B2B', pad=8, fontweight='bold', loc='left')
    fig.tight_layout(pad=0.8)
    return fig

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

    all_readers   = data['readers_all']
    ROWS_PER_PAGE = 38
    reader_pages  = max(1, -(-len(all_readers) // ROWS_PER_PAGE))
    total_pages   = 1 + reader_pages

    def draw_header():
        c.setFillColor(PWHITE); c.rect(0, H-38*mm, W, 38*mm, fill=1, stroke=0)
        c.setFillColor(PBLUE);  c.rect(0, H-1.5*mm, W, 1.5*mm, fill=1, stroke=0)
        if os.path.exists(LOGO_PATH):
            c.drawImage(LOGO_PATH, M, H-34*mm, width=54*mm, height=24*mm,
                        preserveAspectRatio=True, mask='auto')
        c.setFont("Helvetica-Bold", 15); c.setFillColor(PDARK)
        c.drawRightString(W-M, H-17*mm, "Readership Analytics Report")
        c.setFont("Helvetica-Bold", 9); c.setFillColor(PBLUE)
        c.drawRightString(W-M, H-24*mm,
                          f"{account_name}  |  {data['date_min']} – {data['date_max']}")
        c.setFont("Helvetica", 7.5); c.setFillColor(PMGREY)
        c.drawRightString(W-M, H-30*mm,
                          f"Generated {datetime.now().strftime('%d %b %Y')}")
        c.setFillColor(PRULE); c.rect(0, H-39*mm, W, 0.5*mm, fill=1, stroke=0)

    def draw_footer(page_num):
        c.setFillColor(PDARK); c.rect(0, 0, W, 9*mm, fill=1, stroke=0)
        c.setFillColor(PBLUE); c.rect(0, 9*mm, W, 0.8*mm, fill=1, stroke=0)
        c.setFont("Helvetica", 6.5); c.setFillColor(colors.HexColor("#AAAAAA"))
        c.drawString(M, 3.2*mm,
                     f"{account_name}  |  Readership Analytics  |  Confidential")
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

    # ── 5 KPI tiles ───────────────────────────────────────────────────────────
    kpis = [
        ("Email Reach",       f"{data['email_reach']:,}",     PBLUE,  "initial_open"),
        ("Unique Readers",    f"{data['unique_readers']:,}",   PDARK,  "all channels"),
        ("Portal Views",      f"{data['portal_views']:,}",     PNBLUE, "OpenLandingPage"),
        ("Portal Deep Reads", f"{data['portal_reads']:,}",     colors.HexColor("#33A1E0"),
                                                                        "OpenChapter"),
        ("Portal Downloads",  f"{data['portal_downloads']:,}", PMGREY, "DownloadProduct"),
    ]
    bw = (W - 2*M) / len(kpis); by = H-59*mm; bh = 17*mm
    for i, (lbl, val, col, hint) in enumerate(kpis):
        bx = M + i * bw
        c.setFillColor(PLBLU); c.roundRect(bx, by, bw-1.5*mm, bh, 3, fill=1, stroke=0)
        c.setFillColor(col);   c.rect(bx, by, 2*mm, bh, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 13); c.setFillColor(PDARK)
        c.drawCentredString(bx+2*mm+(bw-3.5*mm)/2, by+8.5*mm, val)
        c.setFont("Helvetica", 6); c.setFillColor(colors.HexColor("#666666"))
        c.drawCentredString(bx+2*mm+(bw-3.5*mm)/2, by+4.5*mm, lbl.upper())
        c.setFont("Helvetica-Oblique", 5.5); c.setFillColor(PMGREY)
        c.drawCentredString(bx+2*mm+(bw-3.5*mm)/2, by+1.8*mm, hint)

    # ── Channel labels ────────────────────────────────────────────────────────
    sep_y = by - 4*mm
    c.setFillColor(PBLUE);  c.rect(M,       sep_y, bw*1-1.5*mm, 0.8*mm, fill=1, stroke=0)
    c.setFillColor(PNBLUE); c.rect(M+bw*2,  sep_y, bw*3-1.5*mm, 0.8*mm, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 6.5)
    c.setFillColor(PBLUE);  c.drawString(M+1*mm,      sep_y-3.5*mm, "EMAIL CHANNEL")
    c.setFillColor(PNBLUE); c.drawString(M+bw*2+1*mm, sep_y-3.5*mm, "PORTAL CHANNEL")

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

    rh = 6*mm; ty = dy - 4*mm; y1 = ty

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
    c.drawString(rx+3*mm, y2+1.5*mm,
                 f"MOST ACTIVE READERS  (Top 10 of {len(all_readers)})")
    y2 -= rh

    c.setFont("Helvetica", 5.5); c.setFillColor(PMGREY)
    c.drawString(rx+6*mm,          y2+1.8*mm, "Reader")
    c.drawRightString(rx+rw-28*mm, y2+1.8*mm, "Email Reach")
    c.drawRightString(rx+rw-14*mm, y2+1.8*mm, "Portal Views")
    c.drawRightString(rx+rw-2*mm,  y2+1.8*mm, "Total")
    y2 -= rh

    for rank, (_, row) in enumerate(data['readers_top'].iterrows(), start=1):
        shade = rank % 2 == 0
        if shade:
            c.setFillColor(PLBLU); c.rect(rx, y2, rw, rh-0.5*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7); c.setFillColor(PBLUE)
        c.drawString(rx+1.5*mm, y2+1.8*mm, str(rank))
        c.setFont("Helvetica", 6.5); c.setFillColor(PDARK)
        name = row['Contact Name']
        c.drawString(rx+6*mm, y2+1.8*mm, name[:26]+'...' if len(name) > 26 else name)
        c.setFont("Helvetica", 6.5); c.setFillColor(colors.HexColor("#555555"))
        c.drawRightString(rx+rw-28*mm, y2+1.8*mm, f"{int(row['Email Reach']):,}")
        c.drawRightString(rx+rw-14*mm, y2+1.8*mm, f"{int(row['Portal Views']):,}")
        c.setFont("Helvetica-Bold", 7); c.setFillColor(PNBLUE)
        c.drawRightString(rx+rw-2*mm,  y2+1.8*mm, f"{int(row['Total']):,}")
        y2 -= rh

    c.setFont("Helvetica-Oblique", 6.5); c.setFillColor(PMGREY)
    c.drawString(rx+1.5*mm, y2+1*mm,
                 f"→ See pages 2–{total_pages} for complete reader list")

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

        c.setFillColor(PWHITE); c.rect(0, H-22*mm, W, 22*mm, fill=1, stroke=0)
        c.setFillColor(PBLUE);  c.rect(0, H-1.5*mm, W, 1.5*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 11); c.setFillColor(PDARK)
        c.drawString(M, H-13*mm, f"Complete Reader List  ({len(all_readers)} readers)")
        c.setFont("Helvetica-Bold", 8); c.setFillColor(PBLUE)
        c.drawRightString(W-M, H-13*mm,
                          f"{account_name}  |  {data['date_min']} – {data['date_max']}")
        c.setFillColor(PRULE); c.rect(0, H-23*mm, W, 0.5*mm, fill=1, stroke=0)

        ty2 = H - 30*mm
        c.setFillColor(PBLUE); c.rect(M, ty2, full_rw, 6*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7); c.setFillColor(PWHITE)
        c.drawString(M+10*mm,              ty2+1.5*mm, "READER")
        c.drawRightString(M+full_rw-38*mm, ty2+1.5*mm, "EMAIL REACH")
        c.drawRightString(M+full_rw-20*mm, ty2+1.5*mm, "PORTAL VIEWS")
        c.drawRightString(M+full_rw-1*mm,  ty2+1.5*mm, "TOTAL")

        ry = ty2 - rh2
        for local_i, (_, row) in enumerate(chunk.iterrows()):
            global_rank = page_idx * ROWS_PER_PAGE + local_i + 1
            shade = global_rank % 2 == 0
            if shade:
                c.setFillColor(PLBLU)
                c.rect(M, ry, full_rw, rh2-0.5*mm, fill=1, stroke=0)
            c.setFont("Helvetica-Bold", 7.5); c.setFillColor(PBLUE)
            c.drawString(M+1.5*mm, ry+1.8*mm, str(global_rank))
            c.setFont("Helvetica", 7.5); c.setFillColor(PDARK)
            name = row['Contact Name']
            c.drawString(M+10*mm, ry+1.8*mm,
                         name[:52]+'...' if len(name) > 52 else name)
            c.setFont("Helvetica", 7); c.setFillColor(colors.HexColor("#555555"))
            c.drawRightString(M+full_rw-38*mm, ry+1.8*mm, f"{int(row['Email Reach']):,}")
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
    uploaded = st.file_uploader(
        "Upload analytics file (.xlsx, .xls, .csv)", type=["xlsx","xls","csv"])
with cn:
    account_name_input = st.text_input(
        "Account / Client name (optional)", placeholder="e.g. Richardson Wealth")

if uploaded:
    # Account name: manual input takes priority, otherwise derived from filename
    account_name = (account_name_input.strip() if account_name_input.strip()
                    else account_from_filename(uploaded.name))

    with st.spinner("Analysing..."):
        df   = load_clean(uploaded.read(), uploaded.name)
        data = analyse(df)

    # ── Summary card ──────────────────────────────────────────────────────────
    st.markdown(
        f'<div class="summary-card">'
        f'<div class="summary-account">{account_name}</div>'
        f'<div class="summary-meta">'
        f'{data["date_min"]} — {data["date_max"]}'
        f'&emsp;·&emsp;<span>{len(data["readers_all"])} readers</span>'
        f'&emsp;·&emsp;<span>{data["email_reach"]:,} email opens</span>'
        f'&emsp;·&emsp;<span>{data["portal_views"]:,} portal views</span>'
        f'&emsp;·&emsp;<span>{data["portal_downloads"]:,} downloads</span>'
        f'</div></div>',
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
        '**Accepted formats:** <span class="chip">.xlsx</span> '
        '<span class="chip">.xls</span> <span class="chip">.csv</span><br><br>'
        '**Required columns:** <span class="chip">Contact Name</span> '
        '<span class="chip">Contact Email</span> '
        '<span class="chip">EventSource</span> '
        '<span class="chip">EventAction</span> '
        '<span class="chip">EventDate</span> '
        '<span class="chip">Report Title</span> '
        '<span class="chip">Authors</span> '
        '<span class="chip">Leaf product</span>',
        unsafe_allow_html=True)
