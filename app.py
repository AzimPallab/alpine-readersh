"""
Alpine Macro — Readership Analytics Web App v5
Run with: streamlit run app.py
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
logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height:44px;">' if logo_b64 else "<strong>ALPINE MACRO</strong>"

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
.am-nav {{ background:#fff; border-bottom:3px solid {C_BLUE}; padding:14px 28px; display:flex; align-items:center; justify-content:space-between; margin-bottom:28px; }}
.am-nav-right {{ font-size:0.72rem; color:#888; text-transform:uppercase; letter-spacing:0.1em; font-weight:500; }}
.kpi-card {{ background:#fff; border-radius:10px; padding:16px 14px 12px; border-left:4px solid var(--ac); box-shadow:0 1px 4px rgba(0,119,200,0.08); }}
.kpi-val {{ font-family:'Source Serif 4',serif; font-size:1.9rem; font-weight:600; color:{C_DARK}; line-height:1; margin-bottom:4px; }}
.kpi-lbl {{ font-size:0.68rem; color:#6B6B6B; text-transform:uppercase; letter-spacing:0.07em; font-weight:500; }}
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
</style>
""", unsafe_allow_html=True)

st.markdown(f'<div class="am-nav">{logo_html}<span class="am-nav-right">Readership Analytics</span></div>', unsafe_allow_html=True)

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

def load_clean(file_bytes, filename):
    issues = []
    try:
        df = pd.read_csv(io.BytesIO(file_bytes)) if filename.endswith('.csv') else pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        st.error(f"Could not read file: {e}"); st.stop()

    df.columns = df.columns.str.strip()
    rmap = {c: ALIASES[c.lower().strip()] for c in df.columns
            if c.lower().strip() in ALIASES and ALIASES[c.lower().strip()] not in df.columns}
    if rmap:
        df = df.rename(columns=rmap)
        issues.append(f"Auto-renamed columns: {', '.join(rmap.keys())}")

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {', '.join(missing)}"); st.stop()

    before = len(df)
    df = df.dropna(how='all')
    if before - len(df):
        issues.append(f"Removed {before - len(df)} blank rows")

    for col in ['Contact Name','Contact Email','EventAction','EventSource','Report Title','Authors','Leaf product']:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({'nan':'', 'NaN':'', 'NULL':'', 'none':'', 'None':''})
        df[col] = df[col].replace('', None)

    before = len(df)
    df = df.dropna(subset=['Contact Name','Contact Email','Report Title','EventAction','EventDate'])
    if before - len(df):
        issues.append(f"Removed {before - len(df)} rows with missing critical fields")

    df['EventAction'] = df['EventAction'].str.lower().str.strip()
    df['EventSource']  = df['EventSource'].str.lower().str.strip()

    df['EventDate'] = pd.to_datetime(df['EventDate'], errors='coerce', infer_datetime_format=True)
    bad = df['EventDate'].isna().sum()
    if bad:
        issues.append(f"Removed {bad} rows with unparseable dates")
        df = df.dropna(subset=['EventDate'])

    df = df[~df['EventAction'].isin(['login'])]
    df = df[~df['EventSource'].isin(['login.oxfordeconomics.com'])]

    before = len(df)
    df = df.drop_duplicates(subset=['Contact Email','EventAction','Report Title','EventDate'])
    if before - len(df):
        issues.append(f"Removed {before - len(df)} duplicate events")

    df['Month'] = df['EventDate'].dt.to_period('M')
    return df, issues

def analyse(df):
    opens  = df[df['EventAction'] == 'initial_open']
    clicks = df[df['EventAction'] == 'click']

    email_op  = opens[opens['EventSource'] == 'email']
    portal_op = opens[opens['EventSource'].isin(['my oxford','myoxford','portal'])]
    email_cl  = clicks[clicks['EventSource'] == 'email']
    portal_cl = clicks[clicks['EventSource'].isin(['my oxford','myoxford','portal'])]

    top_products = (
        email_op[email_op['Leaf product'].notna()]
        ['Leaf product'].str.strip().value_counts().head(6)
    )
    top_reports = (
        email_op[email_op['Report Title'].notna()]
        ['Report Title'].str.strip().value_counts().head(5)
    )
    top_authors = (
        email_op[email_op['Authors'].notna()]
        ['Authors'].str.strip().value_counts().head(5)
    )

    e_reads = email_op.groupby('Contact Name').size().rename('Email Opens')
    p_reads = portal_op.groupby('Contact Name').size().rename('Portal Opens')
    readers = pd.DataFrame({'Email Opens': e_reads, 'Portal Opens': p_reads}).fillna(0).astype(int)
    readers['Total'] = readers['Email Opens'] + readers['Portal Opens']
    readers_all = readers.sort_values('Total', ascending=False).reset_index()
    readers_top = readers_all.head(10).reset_index(drop=True)

    total_opens = len(opens)

    return {
        'date_min':        df['EventDate'].min().strftime('%d %b %Y'),
        'date_max':        df['EventDate'].max().strftime('%d %b %Y'),
        'total_opens':     total_opens,
        'unique_readers':  opens['Contact Name'].nunique(),
        'total_clicks':    len(clicks),
        'total_downloads': len(df[df['EventAction'] == 'downloadproduct']),
        'click_rate':      round(len(clicks) / total_opens * 100, 1) if total_opens else 0,
        'email_opens':     len(email_op),
        'email_clicks':    len(email_cl),
        'portal_opens':    len(portal_op),
        'portal_clicks':   len(portal_cl),
        'top_products':    top_products,
        'top_reports':     top_reports,
        'top_authors':     top_authors,
        'readers_top':     readers_top,
        'readers_all':     readers_all,
    }

def make_channel_chart(data):
    fig, ax = plt.subplots(figsize=(5.5, 2.4))
    categories  = ['Opens', 'Clicks']
    email_vals  = [data['email_opens'],  data['email_clicks']]
    portal_vals = [data['portal_opens'], data['portal_clicks']]
    x, bw = np.arange(2), 0.32

    ax.bar(x - bw/2, email_vals,  bw, color='#0077C8', zorder=3, label='Email')
    ax.bar(x + bw/2, portal_vals, bw, color='#80C4ED', zorder=3, label='Portal')

    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    ax.yaxis.grid(True, color='#CBD2DA', linewidth=0.5, zorder=0); ax.set_axisbelow(True)
    for s in ['top','right','left']: ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color('#E0E8F0')
    ax.tick_params(axis='y', labelsize=7, colors='#555')
    ax.set_xticks(x); ax.set_xticklabels(categories, fontsize=9, color='#333', fontweight='500')

    mx = max(email_vals + portal_vals) or 1
    for vals, offset in [(email_vals, -bw/2), (portal_vals, bw/2)]:
        for xi, v in zip(x + offset, vals):
            if v > 0:
                ax.text(xi, v + mx * 0.02, f'{v:,}', ha='center', va='bottom',
                        fontsize=7, color='#333', fontweight='bold')

    ax.legend(fontsize=7.5, frameon=False, loc='upper right')
    ax.set_title('Email vs Portal Activity', fontsize=9, color='#2B2B2B', pad=8, fontweight='bold', loc='left')
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
    ax.set_title('Opens by Product Category', fontsize=9, color='#2B2B2B',
                 pad=8, fontweight='bold', loc='left')
    fig.tight_layout(pad=0.8)
    return fig

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
           '<span class="t-sub">Email</span>'
           '<span class="t-sub">Portal</span>'
           '<span class="t-num">Total</span>'
           '</div>')
    body = ''
    for rank, (_, row) in enumerate(readers_df.iterrows(), start=1):
        body += (f'<div class="tbl-row">'
                 f'<span class="t-rank">{rank}</span>'
                 f'<span class="t-name">{row["Contact Name"]}</span>'
                 f'<span class="t-sub">{int(row["Email Opens"]):,}</span>'
                 f'<span class="t-sub">{int(row["Portal Opens"]):,}</span>'
                 f'<span class="t-num">{int(row["Total"]):,}</span>'
                 f'</div>')
    return f'<div class="tbl-wrap">{hdr}{body}</div>'

def kpi_card(label, value, ac):
    return (f'<div class="kpi-card" style="--ac:{ac}">'
            f'<div class="kpi-val">{value}</div>'
            f'<div class="kpi-lbl">{label}</div>'
            f'</div>')

PBLUE  = colors.HexColor("#0077C8")
PDARK  = colors.HexColor("#2B2B2B")
PLBLU  = colors.HexColor("#F0F7FD")
PRULE  = colors.HexColor("#D0DCE8")
PWHITE = colors.white
PMGREY = colors.HexColor("#999999")

def fig_to_ir(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=160, bbox_inches='tight', transparent=True)
    buf.seek(0); plt.close(fig)
    return ImageReader(buf)

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
        c.setFont("Helvetica-Bold", 7); c.setFillColor(colors.HexColor("#005A96"))
        c.drawRightString(x+w-1*mm, y+1.8*mm, str(val))
        return y - rh

    # ── PAGE 1 ────────────────────────────────────────────────────────────────
    draw_header()

    # KPI tiles
    kpis = [
        ("Total Opens",    f"{data['total_opens']:,}",    PBLUE),
        ("Unique Readers", f"{data['unique_readers']:,}", PDARK),
        ("Clicks",         f"{data['total_clicks']:,}",   colors.HexColor("#005A96")),
        ("Click Rate",     f"{data['click_rate']}%",      colors.HexColor("#33A1E0")),
        ("Downloads",      f"{data['total_downloads']:,}", PMGREY),
    ]
    bw = (W - 2*M) / len(kpis); by = H-59*mm; bh = 16*mm
    for i, (lbl, val, col) in enumerate(kpis):
        bx = M + i * bw
        c.setFillColor(PLBLU); c.roundRect(bx, by, bw-1.5*mm, bh, 3, fill=1, stroke=0)
        c.setFillColor(col);   c.rect(bx, by, 2*mm, bh, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 13); c.setFillColor(PDARK)
        c.drawCentredString(bx+2*mm+(bw-3.5*mm)/2, by+6.5*mm, val)
        c.setFont("Helvetica", 6.5); c.setFillColor(colors.HexColor("#666666"))
        c.drawCentredString(bx+2*mm+(bw-3.5*mm)/2, by+2.5*mm, lbl.upper())

    # Charts
    cy = H - 122*mm
    c.drawImage(fig_to_ir(make_channel_chart(data)), M, cy,
                width=105*mm, height=54*mm, preserveAspectRatio=True, mask='auto')
    c.drawImage(fig_to_ir(make_donut(data['top_products'])), M+109*mm, cy-14*mm,
                width=76*mm, height=70*mm, preserveAspectRatio=True, mask='auto')

    # Divider
    dy = cy - 8*mm
    c.setStrokeColor(PRULE); c.setLineWidth(0.5)
    c.line(M, dy, W-M, dy)

    # Left column — Top Reports + Top Authors
    rh = 6*mm; ty = dy - 4*mm; y1 = ty

    c.setFillColor(PBLUE); c.rect(M, y1, CW, 5.5*mm, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 7.5); c.setFillColor(PWHITE)
    c.drawString(M+3*mm, y1+1.5*mm, "TOP REPORTS BY OPENS")
    y1 -= rh
    for i, (t, v) in enumerate(data['top_reports'].items()):
        y1 = draw_table_row(M, y1, CW, i+1, t, f"{v:,}", i%2==1, rh)

    y1 -= 2.5*mm
    c.setFillColor(PBLUE); c.rect(M, y1, CW, 5.5*mm, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 7.5); c.setFillColor(PWHITE)
    c.drawString(M+3*mm, y1+1.5*mm, "TOP AUTHORS BY OPENS")
    y1 -= rh
    for i, (a, v) in enumerate(data['top_authors'].items()):
        y1 = draw_table_row(M, y1, CW, i+1, a, f"{v:,}", i%2==1, rh)

    # Right column — Top 10 readers summary
    rx = M + CW + 4*mm; rw = CW; y2 = ty

    c.setFillColor(PBLUE); c.rect(rx, y2, rw, 5.5*mm, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 7.5); c.setFillColor(PWHITE)
    c.drawString(rx+3*mm, y2+1.5*mm, f"MOST ACTIVE READERS  (Top 10 of {len(all_readers)})")
    y2 -= rh

    c.setFont("Helvetica", 6); c.setFillColor(PMGREY)
    c.drawString(rx+6*mm,          y2+1.8*mm, "Reader")
    c.drawRightString(rx+rw-22*mm, y2+1.8*mm, "Email")
    c.drawRightString(rx+rw-11*mm, y2+1.8*mm, "Portal")
    c.drawRightString(rx+rw-1*mm,  y2+1.8*mm, "Total")
    y2 -= rh

    for rank, (_, row) in enumerate(data['readers_top'].iterrows(), start=1):
        shade = rank % 2 == 0
        if shade:
            c.setFillColor(PLBLU); c.rect(rx, y2, rw, rh-0.5*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7); c.setFillColor(PBLUE)
        c.drawString(rx+1.5*mm, y2+1.8*mm, str(rank))
        c.setFont("Helvetica", 7); c.setFillColor(PDARK)
        name = row['Contact Name']
        c.drawString(rx+6*mm, y2+1.8*mm, name[:28]+'...' if len(name)>28 else name)
        c.setFont("Helvetica", 6.5); c.setFillColor(colors.HexColor("#555555"))
        c.drawRightString(rx+rw-22*mm, y2+1.8*mm, f"{int(row['Email Opens']):,}")
        c.drawRightString(rx+rw-11*mm, y2+1.8*mm, f"{int(row['Portal Opens']):,}")
        c.setFont("Helvetica-Bold", 7); c.setFillColor(colors.HexColor("#005A96"))
        c.drawRightString(rx+rw-1*mm,  y2+1.8*mm, f"{int(row['Total']):,}")
        y2 -= rh

    c.setFont("Helvetica-Oblique", 6.5); c.setFillColor(PMGREY)
    c.drawString(rx+1.5*mm, y2+1*mm, f"→ See pages 2–{total_pages} for complete reader list")

    draw_footer(1)

    # ── PAGES 2+ — Complete reader list ──────────────────────────────────────
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

        # Table header
        ty2 = H - 30*mm
        c.setFillColor(PBLUE); c.rect(M, ty2, full_rw, 6*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7.5); c.setFillColor(PWHITE)
        c.drawString(M+10*mm,              ty2+1.5*mm, "READER")
        c.drawRightString(M+full_rw-42*mm, ty2+1.5*mm, "EMAIL")
        c.drawRightString(M+full_rw-22*mm, ty2+1.5*mm, "PORTAL")
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
            c.drawString(M+10*mm, ry+1.8*mm, name[:55]+'...' if len(name)>55 else name)
            c.setFont("Helvetica", 7); c.setFillColor(colors.HexColor("#555555"))
            c.drawRightString(M+full_rw-42*mm, ry+1.8*mm, f"{int(row['Email Opens']):,}")
            c.drawRightString(M+full_rw-22*mm, ry+1.8*mm, f"{int(row['Portal Opens']):,}")
            c.setFont("Helvetica-Bold", 7.5); c.setFillColor(colors.HexColor("#005A96"))
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
            + '</div>', unsafe_allow_html=True)

    st.caption(f"Period: **{data['date_min']} — {data['date_max']}**  ·  Account: **{account_name}**  ·  {len(data['readers_all'])} readers total")

    c1, c2, c3, c4, c5 = st.columns(5)
    for col, lbl, val, ac in [
        (c1, "Total Opens",    f"{data['total_opens']:,}",    C_BLUE),
        (c2, "Unique Readers", f"{data['unique_readers']:,}", C_DARK),
        (c3, "Clicks",         f"{data['total_clicks']:,}",   "#005A96"),
        (c4, "Click Rate",     f"{data['click_rate']}%",      "#33A1E0"),
        (c5, "Downloads",      f"{data['total_downloads']:,}", "#6B6B6B"),
    ]:
        with col: st.markdown(kpi_card(lbl, val, ac), unsafe_allow_html=True)

    st.markdown("")
    st.markdown('<div class="sec-title">Rankings</div>', unsafe_allow_html=True)
    t1, t2, t3 = st.columns(3)
    with t1:
        st.markdown("**Top Reports**")
        st.markdown(render_simple(
            [(t, f"{v:,}") for t, v in data['top_reports'].items()],
            ['#','Report','Opens']), unsafe_allow_html=True)
    with t2:
        st.markdown("**Top Authors**")
        st.markdown(render_simple(
            [(a, f"{v:,}") for a, v in data['top_authors'].items()],
            ['#','Author','Opens']), unsafe_allow_html=True)
    with t3:
        st.markdown(f"**Most Active Readers** *(top 10 of {len(data['readers_all'])})*")
        st.markdown(render_readers(data['readers_top']), unsafe_allow_html=True)

    st.markdown("---")
    st.info(f"📄 PDF includes charts + complete list of all {len(data['readers_all'])} readers.")
    dl, _ = st.columns([1, 3])
    with dl:
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
