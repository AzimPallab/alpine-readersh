"""
Alpine Macro — Readership Analytics Web App v3
Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import io, os, base64
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime, timedelta

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
CHART_COLORS = ["#0077C8","#005A96","#33A1E0","#80C4ED","#B3DDF5","#E6F4FC"]

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

# ── Column aliases ─────────────────────────────────────────────────────────────
REQUIRED = ['Contact Name','Contact Email','EventSource','EventAction','EventDate','Report Title','Authors','Leaf product']
ALIASES  = {
    'contact name':'Contact Name','name':'Contact Name',
    'contact email':'Contact Email','email':'Contact Email',
    'eventsource':'EventSource','source':'EventSource','event source':'EventSource',
    'eventaction':'EventAction','action':'EventAction','event action':'EventAction',
    'eventdate':'EventDate','date':'EventDate','event date':'EventDate',
    'report title':'Report Title','title':'Report Title','report':'Report Title',
    'authors':'Authors','author':'Authors',
    'leaf product':'Leaf product','product':'Leaf product','category':'Leaf product',
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
    if rmap: df = df.rename(columns=rmap); issues.append(f"Auto-renamed: {', '.join(rmap.keys())}")

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing: st.error(f"Missing columns: {', '.join(missing)}"); st.stop()

    before = len(df); df = df.dropna(how='all')
    if before-len(df): issues.append(f"Removed {before-len(df)} blank rows")

    for col in ['Contact Name','Contact Email','EventAction','EventSource','Report Title','Authors','Leaf product']:
        df[col] = df[col].astype(str).str.strip()
    df['EventAction'] = df['EventAction'].str.lower()
    df['EventSource']  = df['EventSource'].str.lower()

    df['EventDate'] = pd.to_datetime(df['EventDate'], errors='coerce', infer_datetime_format=True)
    bad = df['EventDate'].isna().sum()
    if bad: issues.append(f"Removed {bad} unparseable dates"); df = df.dropna(subset=['EventDate'])

    df = df[~df['EventAction'].isin(['login'])]
    df = df[~df['EventSource'].isin(['login.oxfordeconomics.com'])]

    before = len(df)
    df = df.drop_duplicates(subset=['Contact Email','EventAction','Report Title','EventDate'])
    if before-len(df): issues.append(f"Removed {before-len(df)} duplicate events")

    df['Month'] = df['EventDate'].dt.to_period('M')
    return df, issues

def analyse(df):
    opens  = df[df['EventAction']=='initial_open']
    clicks = df[df['EventAction']=='click']

    email_op  = opens[opens['EventSource']=='email']
    portal_op = opens[opens['EventSource'].isin(['my oxford','myoxford','portal'])]

    # Monthly opens (email only for trend chart)
    monthly = email_op.groupby('Month').size().reset_index(name='Opens')
    monthly['Label'] = monthly['Month'].astype(str).str[-2:].map(
        {'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'May','06':'Jun',
         '07':'Jul','08':'Aug','09':'Sep','10':'Oct','11':'Nov','12':'Dec'})

    # Top products
    top_products = email_op['Leaf product'].str.strip().value_counts().head(6)

    # Top reports (email opens)
    top_reports = email_op['Report Title'].str.strip().value_counts().head(5)

    # Top authors (email opens)
    top_authors = email_op['Authors'].str.strip().value_counts().head(5)

    # Most active readers — email + portal opens side by side
    e_reads = email_op.groupby('Contact Name').size().rename('Email Opens')
    p_reads = portal_op.groupby('Contact Name').size().rename('Portal Opens')
    readers = pd.DataFrame({'Email Opens': e_reads, 'Portal Opens': p_reads}).fillna(0).astype(int)
    readers['Total'] = readers['Email Opens'] + readers['Portal Opens']
    readers = readers.sort_values('Total', ascending=False).head(10).reset_index()

    return {
        'date_min':        df['EventDate'].min().strftime('%d %b %Y'),
        'date_max':        df['EventDate'].max().strftime('%d %b %Y'),
        'total_opens':     len(opens),
        'unique_readers':  opens['Contact Name'].nunique(),
        'total_clicks':    len(clicks),
        'total_downloads': len(df[df['EventAction']=='downloadproduct']),
        'click_rate':      round(len(clicks)/len(opens)*100,1) if len(opens) else 0,
        'monthly':         monthly,
        'top_products':    top_products,
        'top_reports':     top_reports,
        'top_authors':     top_authors,
        'readers':         readers,
        'email_opens':     len(email_op),
        'portal_opens':    len(portal_op),
    }

# ── Charts — original style ───────────────────────────────────────────────────
def make_bar(monthly):
    fig, ax = plt.subplots(figsize=(5, 2.2))
    labels, values = monthly['Label'].tolist(), monthly['Opens'].tolist()
    bars = ax.bar(labels, values, color=C_BLUE, width=0.55, zorder=3)
    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    ax.yaxis.grid(True, color='#CBD2DA', linewidth=0.5, zorder=0); ax.set_axisbelow(True)
    for s in ['top','right','left']: ax.spines[s].set_visible(False)
    ax.tick_params(axis='y', labelsize=7, colors='#555')
    ax.tick_params(axis='x', labelsize=8, colors='#333')
    mx = max(values) if values else 1
    for bar, val in zip(bars, values):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+mx*0.03,
                f'{val:,}', ha='center', va='bottom', fontsize=7, color='#333', fontweight='bold')
    ax.set_title('Monthly Email Opens', fontsize=9, color='#2B2B2B', pad=6, fontweight='bold')
    fig.tight_layout(); return fig

def make_donut(top_products):
    fig, ax = plt.subplots(figsize=(3.2, 2.8))
    labels = [l[:22]+'...' if len(l)>22 else l for l in top_products.index]
    vals   = top_products.values
    ax.pie(vals, labels=None, colors=CHART_COLORS, startangle=90, wedgeprops=dict(width=0.55))
    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    patches = [mpatches.Patch(color=CHART_COLORS[i], label=f'{labels[i]}  ({vals[i]:,})')
               for i in range(len(labels))]
    ax.legend(handles=patches, loc='lower center', bbox_to_anchor=(0.5,-0.35),
              fontsize=6.5, frameon=False, ncol=1)
    ax.set_title('Opens by Product Category', fontsize=9, color='#2B2B2B', pad=6, fontweight='bold')
    fig.tight_layout(); return fig

# ── Table renderers ───────────────────────────────────────────────────────────
def render_simple(rows, cols):
    hdr = '<div class="tbl-hdr">' + ''.join([
        f'<span class="t-rank">#</span>' if i==0 else
        f'<span class="t-name">{c}</span>' if i==1 else
        f'<span class="t-num">{c}</span>'
        for i,c in enumerate(cols)]) + '</div>'
    body = ""
    for i, row in enumerate(rows):
        v = list(row)
        cells = f'<span class="t-rank">{i+1}</span>'
        cells += f'<span class="t-name">{v[0]}</span>'
        for x in v[1:]: cells += f'<span class="t-num">{x}</span>'
        body += f'<div class="tbl-row">{cells}</div>'
    return f'<div class="tbl-wrap">{hdr}{body}</div>'

def render_readers(readers_df):
    hdr = '<div class="tbl-hdr"><span class="t-rank">#</span><span class="t-name">Reader</span><span class="t-sub">Email</span><span class="t-sub">Portal</span><span class="t-num">Total</span></div>'
    body = ""
    for i, row in readers_df.iterrows():
        body += f'''<div class="tbl-row">
            <span class="t-rank">{i+1}</span>
            <span class="t-name">{row["Contact Name"]}</span>
            <span class="t-sub">{int(row["Email Opens"]):,}</span>
            <span class="t-sub">{int(row["Portal Opens"]):,}</span>
            <span class="t-num">{int(row["Total"]):,}</span>
        </div>'''
    return f'<div class="tbl-wrap">{hdr}{body}</div>'

def kpi_card(label, value, ac):
    return f'<div class="kpi-card" style="--ac:{ac}"><div class="kpi-val">{value}</div><div class="kpi-lbl">{label}</div></div>'

# ── PDF ───────────────────────────────────────────────────────────────────────
PBLUE=colors.HexColor("#0077C8"); PDARK=colors.HexColor("#2B2B2B")
PLBLU=colors.HexColor("#F0F7FD"); PGREY=colors.HexColor("#F4F8FC")
PRULE=colors.HexColor("#D0DCE8"); PWHITE=colors.white; PMGREY=colors.HexColor("#999999")

def fig_to_ir(fig):
    buf=io.BytesIO(); fig.savefig(buf,format='png',dpi=150,bbox_inches='tight',transparent=True)
    buf.seek(0); plt.close(fig); return ImageReader(buf)

def generate_pdf(data, account_name):
    buf=io.BytesIO(); W,H=A4; c=canvas.Canvas(buf,pagesize=A4)
    M=12*mm; CW=(W-2*M-4*mm)/2

    # Header
    c.setFillColor(PWHITE); c.rect(0,H-38*mm,W,38*mm,fill=1,stroke=0)
    c.setFillColor(PBLUE);  c.rect(0,H-1.5*mm,W,1.5*mm,fill=1,stroke=0)
    if os.path.exists(LOGO_PATH):
        c.drawImage(LOGO_PATH,M,H-34*mm,width=54*mm,height=24*mm,preserveAspectRatio=True,mask='auto')
    c.setFont("Helvetica-Bold",15); c.setFillColor(PDARK)
    c.drawRightString(W-M,H-17*mm,"Readership Analytics Report")
    c.setFont("Helvetica",8.5); c.setFillColor(PBLUE)
    c.drawRightString(W-M,H-24*mm,f"{account_name}  |  {data['date_min']} – {data['date_max']}")
    c.setFont("Helvetica",7.5); c.setFillColor(PMGREY)
    c.drawRightString(W-M,H-30*mm,f"Generated {datetime.now().strftime('%d %b %Y')}")
    c.setFillColor(PRULE); c.rect(0,H-39*mm,W,0.5*mm,fill=1,stroke=0)

    # KPI tiles
    kpis = [
        ("Total Opens",    f"{data['total_opens']:,}",    PBLUE),
        ("Unique Readers", f"{data['unique_readers']:,}", PDARK),
        ("Clicks",         f"{data['total_clicks']:,}",   colors.HexColor("#005A96")),
        ("Click Rate",     f"{data['click_rate']}%",      colors.HexColor("#33A1E0")),
        ("Downloads",      f"{data['total_downloads']:,}", PMGREY),
    ]
    bw=(W-2*M)/len(kpis); by=H-59*mm; bh=16*mm
    for i,(lbl,val,col) in enumerate(kpis):
        bx=M+i*bw
        c.setFillColor(PLBLU); c.roundRect(bx,by,bw-1.5*mm,bh,3,fill=1,stroke=0)
        c.setFillColor(col);   c.rect(bx,by,2*mm,bh,fill=1,stroke=0)
        c.setFont("Helvetica-Bold",13); c.setFillColor(PDARK)
        c.drawCentredString(bx+2*mm+(bw-3.5*mm)/2,by+6.5*mm,val)
        c.setFont("Helvetica",6.5); c.setFillColor(colors.HexColor("#666666"))
        c.drawCentredString(bx+2*mm+(bw-3.5*mm)/2,by+2.5*mm,lbl.upper())

    # Charts
    cy=H-115*mm
    c.drawImage(fig_to_ir(make_bar(data['monthly'])),   M,cy,width=108*mm,height=50*mm,preserveAspectRatio=True,mask='auto')
    c.drawImage(fig_to_ir(make_donut(data['top_products'])),M+112*mm,cy-6*mm,width=72*mm,height=58*mm,preserveAspectRatio=True,mask='auto')

    # Divider
    dy=cy-6*mm
    c.setStrokeColor(PRULE); c.setLineWidth(0.5); c.line(M,dy,W-M,dy)

    # Tables
    ty=dy-4*mm; col_w=W/2-16*mm; rh=6*mm

    def shdr(x,y,title,col):
        c.setFillColor(col); c.rect(x,y,col_w,5.5*mm,fill=1,stroke=0)
        c.setFont("Helvetica-Bold",7.5); c.setFillColor(PWHITE)
        c.drawString(x+3*mm,y+1.5*mm,title)
        return y-rh

    def trow(x,y,rank,label,val,shade):
        if shade: c.setFillColor(PLBLU); c.rect(x,y,col_w,rh-0.5*mm,fill=1,stroke=0)
        c.setFont("Helvetica-Bold",7); c.setFillColor(PBLUE); c.drawString(x+1.5*mm,y+1.8*mm,str(rank))
        c.setFont("Helvetica",7); c.setFillColor(PDARK)
        c.drawString(x+6*mm,y+1.8*mm,label[:40]+'...' if len(label)>40 else label)
        c.setFont("Helvetica-Bold",7); c.setFillColor(colors.HexColor("#005A96"))
        c.drawRightString(x+col_w-1*mm,y+1.8*mm,str(val))
        return y-rh

    # Left: Top Reports + Top Authors
    y1=shdr(M,ty,"TOP REPORTS BY OPENS",PBLUE)
    for i,(t,v) in enumerate(data['top_reports'].items()): y1=trow(M,y1,i+1,t,f"{v:,}",i%2==1)
    y1-=2.5*mm
    y1=shdr(M,y1,"TOP AUTHORS BY OPENS",PBLUE)
    for i,(a,v) in enumerate(data['top_authors'].items()): y1=trow(M,y1,i+1,a,f"{v:,}",i%2==1)

    # Right: Most Active Readers with email/portal split
    rx=M+CW+4*mm; rw=col_w
    y2=ty
    c.setFillColor(PBLUE); c.rect(rx,y2,rw,5.5*mm,fill=1,stroke=0)
    c.setFont("Helvetica-Bold",7.5); c.setFillColor(PWHITE)
    c.drawString(rx+3*mm,y2+1.5*mm,"MOST ACTIVE READERS")
    y2-=rh
    # Sub-header
    c.setFont("Helvetica",6); c.setFillColor(PMGREY)
    c.drawString(rx+6*mm,y2+1.8*mm,"Reader")
    c.drawRightString(rx+rw-22*mm,y2+1.8*mm,"Email")
    c.drawRightString(rx+rw-11*mm,y2+1.8*mm,"Portal")
    c.drawRightString(rx+rw-1*mm,y2+1.8*mm,"Total")
    y2-=rh

    for i, row in data['readers'].iterrows():
        if i%2==1: c.setFillColor(PLBLU); c.rect(rx,y2,rw,rh-0.5*mm,fill=1,stroke=0)
        c.setFont("Helvetica-Bold",7); c.setFillColor(PBLUE); c.drawString(rx+1.5*mm,y2+1.8*mm,str(i+1))
        c.setFont("Helvetica",7); c.setFillColor(PDARK)
        name = row['Contact Name']
        c.drawString(rx+6*mm,y2+1.8*mm,name[:26]+'...' if len(name)>26 else name)
        c.setFont("Helvetica",6.5); c.setFillColor(colors.HexColor("#555555"))
        c.drawRightString(rx+rw-22*mm,y2+1.8*mm,f"{int(row['Email Opens']):,}")
        c.drawRightString(rx+rw-11*mm,y2+1.8*mm,f"{int(row['Portal Opens']):,}")
        c.setFont("Helvetica-Bold",7); c.setFillColor(colors.HexColor("#005A96"))
        c.drawRightString(rx+rw-1*mm,y2+1.8*mm,f"{int(row['Total']):,}")
        y2-=rh

    # Footer
    c.setFillColor(PDARK); c.rect(0,0,W,9*mm,fill=1,stroke=0)
    c.setFillColor(PBLUE); c.rect(0,9*mm,W,0.8*mm,fill=1,stroke=0)
    c.setFont("Helvetica",6.5); c.setFillColor(colors.HexColor("#AAAAAA"))
    c.drawString(M,3.2*mm,f"{account_name}  |  Readership Analytics  |  Confidential")
    c.drawRightString(W-M,3.2*mm,"Alpine Macro — An Oxford Economics Company")

    c.save(); buf.seek(0)
    return buf.read()

# ── UI ─────────────────────────────────────────────────────────────────────────
cu, cn = st.columns([2,1])
with cu: uploaded = st.file_uploader("Upload analytics file (.xlsx, .xls, .csv)", type=["xlsx","xls","csv"])
with cn: account_name = st.text_input("Account / Client name", placeholder="e.g. Richardson Wealth")

if uploaded:
    if not account_name:
        account_name = os.path.splitext(uploaded.name)[0].split('_')[-1]

    with st.spinner("Cleaning and analysing..."):
        df, issues = load_clean(uploaded.read(), uploaded.name)
        data = analyse(df)

    if issues:
        st.markdown(f'<div class="quality-box"><strong>⚠️ Data Quality Notes</strong>{"".join(f"<div>• {i}</div>" for i in issues)}</div>', unsafe_allow_html=True)

    st.caption(f"Period: **{data['date_min']} — {data['date_max']}**  ·  Account: **{account_name}**")

    # KPIs
    st.markdown('<div class="sec-title">Key Metrics</div>', unsafe_allow_html=True)
    c1,c2,c3,c4,c5 = st.columns(5)
    for col,lbl,val,ac in [
        (c1,"Total Opens",    f"{data['total_opens']:,}",    C_BLUE),
        (c2,"Unique Readers", f"{data['unique_readers']:,}", C_DARK),
        (c3,"Clicks",         f"{data['total_clicks']:,}",   "#005A96"),
        (c4,"Click Rate",     f"{data['click_rate']}%",      "#33A1E0"),
        (c5,"Downloads",      f"{data['total_downloads']:,}", "#6B6B6B"),
    ]:
        with col: st.markdown(kpi_card(lbl,val,ac), unsafe_allow_html=True)

    st.markdown("")

    # Charts
    st.markdown('<div class="sec-title">Trends & Distribution</div>', unsafe_allow_html=True)
    ch1, ch2 = st.columns([3,2])
    with ch1: st.pyplot(make_bar(data['monthly']), use_container_width=True)
    with ch2: st.pyplot(make_donut(data['top_products']), use_container_width=True)

    # Tables
    st.markdown('<div class="sec-title">Rankings</div>', unsafe_allow_html=True)
    t1,t2,t3 = st.columns(3)
    with t1:
        st.markdown("**Top Reports**")
        rows = [(t,f"{v:,}") for t,v in data['top_reports'].items()]
        st.markdown(render_simple(rows,['#','Report','Opens']), unsafe_allow_html=True)
    with t2:
        st.markdown("**Top Authors**")
        rows = [(a,f"{v:,}") for a,v in data['top_authors'].items()]
        st.markdown(render_simple(rows,['#','Author','Opens']), unsafe_allow_html=True)
    with t3:
        st.markdown("**Most Active Readers**")
        st.markdown(render_readers(data['readers']), unsafe_allow_html=True)

    # Export
    st.markdown("---")
    dl,_ = st.columns([1,3])
    with dl:
        with st.spinner("Building PDF..."): pdf_bytes = generate_pdf(data, account_name)
        st.markdown('<div class="dl-btn">', unsafe_allow_html=True)
        st.download_button("⬇️  Download PDF Report", data=pdf_bytes,
                           file_name=f"{account_name.replace(' ','_')}_Readership_Report.pdf",
                           mime="application/pdf")
        st.markdown('</div>', unsafe_allow_html=True)

else:
    st.info("Upload an analytics export above to get started.")
    st.markdown("""**Accepted formats:** <span class="chip">.xlsx</span> <span class="chip">.xls</span> <span class="chip">.csv</span><br><br>
    **Required columns:** <span class="chip">Contact Name</span> <span class="chip">Contact Email</span>
    <span class="chip">EventSource</span> <span class="chip">EventAction</span> <span class="chip">EventDate</span>
    <span class="chip">Report Title</span> <span class="chip">Authors</span> <span class="chip">Leaf product</span>
    """, unsafe_allow_html=True)
