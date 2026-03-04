"""
Alpine Macro — Readership Analytics Web App
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
from datetime import datetime

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Alpine Macro — Readership Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Logo as base64 for embedding in HTML ─────────────────────────────────────
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "alpine_logo.png")
def get_logo_b64():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None

logo_b64 = get_logo_b64()
logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height:48px;">' if logo_b64 else "<strong>ALPINE MACRO</strong>"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@400;600&family=Inter:wght@300;400;500;600&display=swap');

  html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
    background: #F4F8FC;
  }}

  /* Top nav bar */
  .am-navbar {{
    background: #fff;
    border-bottom: 3px solid #0077C8;
    padding: 14px 28px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 28px;
    border-radius: 0 0 8px 8px;
  }}
  .am-navbar-right {{
    font-size: 0.78rem;
    color: #6B6B6B;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-weight: 500;
  }}

  /* KPI cards */
  .kpi-grid {{ display: flex; gap: 12px; margin: 16px 0; }}
  .kpi-card {{
    flex: 1;
    background: #fff;
    border-radius: 10px;
    padding: 16px 14px 12px;
    box-shadow: 0 1px 4px rgba(0,119,200,0.08);
    border-left: 4px solid var(--accent, #0077C8);
    min-width: 0;
  }}
  .kpi-value {{
    font-family: 'Source Serif 4', serif;
    font-size: 1.9rem;
    font-weight: 600;
    color: #2B2B2B;
    line-height: 1;
    margin-bottom: 4px;
  }}
  .kpi-label {{
    font-size: 0.68rem;
    color: #6B6B6B;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    font-weight: 500;
  }}

  /* Section headings */
  .section-heading {{
    font-family: 'Source Serif 4', serif;
    font-size: 1.05rem;
    font-weight: 600;
    color: #0077C8;
    border-bottom: 2px solid #0077C8;
    padding-bottom: 5px;
    margin: 24px 0 14px;
    display: inline-block;
  }}

  /* Upload zone */
  div[data-testid="stFileUploader"] {{
    border: 2px dashed #0077C8 !important;
    border-radius: 10px !important;
    background: #F0F7FD !important;
  }}

  /* Download button */
  .dl-btn > button {{
    background: #0077C8 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    padding: 0.55rem 1.8rem !important;
    width: 100%;
    transition: background 0.2s;
  }}
  .dl-btn > button:hover {{ background: #005A96 !important; }}

  /* Tables */
  .stDataFrame thead th {{
    background: #0077C8 !important;
    color: #fff !important;
    font-size: 0.75rem !important;
  }}

  /* Chips */
  .chip {{
    display: inline-block;
    background: #E6F2FB;
    color: #005A96;
    border-radius: 20px;
    padding: 3px 10px;
    font-size: 0.73rem;
    font-weight: 500;
    margin: 2px;
  }}
</style>
""", unsafe_allow_html=True)

# ── Navbar ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="am-navbar">
  {logo_html}
  <span class="am-navbar-right">Readership Analytics</span>
</div>
""", unsafe_allow_html=True)

# ── Brand / PDF colours ───────────────────────────────────────────────────────
ALPINE_BLUE = colors.HexColor("#0077C8")
DARK_GREY   = colors.HexColor("#2B2B2B")
MID_GREY    = colors.HexColor("#6B6B6B")
LIGHT_BG    = colors.HexColor("#F0F7FD")
RULE_GREY   = colors.HexColor("#D0DCE8")
WHITE       = colors.white
CHART_HEX   = ["#0077C8","#005A96","#33A1E0","#80C4ED","#B3DDF5","#E6F4FC"]

# ── Data ──────────────────────────────────────────────────────────────────────
@st.cache_data
def load_and_analyse(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    df['EventDate'] = pd.to_datetime(df['EventDate'])
    df['Month']     = df['EventDate'].dt.to_period('M')
    df = df.drop_duplicates(subset=['Contact Email','EventAction','Report Title','EventDate'])

    opens     = df[df['EventAction'] == 'initial_open']
    clicks    = df[df['EventAction'] == 'click']
    downloads = df[df['EventAction'] == 'DownloadProduct']

    monthly = opens.groupby('Month').size().reset_index(name='Opens')
    monthly['Label'] = monthly['Month'].astype(str).str[-2:].map(
        {'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'May','06':'Jun',
         '07':'Jul','08':'Aug','09':'Sep','10':'Oct','11':'Nov','12':'Dec'})

    top_products = opens['Leaf product'].str.strip().value_counts().head(6)

    tr = opens['Report Title'].str.strip().value_counts().head(5).reset_index()
    tr.columns = ['Report Title','Opens']
    ta = opens['Authors'].str.strip().value_counts().head(5).reset_index()
    ta.columns = ['Author','Opens']
    trd = opens.groupby('Contact Name').size().sort_values(ascending=False).head(10).reset_index()
    trd.columns = ['Reader','Opens']

    return {
        'date_min':       df['EventDate'].min().strftime('%d %b %Y'),
        'date_max':       df['EventDate'].max().strftime('%d %b %Y'),
        'total_opens':    len(opens),
        'unique_readers': df['Contact Name'].nunique(),
        'total_clicks':   len(clicks),
        'total_downloads':len(downloads),
        'click_rate':     round(len(clicks)/len(opens)*100,1) if len(opens) else 0,
        'monthly_opens':  monthly,
        'top_products':   top_products,
        'top_reports':    tr,
        'top_authors':    ta,
        'top_readers':    trd,
    }

# ── Streamlit charts ──────────────────────────────────────────────────────────
def plot_bar(monthly):
    fig, ax = plt.subplots(figsize=(7, 2.8))
    bars = ax.bar(monthly['Label'], monthly['Opens'], color="#0077C8", width=0.55, zorder=3)
    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    ax.yaxis.grid(True, color='#D0DCE8', linewidth=0.6, zorder=0); ax.set_axisbelow(True)
    for s in ['top','right','left']: ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color('#D0DCE8')
    ax.tick_params(axis='y', labelsize=9, colors='#6B6B6B')
    ax.tick_params(axis='x', labelsize=10, colors='#2B2B2B')
    for bar, val in zip(bars, monthly['Opens']):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+15,
                f'{val:,}', ha='center', va='bottom', fontsize=8, fontweight='bold', color='#2B2B2B')
    fig.tight_layout()
    return fig

def plot_donut(top_products):
    fig, ax = plt.subplots(figsize=(5, 4.2))
    labels = [l[:26]+'...' if len(l)>26 else l for l in top_products.index]
    vals   = top_products.values
    ax.pie(vals, labels=None, colors=CHART_HEX, startangle=90, wedgeprops=dict(width=0.55))
    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    patches = [mpatches.Patch(color=CHART_HEX[i], label=f'{labels[i]}  ({vals[i]:,})')
               for i in range(len(labels))]
    ax.legend(handles=patches, loc='lower center', bbox_to_anchor=(0.5,-0.28),
              fontsize=8.5, frameon=False)
    fig.tight_layout()
    return fig

# ── PDF helpers ───────────────────────────────────────────────────────────────
def fig_to_ir(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', transparent=True)
    buf.seek(0); plt.close(fig)
    return ImageReader(buf)

def pdf_bar(monthly):
    fig, ax = plt.subplots(figsize=(5, 2.1))
    bars = ax.bar(monthly['Label'], monthly['Opens'], color="#0077C8", width=0.55, zorder=3)
    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    ax.yaxis.grid(True, color='#D0DCE8', linewidth=0.5, zorder=0); ax.set_axisbelow(True)
    for s in ['top','right','left']: ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color('#D0DCE8')
    ax.tick_params(axis='y', labelsize=7, colors='#6B6B6B')
    ax.tick_params(axis='x', labelsize=8, colors='#2B2B2B')
    for bar, val in zip(bars, monthly['Opens']):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+18,
                f'{val:,}', ha='center', va='bottom', fontsize=7, color='#2B2B2B', fontweight='bold')
    ax.set_title('Monthly Email Opens', fontsize=9, color='#2B2B2B', pad=6, fontweight='bold', loc='left')
    return fig_to_ir(fig)

def pdf_donut(top_products):
    fig, ax = plt.subplots(figsize=(3.2, 2.9))
    labels = [l[:24]+'...' if len(l)>24 else l for l in top_products.index]
    vals   = top_products.values
    ax.pie(vals, labels=None, colors=CHART_HEX, startangle=90, wedgeprops=dict(width=0.55))
    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    patches = [mpatches.Patch(color=CHART_HEX[i], label=f'{labels[i]}  ({vals[i]:,})')
               for i in range(len(labels))]
    ax.legend(handles=patches, loc='lower center', bbox_to_anchor=(0.5,-0.38), fontsize=6.5, frameon=False)
    ax.set_title('Opens by Product Category', fontsize=9, color='#2B2B2B', pad=6, fontweight='bold', loc='left')
    return fig_to_ir(fig)

def generate_pdf(data, account_name):
    buf = io.BytesIO()
    W, H = A4
    c = canvas.Canvas(buf, pagesize=A4)

    # Header
    header_h = 38*mm
    c.setFillColor(WHITE); c.rect(0, H-header_h, W, header_h, fill=1, stroke=0)
    c.setFillColor(ALPINE_BLUE); c.rect(0, H-2*mm, W, 2*mm, fill=1, stroke=0)
    if os.path.exists(LOGO_PATH):
        c.drawImage(LOGO_PATH, 12*mm, H-header_h+6*mm,
                    width=58*mm, height=26*mm, preserveAspectRatio=True, mask='auto')
    c.setFont("Helvetica-Bold", 15); c.setFillColor(DARK_GREY)
    c.drawRightString(W-12*mm, H-18*mm, "Readership Analytics Report")
    c.setFont("Helvetica", 8.5); c.setFillColor(ALPINE_BLUE)
    c.drawRightString(W-12*mm, H-25*mm, f"{account_name}  |  {data['date_min']} – {data['date_max']}")
    c.setFont("Helvetica", 7.5); c.setFillColor(colors.HexColor("#999999"))
    c.drawRightString(W-12*mm, H-31*mm, f"Generated {datetime.now().strftime('%d %b %Y')}")
    c.setFillColor(ALPINE_BLUE); c.rect(0, H-header_h-1*mm, W, 1*mm, fill=1, stroke=0)

    # KPIs
    kpis = [
        ("Total Opens",    f"{data['total_opens']:,}",     ALPINE_BLUE),
        ("Unique Readers", f"{data['unique_readers']:,}",  DARK_GREY),
        ("Clicks",         f"{data['total_clicks']:,}",    colors.HexColor("#005A96")),
        ("Click Rate",     f"{data['click_rate']}%",       colors.HexColor("#33A1E0")),
        ("Downloads",      f"{data['total_downloads']:,}", MID_GREY),
    ]
    box_w = (W-24*mm)/len(kpis); box_y = H-header_h-24*mm; box_h = 18*mm
    for i, (label, value, col) in enumerate(kpis):
        bx = 12*mm+i*box_w
        c.setFillColor(LIGHT_BG); c.roundRect(bx, box_y, box_w-2*mm, box_h, 3, fill=1, stroke=0)
        c.setFillColor(col); c.rect(bx, box_y, 2*mm, box_h, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 14); c.setFillColor(DARK_GREY)
        c.drawCentredString(bx+2*mm+(box_w-4*mm)/2, box_y+7*mm, value)
        c.setFont("Helvetica", 6.5); c.setFillColor(MID_GREY)
        c.drawCentredString(bx+2*mm+(box_w-4*mm)/2, box_y+2.5*mm, label.upper())

    # Charts
    chart_y = box_y-53*mm
    c.drawImage(pdf_bar(data['monthly_opens']),   12*mm,  chart_y,       width=108*mm, height=50*mm, preserveAspectRatio=True, mask='auto')
    c.drawImage(pdf_donut(data['top_products']),  122*mm, chart_y-8*mm,  width=72*mm,  height=60*mm, preserveAspectRatio=True, mask='auto')

    # Divider
    div_y = chart_y-5*mm
    c.setStrokeColor(RULE_GREY); c.setLineWidth(0.5); c.line(12*mm, div_y, W-12*mm, div_y)

    # Tables
    table_top = div_y-3*mm
    col1_x, col2_x = 12*mm, W/2+2*mm
    col_w, row_h   = W/2-16*mm, 6.5*mm

    def hdr(x, y, title):
        c.setFillColor(ALPINE_BLUE); c.rect(x, y, col_w, 5.5*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7.5); c.setFillColor(WHITE)
        c.drawString(x+3*mm, y+1.5*mm, title)
        return y-row_h

    def row(x, y, rank, label, value, shade):
        if shade: c.setFillColor(LIGHT_BG); c.rect(x, y, col_w, row_h-0.5*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7); c.setFillColor(ALPINE_BLUE)
        c.drawString(x+1.5*mm, y+1.8*mm, str(rank))
        c.setFont("Helvetica", 7); c.setFillColor(DARK_GREY)
        c.drawString(x+6.5*mm, y+1.8*mm, label[:40]+'...' if len(label)>40 else label)
        c.setFont("Helvetica-Bold", 7); c.setFillColor(colors.HexColor("#005A96"))
        c.drawRightString(x+col_w-1*mm, y+1.8*mm, f"{value:,}")
        return y-row_h

    y1 = hdr(col1_x, table_top, "TOP REPORTS BY OPENS")
    for i, r in data['top_reports'].iterrows():
        y1 = row(col1_x, y1, i+1, r['Report Title'], r['Opens'], i%2==1)
    y1 -= 2.5*mm
    y1 = hdr(col1_x, y1, "TOP AUTHORS BY OPENS")
    for i, r in data['top_authors'].iterrows():
        y1 = row(col1_x, y1, i+1, r['Author'], r['Opens'], i%2==1)

    y2 = hdr(col2_x, table_top, "MOST ACTIVE READERS")
    for i, r in data['top_readers'].iterrows():
        y2 = row(col2_x, y2, i+1, r['Reader'], r['Opens'], i%2==1)

    # Footer
    c.setFillColor(DARK_GREY); c.rect(0, 0, W, 9*mm, fill=1, stroke=0)
    c.setFillColor(ALPINE_BLUE); c.rect(0, 9*mm, W, 0.8*mm, fill=1, stroke=0)
    c.setFont("Helvetica", 6.5); c.setFillColor(colors.HexColor("#AAAAAA"))
    c.drawString(12*mm, 3*mm, f"{account_name}  |  Readership Analytics  |  Confidential")
    c.drawRightString(W-12*mm, 3*mm, "Alpine Macro — An Oxford Economics Company")

    c.save(); buf.seek(0)
    return buf.read()

# ── UI ────────────────────────────────────────────────────────────────────────
col_up, col_name = st.columns([2, 1])
with col_up:
    uploaded = st.file_uploader("Upload analytics spreadsheet (.xlsx)", type=["xlsx"])
with col_name:
    account_name = st.text_input("Account / Client name", placeholder="e.g. Richardson Wealth")

if uploaded:
    if not account_name:
        account_name = os.path.splitext(uploaded.name)[0].split('_')[-1]

    with st.spinner("Analysing..."):
        data = load_and_analyse(uploaded.read())

    st.caption(f"Period: **{data['date_min']} — {data['date_max']}**  ·  Deduplicated unique events")

    # KPIs
    st.markdown('<div class="section-heading">Key Metrics</div>', unsafe_allow_html=True)
    kpi_items = [
        ("Total Opens",    data['total_opens'],    "#0077C8"),
        ("Unique Readers", data['unique_readers'], "#2B2B2B"),
        ("Clicks",         data['total_clicks'],   "#005A96"),
        ("Click Rate",     f"{data['click_rate']}%", "#33A1E0"),
        ("Downloads",      data['total_downloads'], "#6B6B6B"),
    ]
    cols = st.columns(5)
    for col, (label, value, accent) in zip(cols, kpi_items):
        with col:
            st.markdown(f"""
            <div class="kpi-card" style="--accent:{accent}">
              <div class="kpi-value">{value}</div>
              <div class="kpi-label">{label}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("")

    # Charts
    st.markdown('<div class="section-heading">Trends & Distribution</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([3, 2])
    with c1:
        st.markdown("**Monthly Email Opens**")
        st.pyplot(plot_bar(data['monthly_opens']), use_container_width=True)
    with c2:
        st.markdown("**Opens by Product Category**")
        st.pyplot(plot_donut(data['top_products']), use_container_width=True)

    # Tables
    st.markdown('<div class="section-heading">Rankings</div>', unsafe_allow_html=True)
    t1, t2, t3 = st.columns(3)
    with t1:
        st.markdown("**Top Reports**")
        st.dataframe(data['top_reports'], use_container_width=True, hide_index=True)
    with t2:
        st.markdown("**Top Authors**")
        st.dataframe(data['top_authors'], use_container_width=True, hide_index=True)
    with t3:
        st.markdown("**Most Active Readers**")
        st.dataframe(data['top_readers'], use_container_width=True, hide_index=True)

    # Export
    st.markdown("---")
    st.markdown('<div class="section-heading">Export PDF</div>', unsafe_allow_html=True)
    dl_col, _ = st.columns([1, 3])
    with dl_col:
        with st.spinner("Building PDF..."):
            pdf_bytes = generate_pdf(data, account_name)
        filename = f"{account_name.replace(' ','_')}_Readership_Report.pdf"
        st.markdown('<div class="dl-btn">', unsafe_allow_html=True)
        st.download_button(
            label="⬇️  Download PDF Report",
            data=pdf_bytes,
            file_name=filename,
            mime="application/pdf",
        )
        st.markdown('</div>', unsafe_allow_html=True)

else:
    st.markdown('<div class="section-heading">Get Started</div>', unsafe_allow_html=True)
    st.info("Upload an `.xlsx` analytics export above to generate your report.")
    st.markdown("""
    **Expected columns:**
    <span class="chip">Contact Name</span>
    <span class="chip">Contact Email</span>
    <span class="chip">EventAction</span>
    <span class="chip">EventDate</span>
    <span class="chip">Report Title</span>
    <span class="chip">Authors</span>
    <span class="chip">Leaf product</span>
    """, unsafe_allow_html=True)
