"""
Alpine Macro — Readership Analytics Report Generator
=====================================================
Usage:
    python readership_report.py <spreadsheet.xlsx> [account_name] [output.pdf]

Examples:
    python readership_report.py Analytics_Richardson.xlsx
    python readership_report.py Analytics_Richardson.xlsx "Richardson Wealth"
    python readership_report.py Analytics_Richardson.xlsx "Richardson Wealth" output/report.pdf

Requirements:
    pip install pandas openpyxl reportlab matplotlib pillow
"""

import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import io, sys, os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime

# ── Brand palette ─────────────────────────────────────────────────────────────
ALPINE_BLUE  = colors.HexColor("#0077C8")   # Alpine Macro brand blue
DARK_GREY    = colors.HexColor("#2B2B2B")   # Near-black from logo text
MID_GREY     = colors.HexColor("#6B6B6B")   # "An Oxford Economics Company"
LIGHT_BG     = colors.HexColor("#F0F7FD")   # Very light tint of brand blue
RULE_GREY    = colors.HexColor("#D0DCE8")
WHITE        = colors.white

CHART_COLORS = ["#0077C8","#005A96","#33A1E0","#80C4ED","#B3DDF5","#E6F4FC"]

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "alpine_logo.png")

# ── Data ──────────────────────────────────────────────────────────────────────
def load_and_analyse(filepath):
    df = pd.read_excel(filepath)
    df['EventDate'] = pd.to_datetime(df['EventDate'])
    df['Month']     = df['EventDate'].dt.to_period('M')
    df = df.drop_duplicates(subset=['Contact Email', 'EventAction', 'Report Title', 'EventDate'])

    opens     = df[df['EventAction'] == 'initial_open']
    clicks    = df[df['EventAction'] == 'click']
    downloads = df[df['EventAction'] == 'DownloadProduct']

    monthly_opens = opens.groupby('Month').size().reset_index(name='Opens')
    monthly_opens['Label'] = monthly_opens['Month'].astype(str).str[-2:].map(
        {'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'May','06':'Jun',
         '07':'Jul','08':'Aug','09':'Sep','10':'Oct','11':'Nov','12':'Dec'})

    return {
        'date_min':       df['EventDate'].min().strftime('%d %b %Y'),
        'date_max':       df['EventDate'].max().strftime('%d %b %Y'),
        'total_opens':    len(opens),
        'unique_readers': df['Contact Name'].nunique(),
        'total_clicks':   len(clicks),
        'total_downloads':len(downloads),
        'click_rate':     round(len(clicks)/len(opens)*100, 1) if len(opens) else 0,
        'monthly_opens':  monthly_opens,
        'top_products':   opens['Leaf product'].str.strip().value_counts().head(6),
        'top_reports':    opens['Report Title'].str.strip().value_counts().head(5),
        'top_authors':    opens['Authors'].str.strip().value_counts().head(5),
        'top_readers':    opens.groupby('Contact Name').size().sort_values(ascending=False).head(10),
    }

# ── Charts ────────────────────────────────────────────────────────────────────
def fig_to_image(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', transparent=True)
    buf.seek(0); plt.close(fig)
    return ImageReader(buf)

def make_bar_chart(monthly_opens):
    fig, ax = plt.subplots(figsize=(5, 2.1))
    labels, values = monthly_opens['Label'].tolist(), monthly_opens['Opens'].tolist()
    bars = ax.bar(labels, values, color="#0077C8", width=0.55, zorder=3)
    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    ax.yaxis.grid(True, color='#D0DCE8', linewidth=0.5, zorder=0); ax.set_axisbelow(True)
    for s in ['top','right','left']: ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color('#D0DCE8')
    ax.tick_params(axis='y', labelsize=7, colors='#6B6B6B')
    ax.tick_params(axis='x', labelsize=8, colors='#2B2B2B')
    for bar, val in zip(bars, values):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+18,
                f'{val:,}', ha='center', va='bottom', fontsize=7, color='#2B2B2B', fontweight='bold')
    ax.set_title('Monthly Email Opens', fontsize=9, color='#2B2B2B', pad=6, fontweight='bold', loc='left')
    return fig_to_image(fig)

def make_donut_chart(top_products):
    fig, ax = plt.subplots(figsize=(3.2, 2.9))
    labels = [l[:24]+'...' if len(l)>24 else l for l in top_products.index]
    vals   = top_products.values
    ax.pie(vals, labels=None, colors=CHART_COLORS, startangle=90, wedgeprops=dict(width=0.55))
    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    patches = [mpatches.Patch(color=CHART_COLORS[i], label=f'{labels[i]}  ({vals[i]:,})')
               for i in range(len(labels))]
    ax.legend(handles=patches, loc='lower center', bbox_to_anchor=(0.5,-0.38),
              fontsize=6.5, frameon=False, ncol=1)
    ax.set_title('Opens by Product Category', fontsize=9, color='#2B2B2B', pad=6, fontweight='bold', loc='left')
    return fig_to_image(fig)

# ── PDF ───────────────────────────────────────────────────────────────────────
def build_pdf(data, account_name, output_path):
    W, H = A4
    c = canvas.Canvas(output_path, pagesize=A4)

    # ── Header: white with logo + blue rule ───────────────────────────────────
    header_h = 38*mm
    c.setFillColor(WHITE)
    c.rect(0, H-header_h, W, header_h, fill=1, stroke=0)

    # Blue top rule
    c.setFillColor(ALPINE_BLUE)
    c.rect(0, H-2*mm, W, 2*mm, fill=1, stroke=0)

    # Logo
    if os.path.exists(LOGO_PATH):
        logo_w, logo_h = 58*mm, 26*mm
        c.drawImage(LOGO_PATH, 12*mm, H-header_h+6*mm,
                    width=logo_w, height=logo_h,
                    preserveAspectRatio=True, mask='auto')

    # Report title (right side of header)
    c.setFont("Helvetica-Bold", 15)
    c.setFillColor(DARK_GREY)
    c.drawRightString(W-12*mm, H-18*mm, "Readership Analytics Report")
    c.setFont("Helvetica", 8.5)
    c.setFillColor(ALPINE_BLUE)
    c.drawRightString(W-12*mm, H-25*mm, f"{account_name}  |  {data['date_min']} – {data['date_max']}")
    c.setFont("Helvetica", 7.5)
    c.setFillColor(colors.HexColor("#999999"))
    c.drawRightString(W-12*mm, H-31*mm, f"Generated {datetime.now().strftime('%d %b %Y')}")

    # Blue rule under header
    c.setFillColor(ALPINE_BLUE)
    c.rect(0, H-header_h-1*mm, W, 1*mm, fill=1, stroke=0)

    # ── KPI tiles ─────────────────────────────────────────────────────────────
    kpis = [
        ("Total Opens",    f"{data['total_opens']:,}",    ALPINE_BLUE),
        ("Unique Readers", f"{data['unique_readers']:,}", DARK_GREY),
        ("Clicks",         f"{data['total_clicks']:,}",   colors.HexColor("#005A96")),
        ("Click Rate",     f"{data['click_rate']}%",      colors.HexColor("#33A1E0")),
        ("Downloads",      f"{data['total_downloads']:,}", colors.HexColor("#6B6B6B")),
    ]
    box_w = (W-24*mm)/len(kpis)
    box_y = H-header_h-24*mm
    box_h = 18*mm
    for i, (label, value, col) in enumerate(kpis):
        bx = 12*mm + i*box_w
        c.setFillColor(LIGHT_BG)
        c.roundRect(bx, box_y, box_w-2*mm, box_h, 3, fill=1, stroke=0)
        # Colour left accent bar
        c.setFillColor(col)
        c.rect(bx, box_y, 2*mm, box_h, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 14)
        c.setFillColor(DARK_GREY)
        c.drawCentredString(bx+2*mm+(box_w-4*mm)/2, box_y+7*mm, value)
        c.setFont("Helvetica", 6.5)
        c.setFillColor(MID_GREY)
        c.drawCentredString(bx+2*mm+(box_w-4*mm)/2, box_y+2.5*mm, label.upper())

    # ── Charts ────────────────────────────────────────────────────────────────
    chart_top = box_y-3*mm
    chart_y   = chart_top-50*mm
    c.drawImage(make_bar_chart(data['monthly_opens']),
                12*mm, chart_y, width=108*mm, height=50*mm,
                preserveAspectRatio=True, mask='auto')
    c.drawImage(make_donut_chart(data['top_products']),
                122*mm, chart_y-8*mm, width=72*mm, height=60*mm,
                preserveAspectRatio=True, mask='auto')

    # ── Divider ───────────────────────────────────────────────────────────────
    div_y = chart_y-5*mm
    c.setStrokeColor(RULE_GREY); c.setLineWidth(0.5)
    c.line(12*mm, div_y, W-12*mm, div_y)

    # ── Tables ────────────────────────────────────────────────────────────────
    table_top = div_y-3*mm
    col1_x, col2_x = 12*mm, W/2+2*mm
    col_w, row_h   = W/2-16*mm, 6.5*mm

    def section_hdr(x, y, title):
        c.setFillColor(ALPINE_BLUE)
        c.rect(x, y, col_w, 5.5*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7.5)
        c.setFillColor(WHITE)
        c.drawString(x+3*mm, y+1.5*mm, title)
        return y-row_h

    def table_row(x, y, rank, label, value, shade):
        if shade:
            c.setFillColor(LIGHT_BG)
            c.rect(x, y, col_w, row_h-0.5*mm, fill=1, stroke=0)
        c.setFont("Helvetica-Bold", 7)
        c.setFillColor(ALPINE_BLUE)
        c.drawString(x+1.5*mm, y+1.8*mm, str(rank))
        c.setFont("Helvetica", 7)
        c.setFillColor(DARK_GREY)
        short = label[:40]+'...' if len(label)>40 else label
        c.drawString(x+6.5*mm, y+1.8*mm, short)
        c.setFont("Helvetica-Bold", 7)
        c.setFillColor(colors.HexColor("#005A96"))
        c.drawRightString(x+col_w-1*mm, y+1.8*mm, f"{value:,}")
        return y-row_h

    y1 = section_hdr(col1_x, table_top, "TOP REPORTS BY OPENS")
    for i, (title, cnt) in enumerate(data['top_reports'].items()):
        y1 = table_row(col1_x, y1, i+1, title, cnt, i%2==1)
    y1 -= 2.5*mm
    y1 = section_hdr(col1_x, y1, "TOP AUTHORS BY OPENS")
    for i, (author, cnt) in enumerate(data['top_authors'].items()):
        y1 = table_row(col1_x, y1, i+1, author, cnt, i%2==1)

    y2 = section_hdr(col2_x, table_top, "MOST ACTIVE READERS")
    for i, (name, cnt) in enumerate(data['top_readers'].items()):
        y2 = table_row(col2_x, y2, i+1, name, cnt, i%2==1)

    # ── Footer ────────────────────────────────────────────────────────────────
    c.setFillColor(DARK_GREY)
    c.rect(0, 0, W, 9*mm, fill=1, stroke=0)
    c.setFillColor(ALPINE_BLUE)
    c.rect(0, 9*mm, W, 0.8*mm, fill=1, stroke=0)
    c.setFont("Helvetica", 6.5)
    c.setFillColor(colors.HexColor("#AAAAAA"))
    c.drawString(12*mm, 3*mm, f"{account_name}  |  Readership Analytics  |  Confidential")
    c.drawRightString(W-12*mm, 3*mm, "Alpine Macro — An Oxford Economics Company")

    c.save()
    print(f"Report saved: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)

    filepath     = sys.argv[1]
    default_name = os.path.splitext(os.path.basename(filepath))[0].split('_')[-1]
    account_name = sys.argv[2] if len(sys.argv) > 2 else default_name
    default_out  = f"{account_name.replace(' ','_')}_Readership_Report.pdf"
    output_path  = sys.argv[3] if len(sys.argv) > 3 else default_out

    print(f"Loading: {filepath}")
    data = load_and_analyse(filepath)
    print(f"  {data['total_opens']:,} opens  |  {data['unique_readers']} readers  |  {data['date_min']} – {data['date_max']}")
    build_pdf(data, account_name, output_path)
