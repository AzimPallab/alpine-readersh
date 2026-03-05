"""
Alpine Macro — Readership Analytics Web App v2
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
C_TEAL  = "#00A896"
C_DARK  = "#2B2B2B"
C_LBLUE = "#E6F2FB"
C_LTEAL = "#E6F7F5"
C_LGREY = "#F4F8FC"
C_RULE  = "#D0DCE8"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;600&family=DM+Sans:wght@300;400;500;600&display=swap');
html, body, [class*="css"] {{ font-family: 'DM Sans', sans-serif; background: {C_LGREY}; }}
.am-nav {{ background:#fff; border-bottom:3px solid {C_BLUE}; padding:12px 24px; display:flex; align-items:center; justify-content:space-between; margin-bottom:24px; }}
.am-nav-right {{ font-size:0.72rem; color:#888; text-transform:uppercase; letter-spacing:0.1em; font-weight:500; }}
.kpi-card {{ background:#fff; border-radius:10px; padding:14px 16px 12px; border-left:3px solid var(--ac); box-shadow:0 1px 6px rgba(0,0,0,0.05); margin-bottom:4px; }}
.kpi-val {{ font-family:'Playfair Display',serif; font-size:1.85rem; font-weight:600; color:{C_DARK}; line-height:1; margin-bottom:3px; }}
.kpi-lbl {{ font-size:0.65rem; color:#888; text-transform:uppercase; letter-spacing:0.07em; font-weight:500; }}
.sec-hdr {{ display:flex; align-items:center; gap:10px; padding:10px 16px; border-radius:8px; margin:20px 0 14px; font-size:0.8rem; font-weight:600; letter-spacing:0.06em; text-transform:uppercase; }}
.sec-email {{ background:{C_LBLUE}; color:{C_BLUE}; border-left:4px solid {C_BLUE}; }}
.sec-portal {{ background:{C_LTEAL}; color:{C_TEAL}; border-left:4px solid {C_TEAL}; }}
.sec-readers {{ background:#F0EEF8; color:{C_DARK}; border-left:4px solid {C_DARK}; }}
.tbl-hdr {{ font-size:0.68rem; font-weight:600; color:#aaa; text-transform:uppercase; letter-spacing:0.07em; padding-bottom:5px; border-bottom:1px solid {C_RULE}; margin-bottom:3px; display:flex; gap:6px; }}
.tbl-row {{ display:flex; justify-content:space-between; align-items:center; padding:5px 0; font-size:0.78rem; border-bottom:1px solid #F0F4F8; }}
.tbl-rank {{ font-weight:700; color:var(--ac); min-width:16px; font-size:0.7rem; }}
.tbl-lbl {{ flex:1; color:{C_DARK}; padding:0 6px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:210px; }}
.tbl-val {{ font-weight:600; color:#666; font-size:0.74rem; white-space:nowrap; }}
.quality-box {{ background:#FFFBF0; border:1px solid #F0D080; border-radius:8px; padding:12px 16px; font-size:0.78rem; color:#7A5C00; margin-bottom:16px; }}
div[data-testid="stFileUploader"] {{ border:2px dashed {C_BLUE} !important; border-radius:10px !important; background:{C_LBLUE} !important; }}
.dl-btn > button {{ background:{C_DARK} !important; color:#fff !important; border:none !important; border-radius:8px !important; font-weight:600 !important; font-size:0.9rem !important; padding:0.55rem 2rem !important; width:100%; }}
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
    rmap = {c: ALIASES[c.lower().strip()] for c in df.columns if c.lower().strip() in ALIASES and ALIASES[c.lower().strip()] not in df.columns}
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
    def stream(sdf):
        op = sdf[sdf['EventAction']=='initial_open']
        cl = sdf[sdf['EventAction']=='click']
        readers = pd.concat([op,cl]).groupby('Contact Name').size().sort_values(ascending=False)
        monthly = op.groupby('Month').size().reset_index(name='Opens')
        monthly['Label'] = monthly['Month'].astype(str).str[-2:].map(
            {'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'May','06':'Jun',
             '07':'Jul','08':'Aug','09':'Sep','10':'Oct','11':'Nov','12':'Dec'})
        auths = op.groupby('Authors').size().reset_index(name='Opens')
        auths_cl = cl.groupby('Authors').size().reindex(auths['Authors']).fillna(0).astype(int).values
        auths['Engagement Rate'] = (auths_cl / auths['Opens'].replace(0,1)*100).round(1)
        auths = auths.sort_values('Opens', ascending=False).head(5)
        return {
            'opens': len(op), 'clicks': len(cl),
            'unique_readers': op['Contact Name'].nunique(),
            'ctor': round(len(cl)/len(op)*100,1) if len(op) else 0,
            'top_readers': readers.head(10).reset_index().rename(columns={'Contact Name':'Reader', 0:'Activity'}),
            'top_reports': op['Report Title'].str.strip().value_counts().head(5),
            'top_authors': auths,
            'monthly': monthly,
            'top_products': op['Leaf product'].str.strip().value_counts().head(6),
        }

    email_df  = df[df['EventSource']=='email']
    portal_df = df[df['EventSource'].isin(['my oxford','myoxford','portal'])]
    ao = df[df['EventAction']=='initial_open']
    ac = df[df['EventAction']=='click']

    # Reader intelligence
    activity = pd.concat([ao,ac]).groupby('Contact Name').size().sort_values(ascending=False)
    power = activity[activity >= activity.quantile(0.80)].reset_index()
    power.columns = ['Reader','Activity']

    months = ao['Month'].unique()
    loyal = pd.DataFrame(columns=['Reader','Months Active'])
    if len(months)>1:
        loy = ao.groupby('Contact Name')['Month'].nunique()
        loyal = loy[loy==len(months)].reset_index(); loyal.columns=['Reader','Months Active']

    mx = df['EventDate'].max(); cutoff = mx - timedelta(days=30)
    recent = ao[ao['EventDate']>=cutoff]['Contact Name'].unique()
    before = ao[ao['EventDate']<cutoff]['Contact Name'].unique()
    at_risk = pd.DataFrame({'Reader':[r for r in before if r not in recent]})

    rep_op = ao['Report Title'].value_counts()
    rep_cl = ac['Report Title'].value_counts()
    popularity = (rep_op.fillna(0)+rep_cl.reindex(rep_op.index).fillna(0)*2).sort_values(ascending=False).head(10)

    auth_op = ao.groupby('Authors').size()
    auth_cl = ac.groupby('Authors').size()
    aeng = pd.DataFrame({'Author':auth_op.index,'Opens':auth_op.values,
                          'Clicks':auth_cl.reindex(auth_op.index).fillna(0).astype(int).values})
    aeng['Engagement Rate %'] = (aeng['Clicks']/aeng['Opens'].replace(0,1)*100).round(1)
    aeng = aeng.sort_values('Engagement Rate %',ascending=False).head(8)

    return {
        'global':{'total_opens':len(ao),'total_readers':ao['Contact Name'].nunique(),
                  'avg_reports':round(len(ao)/ao['Contact Name'].nunique(),1) if ao['Contact Name'].nunique() else 0,
                  'global_ctor':round(len(ac)/len(ao)*100,1) if len(ao) else 0},
        'email': stream(email_df), 'portal': stream(portal_df),
        'intel':{'power':power,'loyal':loyal,'at_risk':at_risk,'popularity':popularity,'aeng':aeng},
        'date_min':df['EventDate'].min().strftime('%d %b %Y'),
        'date_max':df['EventDate'].max().strftime('%d %b %Y'),
    }

# ── Charts ─────────────────────────────────────────────────────────────────────
def slim_bar(labels, values, color, title=None):
    fig, ax = plt.subplots(figsize=(6, 2.4))
    bars = ax.bar(labels, values, color=color, width=0.32, zorder=3)
    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    ax.yaxis.grid(True, color='#E8EEF4', linewidth=0.4, zorder=0); ax.set_axisbelow(True)
    for s in ['top','right','left']: ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color('#E0E8F0'); ax.spines['bottom'].set_linewidth(0.5)
    ax.tick_params(axis='y', labelsize=7, colors='#999')
    ax.tick_params(axis='x', labelsize=8.5, colors='#444')
    mx = max(values) if values else 1
    for bar, val in zip(bars, values):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+mx*0.02,
                f'{val:,}', ha='center', va='bottom', fontsize=7, color='#444', fontweight='600')
    if title: ax.set_title(title, fontsize=8.5, color='#333', pad=8, fontweight='600', loc='left')
    fig.tight_layout(pad=0.5); return fig

def slim_donut(labels, values, cols, title=None):
    fig, ax = plt.subplots(figsize=(5, 3.8))
    ax.pie(values, labels=None, colors=cols, startangle=90, wedgeprops=dict(width=0.42, edgecolor='white', linewidth=1.5))
    ax.set_facecolor('none'); fig.patch.set_alpha(0)
    short = [l[:22]+'...' if len(l)>22 else l for l in labels]
    patches = [mpatches.Patch(color=cols[i], label=f'{short[i]}  ({values[i]:,})') for i in range(len(labels))]
    ax.legend(handles=patches, loc='lower center', bbox_to_anchor=(0.5,-0.22), fontsize=7, frameon=False, ncol=1, labelspacing=0.4)
    if title: ax.set_title(title, fontsize=8.5, color='#333', pad=8, fontweight='600', loc='left')
    fig.tight_layout(pad=0.5); return fig

def horiz_bar(labels, values, color, title=None):
    fig, ax = plt.subplots(figsize=(6, 3.2))
    y_pos = range(len(labels))
    bars = ax.barh(list(y_pos), values, color=color, height=0.38, zorder=3)
    ax.set_yticks(list(y_pos))
    ax.set_yticklabels([l[:35]+'...' if len(l)>35 else l for l in labels], fontsize=7.5, color='#333')
    ax.invert_yaxis(); ax.set_facecolor('none'); fig.patch.set_alpha(0)
    ax.xaxis.grid(True, color='#E8EEF4', linewidth=0.4, zorder=0); ax.set_axisbelow(True)
    for s in ['top','right','bottom']: ax.spines[s].set_visible(False)
    ax.spines['left'].set_color('#E0E8F0'); ax.tick_params(axis='x', labelsize=7, colors='#999')
    mx = max(values) if values else 1
    for bar, val in zip(bars, values):
        ax.text(val+mx*0.01, bar.get_y()+bar.get_height()/2, f'{val:,}', va='center', fontsize=7, color='#444', fontweight='600')
    if title: ax.set_title(title, fontsize=8.5, color='#333', pad=8, fontweight='600', loc='left')
    fig.tight_layout(pad=0.5); return fig

def render_table(rows, accent, cols=None):
    if not cols: cols=['#','Name','Count']
    hdr = f'<div class="tbl-hdr">{"".join(f"<span>{c}</span>" for c in cols)}</div>'
    body = ""
    for i, row in enumerate(rows):
        v = list(row)
        body += f'<div class="tbl-row"><span class="tbl-rank" style="--ac:{accent}">{i+1}</span><span class="tbl-lbl">{v[0]}</span>'
        for x in v[1:]: body += f'<span class="tbl-val">{x}</span>'
        body += '</div>'
    return hdr + body

def kpi_html(label, value, ac):
    return f'<div class="kpi-card" style="--ac:{ac}"><div class="kpi-val">{value}</div><div class="kpi-lbl">{label}</div></div>'

# ── PDF ────────────────────────────────────────────────────────────────────────
PBLUE=colors.HexColor("#0077C8"); PTEAL=colors.HexColor("#00A896"); PDARK=colors.HexColor("#2B2B2B")
PLBLU=colors.HexColor("#E6F2FB"); PGREY=colors.HexColor("#F4F8FC"); PRULE=colors.HexColor("#D0DCE8")
PWHITE=colors.white; PMGREY=colors.HexColor("#888888")

def fig_to_ir(fig):
    buf=io.BytesIO(); fig.savefig(buf,format='png',dpi=150,bbox_inches='tight',transparent=True)
    buf.seek(0); plt.close(fig); return ImageReader(buf)

def pdf_nav(c,W,H,account,data):
    c.setFillColor(PWHITE); c.rect(0,H-36*mm,W,36*mm,fill=1,stroke=0)
    c.setFillColor(PBLUE);  c.rect(0,H-1.5*mm,W,1.5*mm,fill=1,stroke=0)
    if os.path.exists(LOGO_PATH):
        c.drawImage(LOGO_PATH,12*mm,H-32*mm,width=52*mm,height=22*mm,preserveAspectRatio=True,mask='auto')
    c.setFont("Helvetica-Bold",13); c.setFillColor(PDARK)
    c.drawRightString(W-12*mm,H-16*mm,"Readership Analytics Report")
    c.setFont("Helvetica",8); c.setFillColor(PBLUE)
    c.drawRightString(W-12*mm,H-22*mm,f"{account}  |  {data['date_min']} – {data['date_max']}")
    c.setFont("Helvetica",7); c.setFillColor(PMGREY)
    c.drawRightString(W-12*mm,H-27*mm,f"Generated {datetime.now().strftime('%d %b %Y')}")
    c.setFillColor(PRULE); c.rect(0,H-37*mm,W,0.5*mm,fill=1,stroke=0)

def pdf_foot(c,W,account,page):
    c.setFillColor(PDARK); c.rect(0,0,W,8*mm,fill=1,stroke=0)
    c.setFillColor(PBLUE); c.rect(0,8*mm,W,0.6*mm,fill=1,stroke=0)
    c.setFont("Helvetica",6); c.setFillColor(colors.HexColor("#AAAAAA"))
    c.drawString(12*mm,2.8*mm,f"{account}  |  Readership Analytics  |  Confidential")
    c.drawRightString(W-12*mm,2.8*mm,f"Alpine Macro — An Oxford Economics Company  |  Page {page}")

def pdf_shdr(c,x,y,w,title,col):
    c.setFillColor(col); c.roundRect(x,y,w,6*mm,2,fill=1,stroke=0)
    c.setFont("Helvetica-Bold",7.5); c.setFillColor(PWHITE)
    c.drawString(x+4*mm,y+1.8*mm,title)

def pdf_kpis(c,kpis,x,y,w):
    bw=(w-(len(kpis)-1)*2*mm)/len(kpis)
    for i,(lbl,val,col) in enumerate(kpis):
        bx=x+i*(bw+2*mm)
        c.setFillColor(col); c.roundRect(bx,y,bw,14*mm,2,fill=1,stroke=0)
        c.setFont("Helvetica-Bold",12); c.setFillColor(PWHITE)
        c.drawCentredString(bx+bw/2,y+6.5*mm,str(val))
        c.setFont("Helvetica",6); c.setFillColor(colors.HexColor("#DDDDDD"))
        c.drawCentredString(bx+bw/2,y+2.5*mm,lbl.upper())

def pdf_tbl(c,rows,x,y,cw,accent,rh=5.8*mm):
    for i,(lbl,val) in enumerate(rows):
        if i%2==1: c.setFillColor(PGREY); c.rect(x,y,cw,rh-0.3*mm,fill=1,stroke=0)
        c.setFont("Helvetica-Bold",6.5); c.setFillColor(accent); c.drawString(x+1.5*mm,y+1.5*mm,str(i+1))
        c.setFont("Helvetica",6.5); c.setFillColor(PDARK)
        c.drawString(x+6*mm,y+1.5*mm,lbl[:42]+'...' if len(lbl)>42 else lbl)
        c.setFont("Helvetica-Bold",6.5); c.setFillColor(accent)
        c.drawRightString(x+cw-1.5*mm,y+1.5*mm,str(val))
        y-=rh
    return y

def generate_pdf(data, account):
    buf=io.BytesIO(); W,H=A4; c=canvas.Canvas(buf,pagesize=A4)
    M=12*mm; CW=(W-2*M-4*mm)/2

    # Page 1
    pdf_nav(c,W,H,account,data)
    g=data['global']
    y=H-55*mm
    pdf_kpis(c,[("Total Opens",f"{g['total_opens']:,}",PBLUE),
                ("Unique Readers",f"{g['total_readers']:,}",PDARK),
                ("Avg Reports/Reader",str(g['avg_reports']),colors.HexColor("#005A96")),
                ("Click-to-Open",f"{g['global_ctor']}%",PTEAL)],M,y,W-2*M)
    y-=20*mm
    pdf_shdr(c,M,y,W-2*M,"EMAIL ENGAGEMENT",PBLUE); y-=8*mm
    e=data['email']
    pdf_kpis(c,[("Email Opens",f"{e['opens']:,}",PBLUE),
                ("Clicks",f"{e['clicks']:,}",colors.HexColor("#005A96")),
                ("Unique Readers",f"{e['unique_readers']:,}",colors.HexColor("#33A1E0")),
                ("CTOR",f"{e['ctor']}%",colors.HexColor("#80C4ED"))],M,y,W-2*M)
    y-=18*mm
    if len(e['monthly'])>0:
        c.drawImage(fig_to_ir(slim_bar(e['monthly']['Label'].tolist(),e['monthly']['Opens'].tolist(),"#0077C8","Monthly Opens")),
                    M,y-42*mm,width=CW,height=42*mm,preserveAspectRatio=True,mask='auto')
    if len(e['top_products'])>0:
        c.drawImage(fig_to_ir(slim_donut(e['top_products'].index.tolist(),e['top_products'].values.tolist(),
                    ["#0077C8","#005A96","#33A1E0","#80C4ED","#B3DDF5","#E6F4FC"])),
                    M+CW+4*mm,y-44*mm,width=CW,height=44*mm,preserveAspectRatio=True,mask='auto')
    y-=46*mm
    c.setFont("Helvetica-Bold",7); c.setFillColor(PBLUE)
    c.drawString(M,y,"TOP REPORTS"); c.drawString(M+CW+4*mm,y,"TOP AUTHORS"); y-=4*mm
    pdf_tbl(c,[(t[:42],f"{v:,}") for t,v in e['top_reports'].head(5).items()],M,y,CW,PBLUE)
    pdf_tbl(c,[(row['Authors'],f"{row['Opens']:,}") for _,row in e['top_authors'].head(5).iterrows()],M+CW+4*mm,y,CW,PBLUE)
    pdf_foot(c,W,account,1); c.showPage()

    # Page 2
    pdf_nav(c,W,H,account,data); y=H-46*mm
    pdf_shdr(c,M,y,W-2*M,"WEB PORTAL ACTIVITY",PTEAL); y-=8*mm
    p=data['portal']
    pdf_kpis(c,[("Portal Opens",f"{p['opens']:,}",PTEAL),
                ("Portal Clicks",f"{p['clicks']:,}",colors.HexColor("#007A6E")),
                ("Unique Readers",f"{p['unique_readers']:,}",colors.HexColor("#00C4B0"))],M,y,W-2*M)
    y-=18*mm
    c.setFont("Helvetica-Bold",7); c.setFillColor(PTEAL)
    c.drawString(M,y,"MOST ACTIVE PORTAL USERS"); c.drawString(M+CW+4*mm,y,"TOP PORTAL REPORTS"); y-=4*mm
    pdf_tbl(c,[(row['Reader'],f"{row.iloc[1]:,}") for _,row in p['top_readers'].head(8).iterrows()],M,y,CW,PTEAL)
    pdf_tbl(c,[(t[:42],f"{v:,}") for t,v in p['top_reports'].head(8).items()],M+CW+4*mm,y,CW,PTEAL)
    y-=52*mm
    pdf_shdr(c,M,y,W-2*M,"READER INTELLIGENCE",PDARK); y-=8*mm
    intel=data['intel']
    c.setFont("Helvetica-Bold",7); c.setFillColor(PDARK)
    c.drawString(M,y,"POWER READERS (TOP 20%)"); c.drawString(M+CW+4*mm,y,"AUTHOR ENGAGEMENT RATE"); y-=4*mm
    pdf_tbl(c,[(row['Reader'],f"{row['Activity']:,}") for _,row in intel['power'].head(8).iterrows()],M,y,CW,PDARK)
    pdf_tbl(c,[(row['Author'],f"{row['Engagement Rate %']}%") for _,row in intel['aeng'].head(8).iterrows()],M+CW+4*mm,y,CW,PDARK)
    y-=52*mm
    if len(intel['popularity'])>0:
        c.drawImage(fig_to_ir(horiz_bar(intel['popularity'].index.tolist(),intel['popularity'].values.tolist(),"#2B2B2B","Report Popularity Score")),
                    M,y-46*mm,width=W-2*M,height=46*mm,preserveAspectRatio=True,mask='auto')
    pdf_foot(c,W,account,2); c.save(); buf.seek(0)
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

    # Global KPIs
    g = data['global']
    c1,c2,c3,c4 = st.columns(4)
    for col,lbl,val,ac in [(c1,"Total Unique Opens",f"{g['total_opens']:,}",C_BLUE),
                            (c2,"Unique Readers",f"{g['total_readers']:,}",C_DARK),
                            (c3,"Avg Reports / Reader",str(g['avg_reports']),"#005A96"),
                            (c4,"Click-to-Open Rate",f"{g['global_ctor']}%",C_TEAL)]:
        with col: st.markdown(kpi_html(lbl,val,ac), unsafe_allow_html=True)

    st.markdown("")
    tab1,tab2,tab3 = st.tabs(["📧  Email","🌐  Web Portal","👥  Reader Intelligence"])

    with tab1:
        st.markdown('<div class="sec-hdr sec-email">📧 Email Engagement</div>', unsafe_allow_html=True)
        e=data['email']
        c1,c2,c3,c4=st.columns(4)
        for col,lbl,val,ac in [(c1,"Email Opens",f"{e['opens']:,}",C_BLUE),(c2,"Email Clicks",f"{e['clicks']:,}","#005A96"),
                                (c3,"Unique Readers",f"{e['unique_readers']:,}","#33A1E0"),(c4,"CTOR",f"{e['ctor']}%","#80C4ED")]:
            with col: st.markdown(kpi_html(lbl,val,ac), unsafe_allow_html=True)
        st.markdown("")
        ch1,ch2=st.columns([3,2])
        with ch1:
            if len(e['monthly'])>0:
                st.pyplot(slim_bar(e['monthly']['Label'].tolist(),e['monthly']['Opens'].tolist(),C_BLUE,"Monthly Email Opens"),use_container_width=True)
        with ch2:
            if len(e['top_products'])>0:
                st.pyplot(slim_donut(e['top_products'].index.tolist(),e['top_products'].values.tolist(),
                          [C_BLUE,"#005A96","#33A1E0","#80C4ED","#B3DDF5","#E6F4FC"],"Opens by Category"),use_container_width=True)
        t1,t2,t3=st.columns(3)
        with t1:
            st.markdown("**Top Reports**")
            st.markdown(render_table([(t,f"{v:,}") for t,v in e['top_reports'].items()],C_BLUE,['#','Report','Opens']),unsafe_allow_html=True)
        with t2:
            st.markdown("**Top Authors**")
            st.markdown(render_table([(row['Authors'],f"{row['Opens']:,}") for _,row in e['top_authors'].iterrows()],C_BLUE,['#','Author','Opens']),unsafe_allow_html=True)
        with t3:
            st.markdown("**Most Active Readers**")
            st.markdown(render_table([(row['Reader'],f"{row.iloc[1]:,}") for _,row in e['top_readers'].head(10).iterrows()],C_BLUE,['#','Reader','Activity']),unsafe_allow_html=True)

    with tab2:
        st.markdown('<div class="sec-hdr sec-portal">🌐 Web Portal Activity</div>', unsafe_allow_html=True)
        p=data['portal']
        if p['opens']==0 and p['clicks']==0:
            st.info("No web portal activity found in this dataset.")
        else:
            c1,c2,c3=st.columns(3)
            for col,lbl,val,ac in [(c1,"Portal Opens",f"{p['opens']:,}",C_TEAL),(c2,"Portal Clicks",f"{p['clicks']:,}","#007A6E"),
                                    (c3,"Unique Readers",f"{p['unique_readers']:,}","#00C4B0")]:
                with col: st.markdown(kpi_html(lbl,val,ac), unsafe_allow_html=True)
            st.markdown("")
            t1,t2=st.columns(2)
            with t1:
                st.markdown("**Most Active Portal Users**")
                st.markdown(render_table([(row['Reader'],f"{row.iloc[1]:,}") for _,row in p['top_readers'].head(10).iterrows()],C_TEAL,['#','Reader','Activity']),unsafe_allow_html=True)
            with t2:
                st.markdown("**Top Reports via Portal**")
                st.markdown(render_table([(t,f"{v:,}") for t,v in p['top_reports'].head(10).items()],C_TEAL,['#','Report','Opens']),unsafe_allow_html=True)

    with tab3:
        st.markdown('<div class="sec-hdr sec-readers">👥 Reader Intelligence</div>', unsafe_allow_html=True)
        intel=data['intel']
        c1,c2,c3=st.columns(3)
        with c1: st.markdown(kpi_html("Power Readers",len(intel['power']),C_DARK), unsafe_allow_html=True)
        with c2: st.markdown(kpi_html("Loyal Readers",len(intel['loyal']),"#444"), unsafe_allow_html=True)
        with c3: st.markdown(kpi_html("At-Risk Readers",len(intel['at_risk']),"#C0392B"), unsafe_allow_html=True)
        st.markdown("")
        ch1,ch2=st.columns([3,2])
        with ch1:
            if len(intel['popularity'])>0:
                st.pyplot(horiz_bar(intel['popularity'].index.tolist(),intel['popularity'].values.tolist(),C_DARK,"Report Popularity Score (Opens + Clicks×2)"),use_container_width=True)
        with ch2:
            st.markdown("**Author Engagement Rate**")
            st.markdown(render_table([(row['Author'],f"{row['Engagement Rate %']}%") for _,row in intel['aeng'].iterrows()],C_DARK,['#','Author','Rate']),unsafe_allow_html=True)
        t1,t2,t3=st.columns(3)
        with t1:
            st.markdown("**Power Readers**")
            st.markdown(render_table([(row['Reader'],f"{row['Activity']:,}") for _,row in intel['power'].head(10).iterrows()],C_DARK,['#','Reader','Activity']),unsafe_allow_html=True)
        with t2:
            st.markdown("**Loyal Readers** *(every month)*")
            if len(intel['loyal'])>0:
                st.markdown(render_table([(row['Reader'],str(row['Months Active'])) for _,row in intel['loyal'].head(10).iterrows()],"#444",['#','Reader','Months']),unsafe_allow_html=True)
            else:
                st.caption("Not enough months of data.")
        with t3:
            st.markdown("**At-Risk Readers** *(silent 30+ days)*")
            if len(intel['at_risk'])>0:
                st.markdown(render_table([(row['Reader'],"⚠️") for _,row in intel['at_risk'].head(10).iterrows()],"#C0392B",['#','Reader','']),unsafe_allow_html=True)
            else:
                st.success("No at-risk readers!")

    st.markdown("---")
    dl,_=st.columns([1,3])
    with dl:
        with st.spinner("Building PDF..."): pdf_bytes=generate_pdf(data,account_name)
        st.markdown('<div class="dl-btn">',unsafe_allow_html=True)
        st.download_button("⬇️  Download PDF Report",data=pdf_bytes,
                           file_name=f"{account_name.replace(' ','_')}_Readership_Report.pdf",mime="application/pdf")
        st.markdown('</div>',unsafe_allow_html=True)

else:
    st.info("Upload an analytics export above to get started.")
    st.markdown("""**Accepted:** <span class="chip">.xlsx</span> <span class="chip">.xls</span> <span class="chip">.csv</span><br><br>
    **Required columns:** <span class="chip">Contact Name</span> <span class="chip">Contact Email</span>
    <span class="chip">EventSource</span> <span class="chip">EventAction</span> <span class="chip">EventDate</span>
    <span class="chip">Report Title</span> <span class="chip">Authors</span> <span class="chip">Leaf product</span>
    """, unsafe_allow_html=True)
