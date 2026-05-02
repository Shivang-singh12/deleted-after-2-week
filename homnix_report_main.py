"""
╔══════════════════════════════════════════════════════════════════╗
║        HOMNIX — AMAZON MONTHLY REPORT GENERATOR v2              ║
║                                                                  ║
║  Usage:                                                          ║
║    python homnix_report_generator.py report.xlsx cogs.csv        ║
║    python homnix_report_generator.py report.xlsx cogs.csv ./out  ║
║    python homnix_report_generator.py report.xlsx                 ║
║                                                                  ║
║  Output: Homnix_Report_<Mon-YYYY>.pdf                            ║
╚══════════════════════════════════════════════════════════════════╝

Requirements:
    pip install reportlab pandas openpyxl

Files needed each month:
    1. report.xlsx  — Amazon Business Report (By ASIN)
                      Seller Central → Reports → Business Reports → By ASIN
    2. cogs.csv     — Your cost-of-goods file (update costs anytime)
                      Columns: sku, asin, product_name, cogs_per_unit, notes
                      Run once with --create-cogs to generate a starter template.

Examples:
    python homnix_report_main.py report_2026-03-15.xlsx homnix_cogs.csv  <--- Always use this one, paste this line in terminal and press enter
    python homnix_report_main.py report_2026-03-15.xlsx            # skips COGS page
    python homnix_report_main.py --create-cogs                  # make starter CSV
"""

import sys
import os
import csv
import pandas as pd
from datetime import datetime

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, HRFlowable, PageBreak)
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
except ImportError:
    print("ERROR: reportlab not installed.\nRun:  pip install reportlab pandas openpyxl")
    sys.exit(1)


# ══════════════════════════════════════════════════════════════════════════════
# COLORS
# ══════════════════════════════════════════════════════════════════════════════
NAVY        = colors.HexColor('#1B2A4A')
BLUE        = colors.HexColor('#2E5BBA')
LIGHT_BLUE  = colors.HexColor('#EBF2FF')
TEAL        = colors.HexColor('#0D7377')
PALE_TEAL   = colors.HexColor('#E8F8F5')
GREEN       = colors.HexColor('#1A7A4A')
PALE_GREEN  = colors.HexColor('#E8F5EE')
RED         = colors.HexColor('#C0392B')
PALE_RED    = colors.HexColor('#FDEDEC')
ORANGE      = colors.HexColor('#D35400')
PALE_ORANGE = colors.HexColor('#FEF0E6')
GOLD        = colors.HexColor('#B7860B')
PALE_GOLD   = colors.HexColor('#FEFAEC')
PURPLE      = colors.HexColor('#6C3483')
PALE_PURPLE = colors.HexColor('#F4ECF7')
GRAY        = colors.HexColor('#5D6D7E')
LIGHT_GRAY  = colors.HexColor('#F4F6F7')
MID_GRAY    = colors.HexColor('#BDC3C7')
DARK        = colors.HexColor('#2C3E50')
WHITE       = colors.white


# ══════════════════════════════════════════════════════════════════════════════
# AMAZON REPORT COLUMN MAP  (0-indexed positions in raw export)
# ══════════════════════════════════════════════════════════════════════════════
COL = {
    'asin'         : 0,
    'sku'          : 1,
    'product_name' : 2,
    'net_proceed'  : 5,
    'total_sales'  : 13,
    'avg_price'    : 14,
    'units_sold'   : 15,
    'refunded'     : 16,
    'net_units'    : 17,
    'return_rate'  : 18,
    'fulfillment'  : 20,
    'selling_fees' : 138,
    'referral_fee' : 144,
    'storage'      : 185,
    'advertising'  : 309,
    'returns_ops'  : 254,
}


# ══════════════════════════════════════════════════════════════════════════════
# STYLE HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def S(name, **kw):
    return ParagraphStyle(name, **kw)

STYLES = {
    'title'  : S('title',  fontSize=20, textColor=WHITE,  fontName='Helvetica-Bold',  alignment=TA_CENTER, leading=26),
    'sub'    : S('sub',    fontSize=10, textColor=colors.HexColor('#C8D8FF'), fontName='Helvetica', alignment=TA_CENTER),
    'sec_hdr': S('sec_hdr',fontSize=12, textColor=WHITE,  fontName='Helvetica-Bold',  leading=16),
    'body'   : S('body',   fontSize=9,  textColor=DARK,   fontName='Helvetica',        leading=13, spaceAfter=2),
    'sm'     : S('sm',     fontSize=8,  textColor=DARK,   fontName='Helvetica',        leading=11),
    'sm_r'   : S('sm_r',   fontSize=8,  textColor=DARK,   fontName='Helvetica',        alignment=TA_RIGHT),
    'sm_c'   : S('sm_c',   fontSize=8,  textColor=DARK,   fontName='Helvetica',        alignment=TA_CENTER),
    'smb'    : S('smb',    fontSize=8,  textColor=DARK,   fontName='Helvetica-Bold'),
    'smb_r'  : S('smb_r',  fontSize=8,  textColor=DARK,   fontName='Helvetica-Bold',   alignment=TA_RIGHT),
    'smb_c'  : S('smb_c',  fontSize=8,  textColor=DARK,   fontName='Helvetica-Bold',   alignment=TA_CENTER),
    'note'   : S('note',   fontSize=7.5,textColor=GRAY,   fontName='Helvetica-Oblique',alignment=TA_CENTER),
    'red_r'  : S('red_r',  fontSize=8,  textColor=RED,    fontName='Helvetica-Bold',   alignment=TA_RIGHT),
    'grn_r'  : S('grn_r',  fontSize=8,  textColor=GREEN,  fontName='Helvetica-Bold',   alignment=TA_RIGHT),
    'og_r'   : S('og_r',   fontSize=8,  textColor=ORANGE, fontName='Helvetica-Bold',   alignment=TA_RIGHT),
    'pur_r'  : S('pur_r',  fontSize=8,  textColor=PURPLE, fontName='Helvetica-Bold',   alignment=TA_RIGHT),
}

_style_counter = [0]

def P(text, style_key='body', **kw):
    if isinstance(style_key, str):
        base = STYLES[style_key]
        if kw:
            _style_counter[0] += 1
            base = ParagraphStyle(f'dyn_{_style_counter[0]}', parent=base, **kw)
    else:
        base = style_key
    return Paragraph(str(text), base)

def fmt(v, prefix='Rs '):
    if v < 0:
        return f'-{prefix}{abs(v):,.2f}'
    return f'{prefix}{v:,.2f}'

def pct(v):
    return f'{v:.1f}%'

def color_pct(val, warn=30, ok=20, align=TA_CENTER):
    c = RED if val > warn else (ORANGE if val > ok else GREEN)
    _style_counter[0] += 1
    st = S(f'cp_{_style_counter[0]}', fontSize=7.5, textColor=c,
           fontName='Helvetica-Bold', alignment=align)
    return Paragraph(pct(val), st)

def color_val(val, good_above=0, warn_below=0, align=TA_RIGHT, size=8):
    c = GREEN if val > good_above else (ORANGE if val > warn_below else RED)
    _style_counter[0] += 1
    st = S(f'cv_{_style_counter[0]}', fontSize=size, textColor=c,
           fontName='Helvetica-Bold', alignment=align)
    return Paragraph(fmt(val), st)

def base_ts():
    return TableStyle([
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [WHITE, LIGHT_GRAY]),
        ('BOX',            (0,0), (-1,-1), 0.5, MID_GRAY),
        ('INNERGRID',      (0,0), (-1,-1), 0.3, MID_GRAY),
        ('TOPPADDING',     (0,0), (-1,-1), 3),
        ('BOTTOMPADDING',  (0,0), (-1,-1), 3),
        ('LEFTPADDING',    (0,0), (-1,-1), 4),
        ('RIGHTPADDING',   (0,0), (-1,-1), 4),
        ('VALIGN',         (0,0), (-1,-1), 'MIDDLE'),
    ])

def banner(W, title, subtitle=''):
    rows = [[P(title, 'title')]]
    if subtitle:
        rows.append([P(subtitle, 'sub')])
    t = Table(rows, colWidths=[W])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), NAVY),
        ('TOPPADDING',    (0,0), (-1,0),  14),
        ('BOTTOMPADDING', (0,-1),(-1,-1), 14),
        ('TOPPADDING',    (0,1), (-1,-1),  3),
        ('LEFTPADDING',   (0,0), (-1,-1), 10),
        ('RIGHTPADDING',  (0,0), (-1,-1), 10),
    ]))
    return t

def sec_hdr(W, text, bg=BLUE):
    t = Table([[P(text, 'sec_hdr')]], colWidths=[W])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), bg),
        ('TOPPADDING',    (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
        ('LEFTPADDING',   (0,0), (-1,-1), 10),
    ]))
    return t

def kpi_box(label, value, sub='', bg=LIGHT_BLUE, val_color=NAVY):
    _style_counter[0] += 1
    n = _style_counter[0]
    vs = S(f'kv_{n}', fontSize=17, textColor=val_color, fontName='Helvetica-Bold',
           alignment=TA_CENTER, leading=21)
    ls = S(f'kl_{n}', fontSize=7.5, textColor=GRAY,      fontName='Helvetica',     alignment=TA_CENTER)
    ss = S(f'ks_{n}', fontSize=7,   textColor=GRAY,      fontName='Helvetica-Oblique', alignment=TA_CENTER)
    rows = [[Paragraph(label, ls)], [Paragraph(str(value), vs)]]
    if sub:
        rows.append([Paragraph(str(sub), ss)])
    t = Table(rows, colWidths=[4*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), bg),
        ('BOX',           (0,0), (-1,-1), 0.5, MID_GRAY),
        ('TOPPADDING',    (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING',    (0,0), (0,0),   8),
    ]))
    return t


# ══════════════════════════════════════════════════════════════════════════════
# COGS CSV  — load & create template
# ══════════════════════════════════════════════════════════════════════════════
COGS_COLUMNS = ['sku', 'asin', 'product_name', 'cogs_per_unit', 'notes']

def create_cogs_template(products, outpath='homnix_cogs.csv'):
    """Write a starter COGS CSV pre-filled with all ASINs/SKUs from the report."""
    with open(outpath, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=COGS_COLUMNS)
        w.writeheader()
        for p in products:
            w.writerow({
                'sku'          : p['sku'],
                'asin'         : p['asin'],
                'product_name' : p['short'],
                'cogs_per_unit': '',          # ← fill this in
                'notes'        : '',
            })
    return outpath

def load_cogs(csv_path):
    """
    Load COGS CSV.  Returns dict keyed by SKU AND by ASIN for flexible lookup.
    Each value is the cost per unit as a float (0.0 if blank/missing).
    """
    cogs = {}   # key → cogs_per_unit
    if not csv_path or not os.path.exists(csv_path):
        return cogs
    try:
        df = pd.read_csv(csv_path, dtype=str).fillna('')
        # normalise column names
        df.columns = [c.strip().lower().replace(' ', '_') for c in df.columns]
        for _, row in df.iterrows():
            cpu_raw = row.get('cogs_per_unit', '').strip()
            try:
                cpu = float(cpu_raw) if cpu_raw else 0.0
            except ValueError:
                cpu = 0.0
            sku  = row.get('sku',  '').strip()
            asin = row.get('asin', '').strip()
            if sku:
                cogs[sku] = cpu
            if asin:
                cogs[asin] = cpu   # ASIN fallback
    except Exception as e:
        print(f"  WARNING: Could not read COGS file — {e}")
    return cogs


# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════
def load_products(filepath, cogs_map=None):
    """
    Parse Amazon Business Report xlsx.
    Returns (products list, totals dict, period string).
    If cogs_map provided, enriches each product with COGS and true-profit fields.
    """
    try:
        ext = os.path.splitext(filepath)[1].lower()
        if ext == '.csv':
            df = pd.read_csv(filepath, header=None)
        else:
            df = pd.read_excel(filepath, sheet_name='Report', header=None)
    except Exception as e:
        print(f"ERROR reading file: {e}")
        sys.exit(1)

    # Infer reporting period from filename
    fname  = os.path.basename(filepath)
    period = datetime.today().strftime('%b %Y')
    import re
    m = re.search(r'(\d{4})-(\d{2})', fname)
    if m:
        try:
            period = datetime(int(m.group(1)), int(m.group(2)), 1).strftime('%b %Y')
        except ValueError:
            pass

    if cogs_map is None:
        cogs_map = {}

    products = []
    for i in range(2, len(df)):
        r = df.iloc[i]
        asin = str(r[COL['asin']]).strip() if pd.notna(r[COL['asin']]) else ''
        if not asin or asin.lower() in ('nan', 'total', ''):
            continue

        def v(col):
            idx = COL.get(col, col)
            val = r[idx] if idx < len(r) else None
            return float(val) if pd.notna(val) else 0.0

        def sv(col):
            idx = COL.get(col, col)
            val = r[idx] if idx < len(r) else None
            return str(val).strip() if pd.notna(val) else ''

        sku           = sv('sku')
        total_sales   = v('total_sales')
        net_proceed   = v('net_proceed')
        fulfillment   = v('fulfillment')
        selling_fees  = v('selling_fees') or v('referral_fee')
        advertising   = v('advertising')
        storage       = v('storage')
        return_refund = v('returns_ops')
        units_sold    = int(v('units_sold'))
        refunded      = int(v('refunded'))
        net_units     = int(v('net_units'))
        return_rate   = v('return_rate')
        avg_price     = v('avg_price')

        total_costs   = fulfillment + selling_fees + advertising + storage + abs(return_refund)
        net_margin    = (net_proceed / total_sales * 100) if total_sales else 0

        name_full = sv('product_name')
        short     = name_full.split('|')[0].strip()
        short     = (short[:57] + '...') if len(short) > 60 else short

        # ── COGS enrichment ──────────────────────────────────────────────────
        # Lookup order: SKU → ASIN → 0
        cogs_per_unit   = cogs_map.get(sku, cogs_map.get(asin, 0.0))
        total_cogs      = cogs_per_unit * net_units          # charged on units actually shipped
        true_profit     = net_proceed - total_cogs
        true_margin     = (true_profit / total_sales * 100) if total_sales else 0
        cogs_pct_rev    = (total_cogs / total_sales * 100)  if total_sales else 0
        has_cogs        = cogs_per_unit > 0

        products.append({
            'asin'         : asin,
            'sku'          : sku,
            'name_full'    : name_full,
            'short'        : short,
            'total_sales'  : total_sales,
            'net_proceed'  : net_proceed,
            'avg_price'    : avg_price,
            'units_sold'   : units_sold,
            'refunded'     : refunded,
            'net_units'    : net_units,
            'return_rate'  : return_rate,
            'fulfillment'  : fulfillment,
            'selling_fees' : selling_fees,
            'advertising'  : advertising,
            'storage'      : storage,
            'return_refund': return_refund,
            'total_costs'  : total_costs,
            'net_margin'   : net_margin,
            # COGS fields
            'cogs_per_unit': cogs_per_unit,
            'total_cogs'   : total_cogs,
            'true_profit'  : true_profit,
            'true_margin'  : true_margin,
            'cogs_pct_rev' : cogs_pct_rev,
            'has_cogs'     : has_cogs,
        })

    if not products:
        print("ERROR: No product rows found. Is the sheet named 'Report'?")
        sys.exit(1)

    num_keys = ['total_sales','net_proceed','units_sold','refunded','net_units',
                'fulfillment','selling_fees','advertising','storage','total_costs',
                'total_cogs','true_profit']
    totals = {k: sum(p[k] for p in products) for k in num_keys}
    totals['avg_price']   = totals['total_sales'] / totals['units_sold'] if totals['units_sold'] else 0
    totals['return_rate'] = totals['refunded']    / totals['units_sold'] if totals['units_sold'] else 0
    totals['net_margin']  = (totals['net_proceed'] / totals['total_sales'] * 100) if totals['total_sales'] else 0
    totals['true_margin'] = (totals['true_profit'] / totals['total_sales'] * 100) if totals['total_sales'] else 0
    totals['cogs_pct_rev']= (totals['total_cogs']  / totals['total_sales'] * 100) if totals['total_sales'] else 0
    totals['cogs_covered'] = sum(1 for p in products if p['has_cogs'])
    totals['cogs_missing'] = len(products) - totals['cogs_covered']

    return products, totals, period


# ══════════════════════════════════════════════════════════════════════════════
# PDF BUILDER
# ══════════════════════════════════════════════════════════════════════════════
def build_report(products, totals, period, output_path, has_cogs_file=False):
    T  = totals
    ps = sorted(products, key=lambda x: x['total_sales'], reverse=True)
    run_date = datetime.today().strftime('%d %B %Y')

    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            rightMargin=1.8*cm, leftMargin=1.8*cm,
                            topMargin=1.8*cm,  bottomMargin=1.8*cm)
    W     = 17.4 * cm
    story = []

    cogs_note = f"  |  COGS: {T['cogs_covered']} of {T['cogs_covered']+T['cogs_missing']} ASINs covered" if has_cogs_file else "  |  COGS: not provided"

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 1 — HEADER + EXECUTIVE SUMMARY
    # ══════════════════════════════════════════════════════════════════════════
    story.append(banner(W,
        'Homnix Amazon — Business Performance Report',
        f'{period} Report  |  Generated: {run_date}  |  {len(products)} ASINs{cogs_note}'
    ))
    story.append(Spacer(1, 0.35*cm))

    # ── KPI Row 1 ─────────────────────────────────────────────────────────────
    kw4 = dict(colWidths=[W/4]*4)
    kpis1 = Table([[
        kpi_box('GROSS REVENUE',     fmt(T['total_sales']),  f"{T['units_sold']} units sold",              LIGHT_BLUE,  NAVY),
        kpi_box('NET PROCEEDS',      fmt(T['net_proceed']),  f"Margin: {pct(T['net_margin'])}",
                PALE_GREEN if T['net_proceed'] >= 0 else PALE_RED,
                GREEN      if T['net_proceed'] >= 0 else RED),
        kpi_box('UNITS SOLD',        str(T['units_sold']),   f"{T['refunded']} refunded / {pct(T['return_rate']*100)} rate", PALE_GOLD,   GOLD),
        kpi_box('AVG SELLING PRICE', fmt(T['avg_price']),    'Across all products',                        PALE_ORANGE, ORANGE),
    ]], **kw4)
    kpis1.setStyle(TableStyle([('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4)]))
    story.append(kpis1)
    story.append(Spacer(1, 0.15*cm))

    # ── KPI Row 2 ─────────────────────────────────────────────────────────────
    kpis2 = Table([[
        kpi_box('FULFILLMENT COST',  fmt(T['fulfillment']),  pct(T['fulfillment'] /T['total_sales']*100)+' of rev', PALE_RED,    RED),
        kpi_box('SELLING FEES',      fmt(T['selling_fees']), pct(T['selling_fees']/T['total_sales']*100)+' of rev', PALE_RED,    RED),
        kpi_box('ADVERTISING SPEND', fmt(T['advertising']),  pct(T['advertising'] /T['total_sales']*100)+' of rev', PALE_ORANGE, ORANGE),
        kpi_box('TRUE PROFIT' if has_cogs_file else 'TOTAL COSTS',
                fmt(T['true_profit']) if has_cogs_file else fmt(T['total_costs']),
                (f"Margin: {pct(T['true_margin'])}" if has_cogs_file else pct(T['total_costs']/T['total_sales']*100)+' of rev'),
                (PALE_GREEN if T['true_profit'] >= 0 else PALE_RED) if has_cogs_file else PALE_RED,
                (GREEN if T['true_profit'] >= 0 else RED) if has_cogs_file else RED),
    ]], **kw4)
    kpis2.setStyle(TableStyle([('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4)]))
    story.append(kpis2)
    story.append(Spacer(1, 0.3*cm))

    # ── P&L Summary ───────────────────────────────────────────────────────────
    story.append(sec_hdr(W, '  P & L  SUMMARY'))

    pnl_rows = [
        [P('<b>Category</b>','smb'), P('<b>Amount (₹)</b>','smb_r'), P('<b>% of Revenue</b>','smb_r')],
        [P('Gross Revenue (Total Sales)','sm'),       P(fmt(T['total_sales']),'sm_r'),  P('100.0%','sm_r')],
        [P('  ↳ Fulfillment Cost (FBA)','sm'),         P(fmt(-T['fulfillment']),'red_r'), P(pct(-T['fulfillment'] /T['total_sales']*100),'red_r')],
        [P('  ↳ Selling Fees (Referral)','sm'),        P(fmt(-T['selling_fees']),'red_r'),P(pct(-T['selling_fees']/T['total_sales']*100),'red_r')],
        [P('  ↳ Advertising Spend','sm'),              P(fmt(-T['advertising']),'red_r'), P(pct(-T['advertising'] /T['total_sales']*100),'red_r')],
        [P('  ↳ Storage Costs','sm'),                  P(fmt(-T['storage']),'red_r'),     P(pct(-T['storage']     /T['total_sales']*100),'red_r')],
        [P('<b>Net Proceeds (before COGS)</b>','smb'),
         P(f'<b>{fmt(T["net_proceed"])}</b>',
           S('np2', fontSize=8, textColor=GREEN if T['net_proceed']>=0 else RED, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
         P(f'<b>{pct(T["net_margin"])}</b>',
           S('nm2', fontSize=8, textColor=GREEN if T['net_margin']>=20 else RED,  fontName='Helvetica-Bold', alignment=TA_RIGHT))],
    ]

    if has_cogs_file:
        pnl_rows.append([
            P('  ↳ Cost of Goods Sold (COGS)','sm'),
            P(fmt(-T['total_cogs']),'pur_r'),
            P(pct(-T['cogs_pct_rev']),'pur_r'),
        ])
        pnl_rows.append([
            P('<b>TRUE NET PROFIT (after COGS)</b>','smb'),
            P(f'<b>{fmt(T["true_profit"])}</b>',
              S('tp2', fontSize=8, textColor=GREEN if T['true_profit']>=0 else RED, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
            P(f'<b>{pct(T["true_margin"])}</b>',
              S('tm2', fontSize=8, textColor=GREEN if T['true_margin']>=10 else RED, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        ])

    pnl = Table(pnl_rows, colWidths=[W*0.55, W*0.23, W*0.22])
    ps_style = base_ts()
    ps_style.add('BACKGROUND', (0,0),  (-1,0),  NAVY)
    ps_style.add('TEXTCOLOR',  (0,0),  (-1,0),  WHITE)
    last = len(pnl_rows) - 1
    if has_cogs_file:
        # net proceeds row highlight
        ps_style.add('BACKGROUND', (0,last-2), (-1,last-2), PALE_GREEN if T['net_proceed']>=0 else PALE_RED)
        ps_style.add('LINEABOVE',  (0,last-2), (-1,last-2), 1, MID_GRAY)
        # true profit row highlight (purple tint)
        ps_style.add('BACKGROUND', (0,last),   (-1,last),   PALE_PURPLE)
        ps_style.add('LINEABOVE',  (0,last),   (-1,last),   2, PURPLE)
    else:
        ps_style.add('BACKGROUND', (0,last), (-1,last), PALE_GREEN if T['net_proceed']>=0 else PALE_RED)
        ps_style.add('LINEABOVE',  (0,last), (-1,last), 1.5, GREEN)
    pnl.setStyle(ps_style)
    story.append(pnl)
    story.append(Spacer(1, 0.3*cm))

    # ── Key Insights ──────────────────────────────────────────────────────────
    story.append(sec_hdr(W, '  KEY INSIGHTS & OBSERVATIONS', bg=TEAL))
    story.append(Spacer(1, 0.05*cm))

    best      = max(products, key=lambda x: x['net_proceed'])
    worst     = min(products, key=lambda x: x['net_proceed'])
    most_sold = max(products, key=lambda x: x['units_sold'])
    high_ret  = max(products, key=lambda x: x['return_rate'])
    adv_asins = [p for p in products if p['advertising'] > 0]
    loss_asins= [p for p in products if p['net_proceed'] < 0]

    insights = [
        f"<b>Top Revenue Driver:</b> {most_sold['short']} ({most_sold['asin']}) — "
        f"{most_sold['units_sold']} units, {fmt(most_sold['total_sales'])} revenue, "
        f"{fmt(most_sold['net_proceed'])} net proceeds.",

        f"<b>Best Net Proceeds:</b> {best['short'][:50]} ({best['asin']}) — "
        f"{fmt(best['net_proceed'])} at {pct(best['net_margin'])} margin.",
    ]

    if worst['net_proceed'] < 0:
        insights.append(
            f"<b>&#9888; LOSS ALERT:</b> {worst['short'][:45]} ({worst['asin']}) — "
            f"NEGATIVE net proceeds of {fmt(worst['net_proceed'])}. "
            f"Return rate: {pct(worst['return_rate']*100)}. Immediate action required."
        )

    if high_ret['return_rate'] > 0.10:
        insights.append(
            f"<b>&#9888; HIGH RETURNS:</b> {high_ret['short'][:45]} ({high_ret['asin']}) — "
            f"{pct(high_ret['return_rate']*100)} return rate "
            f"({high_ret['refunded']} of {high_ret['units_sold']} units). Investigate listing/product."
        )

    if adv_asins:
        tot_adv = sum(p['advertising'] for p in adv_asins)
        tot_rev = sum(p['total_sales'] for p in adv_asins)
        acos    = tot_adv / tot_rev * 100 if tot_rev else 0
        insights.append(
            f"<b>Advertising (blended ACoS):</b> {fmt(tot_adv)} spend, ACoS = {pct(acos)}. "
            f"{'Consider bid optimisation — ACoS above 20% erodes margins.' if acos > 20 else 'ACoS is healthy.'}"
        )

    if has_cogs_file:
        if T['cogs_missing'] > 0:
            insights.append(
                f"<b>&#9888; COGS Gap:</b> {T['cogs_missing']} ASIN(s) have no COGS data — "
                f"true profit is understated. Fill in cogs.csv for complete analysis."
            )
        insights.append(
            f"<b>True Net Profit (after COGS):</b> {fmt(T['true_profit'])} at "
            f"{pct(T['true_margin'])} true margin. "
            f"COGS consumed {pct(T['cogs_pct_rev'])} of revenue ({fmt(T['total_cogs'])})."
        )
    else:
        insights.append(
            "<b>COGS not provided.</b> Run with a cogs.csv to see true profit after product costs. "
            "Use --create-cogs flag to generate a starter template."
        )

    ins_t = Table([[P(f'&#8226;  {i}', 'body')] for i in insights], colWidths=[W])
    ins_t.setStyle(TableStyle([
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [WHITE, PALE_TEAL]),
        ('BOX',            (0,0), (-1,-1), 0.5, MID_GRAY),
        ('TOPPADDING',     (0,0), (-1,-1), 5),
        ('BOTTOMPADDING',  (0,0), (-1,-1), 5),
        ('LEFTPADDING',    (0,0), (-1,-1), 10),
        ('RIGHTPADDING',   (0,0), (-1,-1), 10),
    ]))
    story.append(ins_t)
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 2 — PRODUCT BREAKDOWN + COST STRUCTURE
    # ══════════════════════════════════════════════════════════════════════════
    story.append(sec_hdr(W, '  PRODUCT-LEVEL PERFORMANCE BREAKDOWN'))
    story.append(Spacer(1, 0.15*cm))

    cw1 = [W*0.25, W*0.09, W*0.07, W*0.07, W*0.07, W*0.07, W*0.09, W*0.09, W*0.09, W*0.11]

    def th(txt, align=TA_RIGHT):
        _style_counter[0] += 1
        return P(f'<b>{txt}</b>', S(f'th_{_style_counter[0]}', fontSize=7.5,
                                    textColor=WHITE, fontName='Helvetica-Bold', alignment=align))

    prod_rows = [[
        th('Product (ASIN)', TA_LEFT), th('Gross\nRevenue'), th('Units\nSold'),
        th('Net\nUnits'), th('COGS/\nUnit'), th('Return\nRate'),
        th('Fulfillment'), th('Sell\nFees'), th('Advert.'), th('Net\nProceeds'),
    ]]

    for p in ps:
        rr_c = RED if p['return_rate']>0.15 else (ORANGE if p['return_rate']>0.05 else GREEN)
        np_c = RED if p['net_proceed']<0    else (ORANGE if p['net_proceed']<200   else GREEN)
        _style_counter[0] += 1; n = _style_counter[0]
        prod_rows.append([
            P(f"<b>{p['short']}</b><br/><font size='6.5' color='#6C757D'>{p['asin']}</font>",
              S(f'pn_{n}', fontSize=7.5, textColor=DARK, fontName='Helvetica', leading=11)),
            P(fmt(p['total_sales']), 'sm_r'),
            P(str(p['units_sold']), 'sm_c'),
            P(str(p['net_units']),  'sm_c'),
            P(fmt(p['cogs_per_unit']) if p['has_cogs'] else '-', 'sm_r'),
            P(pct(p['return_rate']*100), S(f'rr_{n}', fontSize=7.5, textColor=rr_c, fontName='Helvetica-Bold', alignment=TA_CENTER)),
            P(fmt(p['fulfillment']),  'sm_r'),
            P(fmt(p['selling_fees']), 'sm_r'),
            P(fmt(p['advertising']) if p['advertising']>0 else '-', 'sm_r'),
            P(fmt(p['net_proceed']),  S(f'np_{n}', fontSize=7.5, textColor=np_c, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
        ])

    prod_rows.append([
        P('<b>TOTAL</b>','smb'),
        P(f'<b>{fmt(T["total_sales"])}</b>',   'smb_r'),
        P(f'<b>{T["units_sold"]}</b>',          'smb_c'),
        P(f'<b>{T["net_units"]}</b>',            'smb_c'),
        P('—', 'sm_c'),
        P(f'<b>{pct(T["return_rate"]*100)}</b>','smb_c'),
        P(f'<b>{fmt(T["fulfillment"])}</b>',     'smb_r'),
        P(f'<b>{fmt(T["selling_fees"])}</b>',    'smb_r'),
        P(f'<b>{fmt(T["advertising"])}</b>',     'smb_r'),
        P(f'<b>{fmt(T["net_proceed"])}</b>',     'smb_r'),
    ])

    pt = Table(prod_rows, colWidths=cw1)
    pt_s = base_ts()
    pt_s.add('BACKGROUND', (0,0),  (-1,0),  NAVY)
    pt_s.add('TEXTCOLOR',  (0,0),  (-1,0),  WHITE)
    pt_s.add('BACKGROUND', (0,-1), (-1,-1), PALE_GREEN)
    pt_s.add('LINEABOVE',  (0,-1), (-1,-1), 1.5, GREEN)
    pt.setStyle(pt_s)
    story.append(pt)
     #story.append(Spacer(1, 0.3*cm))
    story.append(PageBreak())

    # ── Cost Structure ────────────────────────────────────────────────────────
    story.append(sec_hdr(W, '  COST STRUCTURE ANALYSIS (per ASIN)', bg=ORANGE))
    story.append(Spacer(1, 0.2*cm))

    cw2 = [W*0.13, W*0.12, W*0.12, W*0.12, W*0.12, W*0.12, W*0.12, W*0.15]

    def cth(txt):
        _style_counter[0] += 1
        return P(f'<b>{txt}</b>', S(f'cth_{_style_counter[0]}', fontSize=7.5,
                                    textColor=WHITE, fontName='Helvetica-Bold', alignment=TA_CENTER))

    cost_rows = [[cth('ASIN'), cth('Gross\nRevenue'), cth('Fulfil\n% Rev'),
                  cth('Fees\n% Rev'), cth('Advert\n% Rev'), cth('TotCost\n% Rev'),
                  cth('Net\nMargin'), cth('Status')]]

    for p in ps:
        rev   = p['total_sales'] or 1
        f_p   = p['fulfillment']  / rev * 100
        s_p   = p['selling_fees'] / rev * 100
        a_p   = p['advertising']  / rev * 100
        tc_p  = p['total_costs']  / rev * 100
        nm    = p['net_margin']

        if p['net_proceed'] < 0:        flag, fc = '🔴 LOSS',     RED
        elif p['return_rate'] > 0.15:   flag, fc = '⚠️ HIGH RTN', RED
        elif nm < 20:                   flag, fc = '🟡 LOW MGN',  ORANGE
        elif a_p  > 20:                 flag, fc = '🟡 HIGH ACoS',ORANGE
        else:                           flag, fc = '🟢 OK',       GREEN

        _style_counter[0] += 1; n = _style_counter[0]
        cost_rows.append([
            P(p['asin'], 'sm'),
            P(fmt(rev), 'sm_r'),
            color_pct(f_p, 30, 20),
            color_pct(s_p, 15, 10),
            color_pct(a_p, 30, 15),
            color_pct(tc_p,60, 40),
            P(pct(nm), S(f'nm_{n}', fontSize=7.5,
                         textColor=RED if nm<20 else (ORANGE if nm<40 else GREEN),
                         fontName='Helvetica-Bold', alignment=TA_CENTER)),
            P(flag, S(f'fl_{n}', fontSize=7.5, textColor=fc,
                      fontName='Helvetica-Bold', alignment=TA_CENTER)),
        ])

    cost_rows.append([
        P('<b>TOTAL</b>','smb'),
        P(f'<b>{fmt(T["total_sales"])}</b>','smb_r'),
        P(f'<b>{pct(T["fulfillment"]/T["total_sales"]*100)}</b>','smb_c'),
        P(f'<b>{pct(T["selling_fees"]/T["total_sales"]*100)}</b>','smb_c'),
        P(f'<b>{pct(T["advertising"]/T["total_sales"]*100)}</b>','smb_c'),
        P(f'<b>{pct(T["total_costs"]/T["total_sales"]*100)}</b>','smb_c'),
        P(f'<b>{pct(T["net_margin"])}</b>','smb_c'),
        P('', 'sm'),
    ])

    ct = Table(cost_rows, colWidths=cw2)
    ct_s = base_ts()
    ct_s.add('BACKGROUND', (0,0),  (-1,0),  ORANGE)
    ct_s.add('TEXTCOLOR',  (0,0),  (-1,0),  WHITE)
    ct_s.add('BACKGROUND', (0,-1), (-1,-1), PALE_GREEN)
    ct_s.add('LINEABOVE',  (0,-1), (-1,-1), 1.5, GREEN)
    ct.setStyle(ct_s)
    story.append(ct)
    story.append(Spacer(1, 0.2*cm))

    # legend
    leg = Table([[
        P('<b>Color Key:</b>', 'smb'),
        P('🟢 Good / Within range',       S('lg', fontSize=7.5, textColor=GREEN,  fontName='Helvetica')),
        P('🟡 Needs attention',            S('lo', fontSize=7.5, textColor=ORANGE, fontName='Helvetica')),
        P('🔴 Critical / Action needed',   S('lr', fontSize=7.5, textColor=RED,    fontName='Helvetica')),
    ]], colWidths=[W*0.15, W*0.23, W*0.28, W*0.34])
    leg.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1), LIGHT_GRAY),
        ('BOX',(0,0),(-1,-1),0.5,MID_GRAY),
        ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
        ('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
    ]))
    story.append(leg)
    story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # PAGE 3 — TRUE PROFIT ANALYSIS (only if COGS provided)
    # ══════════════════════════════════════════════════════════════════════════
    if has_cogs_file:
        story.append(sec_hdr(W, '  TRUE PROFIT ANALYSIS  (Net Proceeds − Cost of Goods Sold)', bg=PURPLE))
        story.append(Spacer(1, 0.2*cm))

        # Explainer box
        expl = Table([[P(
            '<b>How to read this page:</b>  '
            'Net Proceeds = Revenue after Amazon fees (fulfillment, referral, ads, storage). '
            'COGS = your sourcing/manufacturing cost × net units sold. '
            '<b>True Net Profit = Net Proceeds − COGS</b> — this is what actually goes in your pocket. '
            + (f'  ⚠️  {T["cogs_missing"]} ASIN(s) are missing COGS — their True Profit = Net Proceeds (overstated).' if T['cogs_missing'] else ''),
            'body'
        )]], colWidths=[W])
        expl.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1), PALE_PURPLE),
            ('LINEBEFORE',(0,0),(0,-1),  4, PURPLE),
            ('TOPPADDING',(0,0),(-1,-1), 7),('BOTTOMPADDING',(0,0),(-1,-1), 7),
            ('LEFTPADDING',(0,0),(-1,-1),10),('RIGHTPADDING',(0,0),(-1,-1),10),
        ]))
        story.append(expl)
        story.append(Spacer(1, 0.25*cm))

        # ── True Profit per ASIN table ──────────────────────────────────────
        cw3 = [W*0.28, W*0.09, W*0.10, W*0.10, W*0.10, W*0.11, W*0.11, W*0.11]

        def ph(txt):
            _style_counter[0] += 1
            return P(f'<b>{txt}</b>', S(f'ph_{_style_counter[0]}', fontSize=7.5,
                                        textColor=WHITE, fontName='Helvetica-Bold', alignment=TA_RIGHT))

        tp_rows = [[
            P('<b>Product (ASIN)</b>', S('ph0', fontSize=7.5, textColor=WHITE, fontName='Helvetica-Bold')),
            ph('Net\nUnits'),
            ph('COGS/\nUnit'),
            ph('Total\nCOGS'),
            ph('Net\nProceeds'),
            ph('True\nProfit'),
            ph('True\nMargin'),
            ph('Status'),
        ]]

        for p in ps:
            has  = p['has_cogs']
            tp   = p['true_profit']
            tm   = p['true_margin']
            cp   = p['cogs_per_unit']
            tc   = p['total_cogs']
            np_  = p['net_proceed']

            # Flag
            if not has:
                flag2, fc2 = '❓ NO COGS', GRAY
            elif tp < 0:
                flag2, fc2 = '🔴 LOSS',    RED
            elif tm < 10:
                flag2, fc2 = '🟡 THIN',    ORANGE
            elif tm < 25:
                flag2, fc2 = '🟡 OK',      ORANGE
            else:
                flag2, fc2 = '🟢 GOOD',    GREEN

            _style_counter[0] += 1; n = _style_counter[0]
            tp_rows.append([
                P(f"<b>{p['short'][:50]}</b><br/><font size='6.5' color='#6C757D'>{p['asin']}</font>",
                  S(f'tp_{n}', fontSize=7.5, textColor=DARK, fontName='Helvetica', leading=11)),
                P(str(p['net_units']), 'sm_c'),
                P(fmt(cp) if has else '—', 'sm_r' if has else 'sm_c'),
                P(fmt(tc) if has else '—', S(f'tc_{n}', fontSize=8, textColor=PURPLE, fontName='Helvetica-Bold', alignment=TA_RIGHT) if has else 'sm_c'),
                P(fmt(np_), 'sm_r'),
                color_val(tp, good_above=0, warn_below=-500) if has else P(fmt(np_), 'sm_r'),
                P(pct(tm), S(f'tm_{n}', fontSize=7.5,
                              textColor=RED if tm<10 else (ORANGE if tm<25 else GREEN),
                              fontName='Helvetica-Bold', alignment=TA_RIGHT)) if has else P('—', 'sm_c'),
                P(flag2, S(f'f2_{n}', fontSize=7.5, textColor=fc2,
                           fontName='Helvetica-Bold', alignment=TA_CENTER)),
            ])

        tp_rows.append([
            P('<b>TOTAL</b>','smb'),
            P(f'<b>{T["net_units"]}</b>','smb_c'),
            P('—','sm_c'),
            P(f'<b>{fmt(T["total_cogs"])}</b>',
              S('tot_tc', fontSize=8, textColor=PURPLE, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
            P(f'<b>{fmt(T["net_proceed"])}</b>','smb_r'),
            P(f'<b>{fmt(T["true_profit"])}</b>',
              S('tot_tp', fontSize=8,
                textColor=GREEN if T['true_profit']>=0 else RED,
                fontName='Helvetica-Bold', alignment=TA_RIGHT)),
            P(f'<b>{pct(T["true_margin"])}</b>',
              S('tot_tm', fontSize=8,
                textColor=GREEN if T['true_margin']>=10 else RED,
                fontName='Helvetica-Bold', alignment=TA_RIGHT)),
            P('', 'sm'),
        ])

        tpt = Table(tp_rows, colWidths=cw3)
        tpt_s = base_ts()
        tpt_s.add('BACKGROUND', (0,0),  (-1,0),  PURPLE)
        tpt_s.add('TEXTCOLOR',  (0,0),  (-1,0),  WHITE)
        tpt_s.add('BACKGROUND', (0,-1), (-1,-1), PALE_PURPLE)
        tpt_s.add('LINEABOVE',  (0,-1), (-1,-1), 2, PURPLE)
        tpt.setStyle(tpt_s)
        story.append(tpt)
       # story.append(Spacer(1, 0.3*cm))
        story.append(PageBreak())

        # ── COGS waterfall summary ───────────────────────────────────────────
        story.append(sec_hdr(W, '  PROFIT WATERFALL SUMMARY', bg=NAVY))
        story.append(Spacer(1, 0.1*cm))

        wf_rows = [
            [P('<b>Step</b>','smb'), P('<b>Amount</b>','smb_r'), P('<b>% of Revenue</b>','smb_r'), P('<b>Cumulative</b>','smb_r')],
            [P('Gross Revenue',    'sm'), P(fmt(T['total_sales']),   'sm_r'),  P('100.0%','sm_r'),          P(fmt(T['total_sales']),'sm_r')],
            [P('− Fulfillment',    'sm'), P(fmt(-T['fulfillment']),  'red_r'), P(pct(-T['fulfillment'] /T['total_sales']*100),'red_r'), P(fmt(T['total_sales']-T['fulfillment']),'sm_r')],
            [P('− Selling Fees',   'sm'), P(fmt(-T['selling_fees']), 'red_r'), P(pct(-T['selling_fees']/T['total_sales']*100),'red_r'), P(fmt(T['total_sales']-T['fulfillment']-T['selling_fees']),'sm_r')],
            [P('− Advertising',    'sm'), P(fmt(-T['advertising']),  'red_r'), P(pct(-T['advertising'] /T['total_sales']*100),'red_r'), P(fmt(T['net_proceed']+T['storage']),'sm_r')],
            [P('− Storage',        'sm'), P(fmt(-T['storage']),      'red_r'), P(pct(-T['storage']     /T['total_sales']*100),'red_r'), P(fmt(T['net_proceed']),'sm_r')],
            [P('<b>= Net Proceeds</b>','smb'),
             P(f'<b>{fmt(T["net_proceed"])}</b>', S('wf_np', fontSize=8, textColor=GREEN if T['net_proceed']>=0 else RED, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
             P(f'<b>{pct(T["net_margin"])}</b>',  S('wf_nm', fontSize=8, textColor=GREEN if T['net_margin']>=20 else RED,  fontName='Helvetica-Bold', alignment=TA_RIGHT)),
             P(f'<b>{fmt(T["net_proceed"])}</b>',  S('wf_npc', fontSize=8, textColor=GREEN if T['net_proceed']>=0 else RED, fontName='Helvetica-Bold', alignment=TA_RIGHT))],
            [P('− COGS (product cost)','sm'),
             P(fmt(-T['total_cogs']), S('wf_cogs', fontSize=8, textColor=PURPLE, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
             P(pct(-T['cogs_pct_rev']), S('wf_cogsp', fontSize=8, textColor=PURPLE, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
             P(fmt(T['true_profit']), S('wf_tp_c', fontSize=8, textColor=GREEN if T['true_profit']>=0 else RED, fontName='Helvetica-Bold', alignment=TA_RIGHT))],
            [P('<b>= TRUE NET PROFIT</b>','smb'),
             P(f'<b>{fmt(T["true_profit"])}</b>', S('wf_tp', fontSize=8, textColor=GREEN if T['true_profit']>=0 else RED, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
             P(f'<b>{pct(T["true_margin"])}</b>', S('wf_tm', fontSize=8, textColor=GREEN if T['true_margin']>=10 else RED, fontName='Helvetica-Bold', alignment=TA_RIGHT)),
             P(f'<b>{fmt(T["true_profit"])}</b>', S('wf_tpc', fontSize=8, textColor=GREEN if T['true_profit']>=0 else RED, fontName='Helvetica-Bold', alignment=TA_RIGHT))],
        ]

        wf = Table(wf_rows, colWidths=[W*0.35, W*0.22, W*0.20, W*0.23])
        wf_s = base_ts()
        wf_s.add('BACKGROUND', (0,0),  (-1,0),  NAVY)
        wf_s.add('TEXTCOLOR',  (0,0),  (-1,0),  WHITE)
        wf_s.add('BACKGROUND', (0,6),  (-1,6),  PALE_GREEN if T['net_proceed']>=0 else PALE_RED)
        wf_s.add('LINEABOVE',  (0,6),  (-1,6),  1, MID_GRAY)
        wf_s.add('BACKGROUND', (0,-1), (-1,-1), PALE_PURPLE)
        wf_s.add('LINEABOVE',  (0,-1), (-1,-1), 2, PURPLE)
        wf.setStyle(wf_s)
        story.append(wf)

        if T['cogs_missing'] > 0:
            story.append(Spacer(1, 0.15*cm))
            warn_box = Table([[P(
                f'&#9888;  <b>Note:</b> {T["cogs_missing"]} ASIN(s) have no COGS entry. '
                f'Their True Profit equals Net Proceeds (product cost not deducted). '
                f'Update your cogs.csv to get the full picture.',
                S('warn', fontSize=8.5, textColor=colors.HexColor('#7B3F00'), fontName='Helvetica')
            )]], colWidths=[W])
            warn_box.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,-1), colors.HexColor('#FFF3CD')),
                ('BOX',(0,0),(-1,-1),0.5, colors.HexColor('#FFC107')),
                ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6),
                ('LEFTPADDING',(0,0),(-1,-1),10),
            ]))
            story.append(warn_box)

        story.append(PageBreak())

    # ══════════════════════════════════════════════════════════════════════════
    # LAST PAGE — RECOMMENDATIONS + ACTION CHECKLIST
    # ══════════════════════════════════════════════════════════════════════════
    story.append(sec_hdr(W, '  STRATEGIC RECOMMENDATIONS', bg=TEAL))
    story.append(Spacer(1, 0.2*cm))

    recs = []

    # Loss-making
    for p in loss_asins:
        recs.append((
            f'🔴 URGENT — Fix {p["asin"]} ({p["short"][:35]})',
            f'Net proceeds are NEGATIVE ({fmt(p["net_proceed"])}). '
            f'Return rate: {pct(p["return_rate"]*100)} ({p["refunded"]} of {p["units_sold"]} returned). '
            f'Actions: review return reasons in Seller Central, fix listing/photos, '
            f'check product quality, consider pausing until root cause resolved.',
            PALE_RED, RED
        ))

    # High returns
    for p in [p for p in products if p['return_rate']>0.10 and p not in loss_asins]:
        recs.append((
            f'🔴 HIGH RETURNS — {p["asin"]} ({pct(p["return_rate"]*100)})',
            f'{p["refunded"]} of {p["units_sold"]} units returned. '
            f'Home goods benchmark is 2–5%. Investigate return reasons, listing accuracy, and packaging.',
            PALE_RED, RED
        ))

    # Thin true-profit margin (only if COGS available)
    if has_cogs_file:
        for p in [p for p in products if p['has_cogs'] and 0 <= p['true_margin'] < 10]:
            recs.append((
                f'🟡 THIN MARGIN — {p["asin"]} (True margin: {pct(p["true_margin"])})',
                f'After COGS of {fmt(p["cogs_per_unit"])}/unit, true profit is only {fmt(p["true_profit"])}. '
                f'Options: (1) negotiate lower COGS with supplier, '
                f'(2) raise selling price by ₹20–50, '
                f'(3) reduce ad spend to protect margin.',
                PALE_ORANGE, ORANGE
            ))

    # High ACoS
    for p in adv_asins:
        acos = p['advertising']/p['total_sales']*100 if p['total_sales'] else 0
        if acos > 22:
            recs.append((
                f'🟡 HIGH ACoS — {p["asin"]} ({pct(acos)})',
                f'Ad spend {fmt(p["advertising"])} vs revenue {fmt(p["total_sales"])}. '
                f'Target ACoS < 18–20%. Pause low-converting keywords, reduce bids, test exact-match.',
                PALE_ORANGE, ORANGE
            ))

    # Scale top seller
    recs.append((
        f'🟢 SCALE — Double Down on {most_sold["asin"]}',
        f'Top driver: {fmt(most_sold["total_sales"])} revenue, {most_sold["units_sold"]} units. '
        f'Maintain 30–45 day stock coverage, launch bundle variants (2/3/5-pack), '
        f'test Sponsored Brand & Display ads to capture upper-funnel traffic.',
        PALE_GREEN, GREEN
    ))

    # Organic sellers
    organic = [p for p in products if p['advertising']==0 and p['units_sold']>=3]
    if organic:
        names = ', '.join(p['asin'] for p in organic[:3])
        recs.append((
            '🟢 GROWTH — Test Ads on Organic Sellers',
            f'{names} sell without ads. A Rs 200–500/day PPC test can accelerate rank. '
            f'Higher velocity → better organic rank → lower long-term ACoS.',
            PALE_GREEN, GREEN
        ))

    for title_r, body_r, bg_r, bdr in recs:
        rt = Table([
            [P(f'<b>{title_r}</b>', S(f'rt_{id(title_r)}', fontSize=9.5, textColor=bdr, fontName='Helvetica-Bold'))],
            [P(body_r, 'body')],
        ], colWidths=[W])
        rt.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(-1,-1), bg_r),
            ('LINEBEFORE',    (0,0),(0,-1),  4, bdr),
            ('TOPPADDING',    (0,0),(-1,0),  8),('BOTTOMPADDING',(0,0),(-1,0),  3),
            ('TOPPADDING',    (0,1),(-1,-1), 3),('BOTTOMPADDING',(0,-1),(-1,-1),8),
            ('LEFTPADDING',   (0,0),(-1,-1), 10),('RIGHTPADDING', (0,0),(-1,-1),10),
            ('BOX',           (0,0),(-1,-1), 0.3, MID_GRAY),
        ]))
        story.append(rt)
        story.append(Spacer(1, 0.15*cm))

    story.append(Spacer(1, 0.15*cm))
    story.append(sec_hdr(W, '  ACTION CHECKLIST', bg=NAVY))
    story.append(Spacer(1, 0.1*cm))

    actions = [
        ('This Week', [
            'Download return reasons for flagged ASINs: Seller Central → Orders → Returns',
            'Check inventory — flag any ASIN with < 15 days stock coverage',
            'Review negative reviews on loss/high-return ASINs',
        ]),
        ('This Month', [
            'Optimise PPC — pause keywords with ACoS > 40% for 2 weeks',
            'Update cogs.csv with latest sourcing costs if prices changed',
            'Test small ad campaigns on top organic sellers (Rs 200/day budget)',
        ]),
        ('Next Quarter', [
            'Launch bundle/multi-pack variants of your top-selling ASIN',
            'Add A+ Content & enhanced images for low-velocity ASINs',
            'Renegotiate COGS with suppliers for SKUs with thin true margins',
        ]),
    ]

    at = Table([[
        P(f'<b>{lbl}</b>', S(f'al_{lbl[:4]}', fontSize=9, textColor=NAVY, fontName='Helvetica-Bold')),
        P('<br/>'.join(f'&#9744;  {i}' for i in items),
          S(f'ai_{lbl[:4]}', fontSize=8.5, textColor=DARK, fontName='Helvetica', leading=14)),
    ] for lbl, items in actions], colWidths=[W*0.18, W*0.82])
    at.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(0,-1), LIGHT_BLUE),
        ('ROWBACKGROUNDS',(1,0),(1,-1), [WHITE, LIGHT_GRAY, WHITE]),
        ('BOX',           (0,0),(-1,-1),0.5,MID_GRAY),
        ('INNERGRID',     (0,0),(-1,-1),0.3,MID_GRAY),
        ('TOPPADDING',    (0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),
        ('LEFTPADDING',   (0,0),(-1,-1),10),('RIGHTPADDING', (0,0),(-1,-1),10),
        ('VALIGN',        (0,0),(-1,-1),'TOP'),
    ]))
    story.append(at)
    story.append(Spacer(1, 0.4*cm))

    # Footer
    story.append(HRFlowable(width='100%', thickness=0.5, color=NAVY))
    story.append(Spacer(1, 0.1*cm))
    story.append(P(
        f'Homnix Business Intelligence Report  |  {period}  |  Generated: {run_date}  |  '
        f'Amazon India  |  {len(products)} ASINs  |  '
        f'{"COGS included — True Profit reflects actual margin after product cost." if has_cogs_file else "COGS not provided — supply cogs.csv for true profit analysis."}',
        'note'
    ))

    doc.build(story)
    return output_path


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main():
    args = sys.argv[1:]

    # Special flag: create starter COGS CSV from last known products
    if '--create-cogs' in args:
        xlsx_args = [a for a in args if a.endswith('.xlsx')]
        if not xlsx_args:
            print("ERROR: Provide the xlsx file too:  python homnix_report_generator.py report.xlsx --create-cogs")
            sys.exit(1)
        products, _, _ = load_products(xlsx_args[0])
        out = create_cogs_template(products)
        print(f"\n  Starter COGS template created:  {out}")
        print(f"  Open it, fill in the 'cogs_per_unit' column, and save.")
        print(f"  Then run:  python homnix_report_generator.py {xlsx_args[0]} {out}\n")
        return

    if not args:
        print(__doc__)
        sys.exit(1)

    # Parse positional args:
    # arg 1 MUST be the Amazon report (.xlsx or .csv)
    # arg 2 (optional) MUST be the COGS file (.csv) or an output directory if it doesn't end in .csv
    # arg 3 (optional) is the output directory
    
    report_path = None
    cogs_path   = None
    out_dir     = None

    if len(args) > 0:
        report_path = args[0]
    if len(args) > 1:
        if args[1].endswith('.csv'):
            cogs_path = args[1]
        else:
            out_dir = args[1]
    if len(args) > 2:
        out_dir = args[2]

    if not report_path:
        print("ERROR: No report file provided.")
        sys.exit(1)

    if not os.path.exists(report_path):
        print(f"ERROR: File not found — {report_path}")
        sys.exit(1)

    if cogs_path and not os.path.exists(cogs_path):
        print(f"WARNING: COGS file not found — {cogs_path}. Continuing without COGS.")
        cogs_path = None

    print(f"\n  Loading Amazon report:  {report_path}")
    cogs_map = load_cogs(cogs_path)
    if cogs_path:
        print(f"  Loading COGS file:      {cogs_path}  ({len(cogs_map)//2} entries)")

    products, totals, period = load_products(report_path, cogs_map)

    print(f"\n  Period   : {period}")
    print(f"  ASINs    : {len(products)}")
    print(f"  Revenue  : {fmt(totals['total_sales'])}")
    print(f"  Net Proc : {fmt(totals['net_proceed'])}  ({pct(totals['net_margin'])} margin)")
    if cogs_path:
        print(f"  COGS     : {fmt(totals['total_cogs'])}  ({totals['cogs_covered']} ASINs covered, {totals['cogs_missing']} missing)")
        print(f"  TRUE PROFIT: {fmt(totals['true_profit'])}  ({pct(totals['true_margin'])} true margin)")

    # Determine output directory
    safe_period = period.replace(' ', '-')
    pdf_name    = f'Homnix_Report_{safe_period}.pdf'

    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    else:
        candidate = os.path.dirname(os.path.abspath(report_path))
        try:
            tp = os.path.join(candidate, '.wtest')
            open(tp,'w').close(); os.remove(tp)
            out_dir = candidate
        except OSError:
            out_dir = os.getcwd()

    outfile = os.path.join(out_dir, pdf_name)

    print(f"\n  Building PDF...")
    build_report(products, totals, period, outfile, has_cogs_file=bool(cogs_path))
    print(f"\n  DONE!  Saved to:\n  {outfile}\n")


if __name__ == '__main__':
    main()
