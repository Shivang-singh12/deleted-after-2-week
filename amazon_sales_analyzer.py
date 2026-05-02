"""
Amazon Sales Report Analyzer
==============================
Analyzes the Amazon Seller "Sales & Traffic by ASIN" report (tortuga/EU format)
and generates a clean, concise CSV with:
  1. Overall Sales Summary
  2. Product-by-Product Breakdown
  3. Key Insights & Recommendations

HOW TO USE:
-----------
1. Install dependencies:
       pip install pandas openpyxl

2. Change INPUT_FILE to your report path:
       INPUT_FILE = "your_report.xlsx"

3. Run:
       python amazon_sales_analyzer.py

4. Output CSV will be saved as:
       Amazon_Sales_Report_<StartDate>_<EndDate>.csv
"""

import pandas as pd
import csv
import os
import sys
from datetime import datetime

# ─────────────────────────────────────────────
#  CONFIGURATION — Change these as needed
# ─────────────────────────────────────────────
INPUT_FILE = "9fcf3b99-b62b-468c-b9c4-a0a61e9ebd17.amzn1.tortuga.4.eu.CSV"   # <-- Change this to your file path
SHEET_NAME = "in"                                                               
OUTPUT_DIR = "."                                                                
# ─────────────────────────────────────────────


def load_data(filepath, sheet):
    """Load the report (CSV or Excel) into a DataFrame."""
    if not os.path.exists(filepath):
        print(f"ERROR: File not found: {filepath}")
        sys.exit(1)
    try:
        ext = os.path.splitext(filepath)[1].lower()
        if ext == '.csv':
            df = pd.read_csv(filepath)
        else:
            df = pd.read_excel(filepath, sheet_name=sheet)
    except Exception as e:
        print(f"ERROR reading file: {e}")
        sys.exit(1)
    return df


def safe(val, default=0.0):
    """Return 0 if value is NaN, else return the value."""
    try:
        if pd.isna(val):
            return default
    except Exception:
        pass
    return val


def safe_col_sum(df, col, default=0.0):
    """Sum a column if it exists, else return default."""
    return df[col].apply(safe).sum() if col in df.columns else default


def analyze(df):
    """Extract all key metrics from the DataFrame."""

    # ── Date range ──────────────────────────────────
    try:
        start_date = pd.to_datetime(df['Start date'].iloc[0]).strftime('%d %b %Y') if 'Start date' in df.columns else "N-A"
    except Exception:
        start_date = "N-A"
    try:
        end_date = pd.to_datetime(df['End date'].iloc[0]).strftime('%d %b %Y') if 'End date' in df.columns else "N-A"
    except Exception:
        end_date = "N-A"
    currency = df['Currency code'].iloc[0] if 'Currency code' in df.columns else "INR"

    # ── Overall totals ───────────────────────────────
    gross_sales    = safe_col_sum(df, 'Sales')
    net_sales      = safe_col_sum(df, 'Net sales')
    returns_amt    = gross_sales - net_sales                          # positive number = amount returned
    units_sold     = int(safe_col_sum(df, 'Units sold'))
    units_returned = int(safe_col_sum(df, 'Units returned'))
    net_units      = int(safe_col_sum(df, 'Net units sold'))

    easy_ship_fees = safe_col_sum(df, 'Manufacturer-Fulfilled Postage Fee total')
    referral_fees  = safe_col_sum(df, 'Referral fee total')
    ad_spend       = safe_col_sum(df, 'Sponsored Products charge total')
    total_fees     = easy_ship_fees + referral_fees + ad_spend
    net_proceeds   = safe_col_sum(df, 'Net proceeds total')

    summary = {
        'start_date':    start_date,
        'end_date':      end_date,
        'currency':      currency,
        'gross_sales':   round(gross_sales,   2),
        'returns_amt':   round(-returns_amt,  2),   # stored as negative
        'net_sales':     round(net_sales,     2),
        'easy_ship':     round(-easy_ship_fees, 2),
        'referral_fees': round(-referral_fees,  2),
        'ad_spend':      round(-ad_spend,       2),
        'total_fees':    round(-total_fees,     2),
        'net_proceeds':  round(net_proceeds,  2),
        'units_sold':    units_sold,
        'units_returned':units_returned,
        'net_units':     net_units,
    }

    # ── Per-product breakdown ────────────────────────
    products = []
    for _, row in df.iterrows():
        msku          = row.get('MSKU', 'N/A')
        u_sold        = int(safe(row.get('Units sold', 0), 0))
        u_returned    = int(safe(row.get('Units returned', 0), 0))
        u_net         = int(safe(row.get('Net units sold', 0), 0))
        avg_price     = round(safe(row.get('Average sales price', 0)), 2)
        prod_net_sales= round(safe(row.get('Net sales', 0)), 2)
        ship_fee      = round(safe(row.get('Manufacturer-Fulfilled Postage Fee total', 0)), 2)
        ref_fee       = round(safe(row.get('Referral fee total', 0)), 2)
        ads           = round(safe(row.get('Sponsored Products charge total', 0)), 2)
        net_proc      = round(safe(row.get('Net proceeds total', 0)), 2)

        # Only include rows that have any activity
        if u_sold > 0 or u_returned > 0 or net_proc != 0:
            acos = round((ads / prod_net_sales * 100), 1) if prod_net_sales > 0 and ads > 0 else None
            products.append({
                'msku':         msku,
                'units_sold':   u_sold,
                'units_returned': u_returned,
                'net_units':    u_net,
                'avg_price':    avg_price,
                'net_sales':    prod_net_sales,
                'easy_ship':    ship_fee,
                'referral_fee': ref_fee,
                'ad_spend':     ads,
                'net_proceeds': net_proc,
                'acos_pct':     acos,
            })

    # Sort by net proceeds descending
    products.sort(key=lambda x: x['net_proceeds'], reverse=True)

    # ── Key Insights ─────────────────────────────────
    insights = []

    # ACOS alert — any product with ACOS > 40%
    high_acos = [p for p in products if p['acos_pct'] and p['acos_pct'] > 40]
    for p in high_acos:
        insights.append((
            "HIGH ACOS Alert",
            f"{p['msku']}: {currency} {p['ad_spend']} ad spend vs {currency} {p['net_sales']} net sales = {p['acos_pct']}% ACOS. Review ad campaigns immediately."
        ))

    # Loss-making SKUs
    loss_skus = [p for p in products if p['net_proceeds'] < 0]
    if loss_skus:
        desc = "; ".join([f"{p['msku']} (lost {currency} {abs(p['net_proceeds'])})" for p in loss_skus])
        insights.append(("Loss-Making SKUs", desc))

    # Best margin SKU (no ads, highest proceeds)
    no_ad_products = [p for p in products if p['ad_spend'] == 0 and p['net_proceeds'] > 0]
    if no_ad_products:
        best = no_ad_products[0]
        insights.append((
            "Best Margin SKU (No Ads)",
            f"{best['msku']}: {currency} {best['net_proceeds']} net proceeds from {best['net_units']} units with ZERO ad spend."
        ))

    # Easy ship cost share
    if gross_sales > 0:
        ship_pct = round(easy_ship_fees / gross_sales * 100, 1)
        insights.append((
            "Easy Ship Cost Share",
            f"Total Easy Ship fees: {currency} {round(easy_ship_fees, 2)} = {ship_pct}% of gross sales."
        ))

    # Ad spend as % of net sales
    if net_sales > 0 and ad_spend > 0:
        total_acos = round(ad_spend / net_sales * 100, 1)
        insights.append((
            "Overall Ad Spend",
            f"Total ad spend {currency} {round(ad_spend, 2)} = {total_acos}% of total net sales. {'CRITICAL — reduce spend.' if total_acos > 40 else 'Monitor closely.'}"
        ))

    if not insights:
        insights.append(("General", "No critical issues detected. Review individual SKU performance for optimisation."))

    return summary, products, insights


def write_csv(summary, products, insights, output_dir):
    """Write everything into a single merged CSV file."""
    s   = summary
    cur = s['currency']
    start_clean = s['start_date'].replace(' ', '')
    end_clean   = s['end_date'].replace(' ', '')
    # Remove characters that are not safe in filenames (Windows / Linux)
    for ch in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        start_clean = start_clean.replace(ch, '-')
        end_clean   = end_clean.replace(ch, '-')
    generated_on = datetime.now().strftime('%d%b%Y')
    filename = f"Amazon_Sales_Report_{start_clean}_{end_clean}_Generated_{generated_on}.csv"
    filepath = os.path.join(output_dir, filename)

    with open(filepath, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)

        # ── Header ──────────────────────────────────
        w.writerow([f"AMAZON SALES ANALYSIS REPORT"])
        w.writerow([f"Period: {s['start_date']} – {s['end_date']}  |  Currency: {cur}"])
        w.writerow([])

        # ── Section 1: Sales Summary ─────────────────
        w.writerow(["=" * 50])
        w.writerow(["SECTION 1: SALES SUMMARY"])
        w.writerow(["=" * 50])
        w.writerow([])
        w.writerow(["Metric", f"Amount ({cur})"])
        w.writerow([f"Gross Product Sales ({s['units_sold']} orders)",  s['gross_sales']])
        w.writerow([f"Less: Returns ({s['units_returned']} units)",      s['returns_amt']])
        w.writerow([f"Net Sales ({s['net_units']} net units)",           s['net_sales']])
        w.writerow([])
        w.writerow(["FEES & DEDUCTIONS", f"Amount ({cur})"])
        w.writerow(["Easy Ship / Postage Fees",   s['easy_ship']])
        w.writerow(["Referral Fees",              s['referral_fees']])
        w.writerow(["Sponsored Products (Ads)",   s['ad_spend']])
        w.writerow(["Total Fees",                 s['total_fees']])
        w.writerow([])
        w.writerow(["NET PROCEEDS TO SELLER",     s['net_proceeds']])
        w.writerow([])
        w.writerow(["UNITS SUMMARY", "Count"])
        w.writerow(["Units Sold",     s['units_sold']])
        w.writerow(["Units Returned", s['units_returned']])
        w.writerow(["Net Units Sold", s['net_units']])
        w.writerow([])
        w.writerow([])

        # ── Section 2: Product Breakdown ─────────────
        w.writerow(["=" * 50])
        w.writerow(["SECTION 2: PRODUCT-BY-PRODUCT BREAKDOWN"])
        w.writerow(["=" * 50])
        w.writerow([])
        w.writerow([
            "Product (MSKU)",
            "Units Sold", "Returns", "Net Units",
            f"Avg Price ({cur})",
            f"Net Sales ({cur})",
            f"Easy Ship ({cur})",
            f"Referral Fee ({cur})",
            f"Ad Spend ({cur})",
            f"Net Proceeds ({cur})",
            "ACOS (%)"
        ])

        tot_sold = tot_ret = tot_net = 0
        tot_ns = tot_es = tot_rf = tot_ads = tot_np = 0.0

        for p in products:
            w.writerow([
                p['msku'],
                p['units_sold'],
                p['units_returned'],
                p['net_units'],
                p['avg_price'],
                p['net_sales'],
                p['easy_ship'],
                p['referral_fee'],
                p['ad_spend'],
                p['net_proceeds'],
                f"{p['acos_pct']}%" if p['acos_pct'] is not None else "—"
            ])
            tot_sold += p['units_sold']
            tot_ret  += p['units_returned']
            tot_net  += p['net_units']
            tot_ns   += p['net_sales']
            tot_es   += p['easy_ship']
            tot_rf   += p['referral_fee']
            tot_ads  += p['ad_spend']
            tot_np   += p['net_proceeds']

        overall_acos = round(tot_ads / tot_ns * 100, 1) if tot_ns > 0 and tot_ads > 0 else "—"
        w.writerow([
            "TOTAL",
            tot_sold, tot_ret, tot_net, "—",
            round(tot_ns, 2), round(tot_es, 2), round(tot_rf, 2),
            round(tot_ads, 2), round(tot_np, 2),
            f"{overall_acos}%" if isinstance(overall_acos, float) else overall_acos
        ])
        w.writerow([])
        w.writerow([])

        # ── Section 3: Key Insights ───────────────────
        w.writerow(["=" * 50])
        w.writerow(["SECTION 3: KEY INSIGHTS & RECOMMENDATIONS"])
        w.writerow(["=" * 50])
        w.writerow([])
        w.writerow(["Insight", "Detail"])
        for title, detail in insights:
            w.writerow([title, detail])

    return filepath


def main():
    print(f"\n{'='*55}")
    print("  Amazon Sales Report Analyzer")
    print(f"{'='*55}")
    print(f"  Input  : {INPUT_FILE}")

    df = load_data(INPUT_FILE, SHEET_NAME)
    print(f"  Loaded : {len(df)} rows, {len(df.columns)} columns")

    summary, products, insights = analyze(df)
    print(f"  Found  : {len(products)} active SKUs")

    out_path = write_csv(summary, products, insights, OUTPUT_DIR)
    print(f"  Output : {out_path}")
    print(f"{'='*55}\n")
    print("  QUICK SUMMARY:")
    print(f"  Gross Sales    : {summary['currency']} {summary['gross_sales']:,.2f}")
    print(f"  Net Sales      : {summary['currency']} {summary['net_sales']:,.2f}")
    print(f"  Total Fees     : {summary['currency']} {abs(summary['total_fees']):,.2f}")
    print(f"  Net Proceeds   : {summary['currency']} {summary['net_proceeds']:,.2f}")
    print(f"  Units Sold     : {summary['units_sold']}  |  Returned: {summary['units_returned']}")
    print(f"\n  {len(insights)} insight(s) generated.")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()
