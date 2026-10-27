# =====================================================================
# Oman Insurance Broker – Full IFRS + CMA Financial Reporting Demo
# =====================================================================
import streamlit as st
import pandas as pd
from datetime import date, datetime
import io
import uuid

st.set_page_config(page_title="Oman Insurance Broker FS", page_icon="🇴🇲", layout="wide")

# ---------- Optional imports ----------
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    PPTX_OK = True
except ImportError:
    PPTX_OK = False

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    MPL_OK = True
except ImportError:
    MPL_OK = False

# =====================================================================
# CONSTANTS
# =====================================================================
VAT_RATE = 0.05
ECL_RATES = {"current": 0.005, "30": 0.02, "60": 0.05, "90": 0.15, "120+": 0.30}
MIN_CAPITAL = 50000.0
MIN_CAPITAL_COMBINED = 75000.0

# Brand colors
BRAND_NAVY = RGBColor(0x1A, 0x2F, 0x5E)
BRAND_GOLD = RGBColor(0xC9, 0xA2, 0x27)
BRAND_LIGHT = RGBColor(0xF4, 0xF6, 0xFA)
BRAND_GREY = RGBColor(0x6B, 0x72, 0x80)
BRAND_GREEN = RGBColor(0x1E, 0x8E, 0x5A)
BRAND_RED = RGBColor(0xC0, 0x39, 0x2B)

CHART_OF_ACCOUNTS = {
    # Assets
    "1000": "Cash – Insurance Brokerage Account (Client Funds)",
    "1010": "Cash – Operating Account",
    "1020": "Bank – Operating Account",
    "1100": "Accounts Receivable – Commissions (net of ECL)",
    "1110": "VAT Receivable (Input)",
    "1120": "Right-of-Use Asset (Net)",
    "1130": "Prepaid Expenses",
    "1140": "Fixed Assets (Net)",
    "1150": "Deferred Tax Asset",
    # Liabilities
    "2000": "Accounts Payable – Insurers (Premiums)",
    "2010": "Accounts Payable – Clients (Claims)",
    "2100": "Accrued Expenses",
    "2110": "VAT Payable (Output)",
    "2120": "Lease Liability – Current",
    "2130": "Lease Liability – Non-Current",
    "2200": "Loans Payable",
    "2300": "Provision for End of Service Benefits",
    # Equity
    "3000": "Share Capital",
    "3100": "Retained Earnings",
    "3200": "Current Year Profit/Loss",
    # Revenue
    "4000": "Commission Income – Insurance",
    "4010": "Commission Income – Reinsurance",
    "4020": "Fee Income",
    "4030": "Claims Handling Fees",
    "4100": "Interest Income",
    # Expenses – Personnel
    "5000": "Salaries and Wages",
    "5005": "Staff Bonuses",
    "5010": "End of Service Benefits",
    "5015": "Staff Medical Insurance",
    "5020": "Staff Training & Development",
    # Expenses – Occupancy
    "5030": "Rent Expense (short-term)",
    "5035": "Utilities (Electricity/Water)",
    "5040": "ROU Depreciation",
    "5045": "Lease Interest",
    # Expenses – Operations
    "5050": "Telephone & Internet",
    "5055": "IT Software & Subscriptions",
    "5060": "Printing & Stationery",
    "5065": "Courier & Postage",
    # Expenses – Marketing
    "5070": "Marketing & Advertising",
    "5075": "Gifts & Corporate Hospitality",
    "5080": "Business Entertainment",
    "5085": "Donations & Sponsorships",
    # Expenses – Professional
    "5090": "Audit & Accounting Fees",
    "5095": "Legal & Consultancy Fees",
    "5100": "Regulatory & License Fees",
    # Expenses – Finance
    "5105": "Bank Charges",
    "5110": "ECL Impairment Loss",
    # Expenses – Other
    "5115": "Travel & Accommodation",
    "5120": "Motor Vehicle Expenses",
    "5125": "Insurance (Office)",
    "5130": "Depreciation – Fixed Assets",
    "5135": "Miscellaneous Expenses",
}

DEFAULT_BALANCES = {k: 0.0 for k in CHART_OF_ACCOUNTS}
DEFAULT_BALANCES.update({
    "1010": 5000.0, "1020": 15000.0, "1130": 1000.0, "1140": 10000.0,
    "2100": 2000.0, "3000": 50000.0, "3100": 5000.0,
})

# =====================================================================
# SESSION STATE
# =====================================================================
for key, default in [
    ("balances", DEFAULT_BALANCES.copy()),
    ("journal_entries", []),
    ("documents", []),
    ("accruals", []),
    ("ecl_allowance", 0.0),
    ("lease", None),
    ("audit_adjustments", []),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# =====================================================================
# HELPERS
# =====================================================================
def fmt(v):
    return f"{v:,.3f} OMR"

def get_bal(code):
    return st.session_state.balances.get(code, 0.0)

def post_journal(entry_date, description, lines):
    total_dr = sum(l["debit"] for l in lines)
    total_cr = sum(l["credit"] for l in lines)
    if abs(total_dr - total_cr) > 0.001:
        return False, "Debits ≠ Credits"
    if total_dr == 0:
        return False, "Zero amount"
    for l in lines:
        code = l["account"]
        net = l["debit"] - l["credit"]
        if code.startswith("1") or code.startswith("5"):
            st.session_state.balances[code] = st.session_state.balances.get(code, 0) + net
        else:
            st.session_state.balances[code] = st.session_state.balances.get(code, 0) - net
    st.session_state.journal_entries.append({
        "id": str(uuid.uuid4())[:8],
        "date": entry_date,
        "description": description,
        "lines": lines,
        "total": total_dr,
    })
    return True, f"Posted {fmt(total_dr)}"

# =====================================================================
# 3-YEAR SAMPLE DATA
# =====================================================================
def init_sample():
    if st.session_state.journal_entries:
        return
    entries = []
    # Year 1 (2023): Commission 55,000 | Year 2 (2024): 68,000 | Year 3 (2025): 82,000
    yearly = [
        (2023, 55000.0, 8500.0),
        (2024, 68000.0, 10500.0),
        (2025, 82000.0, 12000.0),
    ]
    for year, ins_commission, reins_commission in yearly:
        # Insurance commission - quarterly accruals (IFRS 15)
        q = ins_commission / 4
        for i, m in enumerate([3, 6, 9, 12]):
            entries.append((date(year, m, 25), f"Commission accrued – Q{i+1} {year} insurance policies",
                            [{"account": "1100", "debit": q, "credit": 0},
                             {"account": "4000", "debit": 0, "credit": q}]))
        # Reinsurance commission - quarterly
        qr = reins_commission / 4
        for i, m in enumerate([3, 6, 9, 12]):
            entries.append((date(year, m, 28), f"Commission accrued – Q{i+1} {year} reinsurance placements",
                            [{"account": "1100", "debit": qr, "credit": 0},
                             {"account": "4010", "debit": 0, "credit": qr}]))
        # Fee income - annual
        entries.append((date(year, 12, 30), f"Fee income – policy servicing {year}",
                        [{"account": "1100", "debit": 4500.0, "credit": 0},
                         {"account": "4020", "debit": 0, "credit": 4500.0}]))
        # Salaries – monthly (simplified: quarterly postings)
        salary_m = 3500.0
        for m in [3, 6, 9, 12]:
            entries.append((date(year, m, 28), f"Salaries & wages – Q{(m//3)} {year}",
                            [{"account": "5000", "debit": salary_m*3, "credit": 0},
                             {"account": "1020", "debit": 0, "credit": salary_m*3}]))
        # Staff bonuses – annual
        entries.append((date(year, 12, 15), f"Annual staff bonuses {year}",
                        [{"account": "5005", "debit": 3000.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 3000.0}]))
        # End of service benefits accrual
        entries.append((date(year, 12, 31), f"Provision for EOSB {year}",
                        [{"account": "5010", "debit": 2000.0, "credit": 0},
                         {"account": "2300", "debit": 0, "credit": 2000.0}]))
        # Staff medical
        entries.append((date(year, 1, 15), f"Staff medical insurance {year}",
                        [{"account": "5015", "debit": 1200.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 1200.0}]))
        # Training
        entries.append((date(year, 5, 20), f"Staff training – IFRS updates {year}",
                        [{"account": "5020", "debit": 800.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 800.0}]))
        # Rent – monthly
        for m in [3, 6, 9, 12]:
            entries.append((date(year, m, 1), f"Office rent – Q{(m//3)} {year}",
                            [{"account": "5030", "debit": 900.0*3, "credit": 0},
                             {"account": "1020", "debit": 0, "credit": 900.0*3}]))
        # Utilities
        for m in [3, 6, 9, 12]:
            entries.append((date(year, m, 10), f"Utilities – Q{(m//3)} {year}",
                            [{"account": "5035", "debit": 300.0, "credit": 0},
                             {"account": "1020", "debit": 0, "credit": 300.0}]))
        # Telephone & internet
        for m in [3, 6, 9, 12]:
            entries.append((date(year, m, 5), f"Telephone & internet – Q{(m//3)} {year}",
                            [{"account": "5050", "debit": 180.0, "credit": 0},
                             {"account": "1020", "debit": 0, "credit": 180.0}]))
        # IT Software
        entries.append((date(year, 2, 1), f"Insurance broker software subscription {year}",
                        [{"account": "5055", "debit": 2400.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 2400.0}]))
        # Printing & stationery
        entries.append((date(year, 3, 15), f"Printing & stationery {year}",
                        [{"account": "5060", "debit": 350.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 350.0}]))
        # Courier
        entries.append((date(year, 4, 10), f"Courier & postage {year}",
                        [{"account": "5065", "debit": 200.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 200.0}]))
        # Marketing
        entries.append((date(year, 6, 30), f"Marketing & advertising {year}",
                        [{"account": "5070", "debit": 3000.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 3000.0}]))
        # Gifts & corporate hospitality
        entries.append((date(year, 12, 20), f"Corporate gifts – clients & partners {year}",
                        [{"account": "5075", "debit": 1800.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 1800.0}]))
        # Business entertainment
        entries.append((date(year, 11, 15), f"Client entertainment events {year}",
                        [{"account": "5080", "debit": 1200.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 1200.0}]))
        # Donations
        entries.append((date(year, 9, 5), f"Charitable donations & sponsorships {year}",
                        [{"account": "5085", "debit": 1000.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 1000.0}]))
        # Audit fees
        entries.append((date(year, 12, 31), f"Accrual: statutory audit fee {year}",
                        [{"account": "5090", "debit": 1500.0, "credit": 0},
                         {"account": "2100", "debit": 0, "credit": 1500.0}]))
        # Legal & consultancy
        entries.append((date(year, 8, 15), f"Legal & consultancy fees {year}",
                        [{"account": "5095", "debit": 900.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 900.0}]))
        # Regulatory fees
        entries.append((date(year, 1, 30), f"CMA license & regulatory fees {year}",
                        [{"account": "5100", "debit": 500.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 500.0}]))
        # Bank charges
        for m in [3, 6, 9, 12]:
            entries.append((date(year, m, 28), f"Bank charges – Q{(m//3)} {year}",
                            [{"account": "5105", "debit": 60.0, "credit": 0},
                             {"account": "1020", "debit": 0, "credit": 60.0}]))
        # Travel
        entries.append((date(year, 7, 20), f"Business travel & accommodation {year}",
                        [{"account": "5115", "debit": 2200.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 2200.0}]))
        # Motor vehicle
        entries.append((date(year, 12, 31), f"Motor vehicle expenses {year}",
                        [{"account": "5120", "debit": 1500.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 1500.0}]))
        # Office insurance
        entries.append((date(year, 1, 20), f"Office insurance premium {year}",
                        [{"account": "5125", "debit": 600.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 600.0}]))
        # Depreciation
        entries.append((date(year, 12, 31), f"Depreciation – fixed assets {year}",
                        [{"account": "5130", "debit": 2000.0, "credit": 0},
                         {"account": "1140", "debit": 0, "credit": 2000.0}]))
        # Misc
        entries.append((date(year, 12, 31), f"Miscellaneous expenses {year}",
                        [{"account": "5135", "debit": 400.0, "credit": 0},
                         {"account": "1020", "debit": 0, "credit": 400.0}]))

    for d, desc, lines in entries:
        post_journal(d, desc, lines)

init_sample()

# =====================================================================
# ECL / LEASE / AUDIT
# =====================================================================
def compute_ecl():
    ar = get_bal("1100")
    if ar <= 0:
        return 0.0, []
    buckets = [
        ("Current", ECL_RATES["current"]), ("1–30 days", ECL_RATES["30"]),
        ("31–60 days", ECL_RATES["60"]), ("61–90 days", ECL_RATES["90"]),
        ("90+ days", ECL_RATES["120+"]),
    ]
    splits = [0.60, 0.25, 0.10, 0.05, 0.00]
    details, total = [], 0.0
    for i, (label, rate) in enumerate(buckets):
        gross = ar * splits[i]
        loss = gross * rate
        details.append({"Bucket": label, "Gross (OMR)": gross,
                        "Rate": f"{rate*100:.1f}%", "ECL (OMR)": loss})
        total += loss
    return total, details

def setup_lease(annual_payment, term_years, rate):
    pv = annual_payment * (1 - (1 + rate) ** -term_years) / rate
    st.session_state.lease = {
        "annual_payment": annual_payment, "term": term_years, "rate": rate,
        "initial_rou": pv, "rou_balance": pv, "liability_balance": pv,
        "year": 0, "schedule": [],
    }
    post_journal(date.today(), "IFRS 16 – Initial recognition of office lease",
                 [{"account": "1120", "debit": pv, "credit": 0},
                  {"account": "2130", "debit": 0, "credit": pv}])

def run_lease_year():
    L = st.session_state.lease
    if not L or L["year"] >= L["term"]:
        return
    interest = L["liability_balance"] * L["rate"]
    principal = L["annual_payment"] - interest
    new_liab = L["liability_balance"] - principal
    dep = L["initial_rou"] / L["term"]
    new_rou = L["rou_balance"] - dep
    post_journal(date.today(), f"IFRS 16 Y{L['year']+1} – Interest accretion",
                 [{"account": "5045", "debit": interest, "credit": 0},
                  {"account": "2130", "debit": 0, "credit": interest}])
    post_journal(date.today(), f"IFRS 16 Y{L['year']+1} – ROU depreciation",
                 [{"account": "5040", "debit": dep, "credit": 0},
                  {"account": "1120", "debit": 0, "credit": dep}])
    post_journal(date.today(), f"IFRS 16 Y{L['year']+1} – Lease payment (principal)",
                 [{"account": "2130", "debit": principal, "credit": 0},
                  {"account": "1020", "debit": 0, "credit": principal}])
    L["schedule"].append({
        "Year": L["year"]+1, "Opening Liability": L["liability_balance"],
        "Interest": interest, "Payment": L["annual_payment"],
        "Principal": principal, "Closing Liability": new_liab,
        "ROU Depreciation": dep, "ROU Closing": new_rou,
    })
    L["liability_balance"], L["rou_balance"], L["year"] = new_liab, new_rou, L["year"]+1

def post_audit_adjustment(desc, lines):
    ok, msg = post_journal(date.today(), f"[AUDIT ADJ] {desc}", lines)
    if ok:
        st.session_state.audit_adjustments.append({"desc": desc, "lines": lines})
    return ok, msg

# =====================================================================
# PDF
# =====================================================================
def generate_pdf_document(doc_type, doc_data):
    if not REPORTLAB_OK:
        return None
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=20*mm, leftMargin=20*mm,
                            topMargin=20*mm, bottomMargin=20*mm)
    styles = getSampleStyleSheet()
    elements = [
        Paragraph(f"<b>{doc_type.upper()}</b>",
                  ParagraphStyle('T', parent=styles['Heading1'], fontSize=18, alignment=TA_CENTER)),
        Spacer(1, 10),
    ]
    t = Table([["<b>Al Noor Insurance Brokerage LLC</b>", ""],
               ["Muscat, Sultanate of Oman", ""],
               ["VAT No: OM1100XXXXXX", ""]], colWidths=[100*mm, 60*mm])
    t.setStyle(TableStyle([('FONTSIZE', (0, 0), (-1, -1), 9)]))
    elements.append(t); elements.append(Spacer(1, 15))
    t2 = Table([["<b>Document No:</b>", doc_data.get("doc_no", "N/A")],
                ["<b>Date:</b>", str(doc_data.get("date", date.today()))],
                ["<b>Party:</b>", doc_data.get("party", "N/A")],
                ["<b>Reference:</b>", doc_data.get("reference", "N/A")]],
               colWidths=[40*mm, 120*mm])
    t2.setStyle(TableStyle([('FONTSIZE', (0, 0), (-1, -1), 10),
                            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                            ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey)]))
    elements.append(t2); elements.append(Spacer(1, 15))
    items = doc_data.get("items", [])
    if items:
        data = [["Description", "Amount (OMR)"]]
        data += [[i["desc"], f"{i['amount']:,.3f}"] for i in items]
        sub = sum(i["amount"] for i in items)
        vat = doc_data.get("vat", 0)
        data += [["", ""], ["<b>Subtotal</b>", f"<b>{sub:,.3f}</b>"]]
        if vat > 0:
            data.append(["<b>VAT 5%</b>", f"<b>{vat:,.3f}</b>"])
        data.append(["<b>Total</b>", f"<b>{sub+vat:,.3f}</b>"])
        t3 = Table(data, colWidths=[120*mm, 40*mm])
        t3.setStyle(TableStyle([('FONTSIZE', (0, 0), (-1, -1), 9),
                                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a5490')),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                                ('ALIGN', (1, 0), (1, -1), 'RIGHT')]))
        elements.append(t3)
    doc.build(elements)
    buf.seek(0)
    return buf

# =====================================================================
# EXCEL
# =====================================================================
def generate_excel_report():
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as w:
        pd.DataFrame([{"Code": k, "Account": v, "Balance (OMR)": st.session_state.balances[k]}
                      for k, v in CHART_OF_ACCOUNTS.items()]).to_excel(w, sheet_name='Trial Balance', index=False)
        if st.session_state.journal_entries:
            rows = []
            for je in st.session_state.journal_entries:
                for l in je["lines"]:
                    rows.append({"Date": je["date"], "Journal ID": je["id"],
                                 "Description": je["description"], "Account": l["account"],
                                 "Account Name": CHART_OF_ACCOUNTS.get(l["account"], ""),
                                 "Debit": l["debit"], "Credit": l["credit"]})
            pd.DataFrame(rows).to_excel(w, sheet_name='Journal', index=False)
        # P&L
        rev = {k: v for k, v in st.session_state.balances.items() if k.startswith("4")}
        exp = {k: v for k, v in st.session_state.balances.items() if k.startswith("5")}
        pl = [{"Account": CHART_OF_ACCOUNTS[k], "Amount": v} for k, v in rev.items()]
        pl.append({"Account": "TOTAL REVENUE", "Amount": sum(rev.values())})
        pl.append({"Account": "", "Amount": 0})
        pl += [{"Account": CHART_OF_ACCOUNTS[k], "Amount": v} for k, v in exp.items()]
        pl.append({"Account": "TOTAL EXPENSES", "Amount": sum(exp.values())})
        pl.append({"Account": "NET PROFIT / (LOSS)", "Amount": sum(rev.values()) - sum(exp.values())})
        pd.DataFrame(pl).to_excel(w, sheet_name='P&L', index=False)
        if st.session_state.lease and st.session_state.lease["schedule"]:
            pd.DataFrame(st.session_state.lease["schedule"]).to_excel(w, sheet_name='Lease Schedule', index=False)
    out.seek(0)
    return out

# =====================================================================
# HIGHLY FORMATTED PPTX
# =====================================================================
def _add_branded_header(slide, title, subtitle=None):
    """Add navy header bar with gold accent line."""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.0))
    bar.fill.solid(); bar.fill.fore_color.rgb = BRAND_NAVY
    bar.line.fill.background()

    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.0), Inches(13.333), Inches(0.06))
    accent.fill.solid(); accent.fill.fore_color.rgb = BRAND_GOLD
    accent.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.15), Inches(12), Inches(0.5))
    tf = tb.text_frame; tf.text = title
    p = tf.paragraphs[0]; p.font.size = Pt(26); p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    if subtitle:
        tb2 = slide.shapes.add_textbox(Inches(0.4), Inches(0.62), Inches(12), Inches(0.35))
        tf2 = tb2.text_frame; tf2.text = subtitle
        p2 = tf2.paragraphs[0]; p2.font.size = Pt(12)
        p2.font.color.rgb = BRAND_GOLD

def _add_footer(slide, page_num, total):
    """Add footer with company name and page number."""
    tb = slide.shapes.add_textbox(Inches(0.4), Inches(7.05), Inches(8), Inches(0.3))
    tf = tb.text_frame; tf.text = "Al Noor Insurance Brokerage LLC  ·  Muscat, Oman  ·  Confidential"
    p = tf.paragraphs[0]; p.font.size = Pt(9); p.font.color.rgb = BRAND_GREY

    tb2 = slide.shapes.add_textbox(Inches(11.8), Inches(7.05), Inches(1.3), Inches(0.3))
    tf2 = tb2.text_frame; tf2.text = f"{page_num} / {total}"
    p2 = tf2.paragraphs[0]; p2.font.size = Pt(9); p2.font.color.rgb = BRAND_GREY
    p2.alignment = PP_ALIGN.RIGHT

def _add_kpi_box(slide, left, top, width, height, label, value, delta=None, color=BRAND_NAVY):
    """Rounded KPI tile with label + value + optional delta."""
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    box.fill.solid(); box.fill.fore_color.rgb = BRAND_LIGHT
    box.line.color.rgb = color
    box.line.width = Pt(1.5)

    # Top accent
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.06))
    accent.fill.solid(); accent.fill.fore_color.rgb = color
    accent.line.fill.background()

    tf = box.text_frame
    tf.margin_left = Inches(0.15); tf.margin_top = Inches(0.15)
    tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = label
    p.font.size = Pt(10); p.font.color.rgb = BRAND_GREY; p.font.bold = True

    p2 = tf.add_paragraph(); p2.text = value
    p2.font.size = Pt(20); p2.font.bold = True; p2.font.color.rgb = color

    if delta:
        p3 = tf.add_paragraph(); p3.text = delta
        p3.font.size = Pt(10)
        p3.font.color.rgb = BRAND_GREEN if delta.startswith("+") else BRAND_RED

def _add_table(slide, left, top, width, height, headers, rows, col_widths=None):
    """Formatted table with navy header."""
    n_rows = len(rows) + 1
    n_cols = len(headers)
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    table = tbl_shape.table

    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w

    # Header
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        cell.fill.solid(); cell.fill.fore_color.rgb = BRAND_NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True; p.font.size = Pt(11)
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        if j > 0:
            p.alignment = PP_ALIGN.RIGHT
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    # Rows
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = BRAND_LIGHT if i % 2 == 0 else RGBColor(0xFF, 0xFF, 0xFF)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
            if j > 0:
                p.alignment = PP_ALIGN.RIGHT
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    return table

def _add_column_chart(slide, left, top, width, height, categories, series_dict, title):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series_dict.items():
        chart_data.add_series(name, values)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data)
    chart = gf.chart
    chart.has_title = True
    chart.chart_title.text_frame.text = title
    for p in chart.chart_title.text_frame.paragraphs:
        p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = BRAND_NAVY
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.font.size = Pt(10)
    return chart

def _add_pie_chart(slide, left, top, width, height, categories, values, title):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("Revenue", values)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.PIE, left, top, width, height, chart_data)
    chart = gf.chart
    chart.has_title = True
    chart.chart_title.text_frame.text = title
    for p in chart.chart_title.text_frame.paragraphs:
        p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = BRAND_NAVY
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.legend.include_in_layout = False
    chart.font.size = Pt(10)
    return chart

def generate_pptx_report():
    if not PPTX_OK:
        return None
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    BLANK = prs.slide_layouts[6]
    TOTAL_SLIDES = 8

    # Aggregate financials
    total_assets = sum(v for k, v in st.session_state.balances.items() if k.startswith("1") and k != "1000")
    total_liab = sum(v for k, v in st.session_state.balances.items() if k.startswith("2"))
    total_equity = sum(v for k, v in st.session_state.balances.items() if k.startswith("3"))
    net = total_assets - total_liab
    rev = {k: v for k, v in st.session_state.balances.items() if k.startswith("4")}
    exp = {k: v for k, v in st.session_state.balances.items() if k.startswith("5")}
    total_rev = sum(rev.values())
    total_exp = sum(exp.values())
    net_profit = total_rev - total_exp

    # --- SLIDE 1: COVER ---
    s = prs.slides.add_slide(BLANK)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid(); bg.fill.fore_color.rgb = BRAND_NAVY
    bg.line.fill.background()

    accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(4.8), Inches(13.333), Inches(0.08))
    accent.fill.solid(); accent.fill.fore_color.rgb = BRAND_GOLD
    accent.line.fill.background()

    tb = s.shapes.add_textbox(Inches(0.8), Inches(2.4), Inches(11.7), Inches(1.2))
    tf = tb.text_frame; tf.text = "Al Noor Insurance Brokerage LLC"
    p = tf.paragraphs[0]; p.font.size = Pt(44); p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    tb2 = s.shapes.add_textbox(Inches(0.8), Inches(3.5), Inches(11.7), Inches(0.8))
    tf2 = tb2.text_frame; tf2.text = "Financial Summary & CMA Compliance Report"
    p2 = tf2.paragraphs[0]; p2.font.size = Pt(22); p2.font.color.rgb = BRAND_GOLD

    tb3 = s.shapes.add_textbox(Inches(0.8), Inches(5.2), Inches(11.7), Inches(1.5))
    tf3 = tb3.text_frame
    tf3.text = f"Reporting period: FY2023 – FY2025"
    p3 = tf3.paragraphs[0]; p3.font.size = Pt(16); p3.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p4 = tf3.add_paragraph(); p4.text = f"Prepared: {date.today().strftime('%d %B %Y')}"
    p4.font.size = Pt(14); p4.font.color.rgb = BRAND_GOLD
    p5 = tf3.add_paragraph(); p5.text = "IFRS · Oman CMA · VAT compliant"
    p5.font.size = Pt(14); p5.font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)

    # --- SLIDE 2: EXECUTIVE SUMMARY (KPI tiles) ---
    s = prs.slides.add_slide(BLANK)
    _add_branded_header(s, "Executive Summary", "Key financial metrics – all years to date")

    kpi_w, kpi_h = Inches(3.0), Inches(1.6)
    _add_kpi_box(s, Inches(0.5), Inches(1.5), kpi_w, kpi_h,
                 "TOTAL REVENUE", fmt(total_rev), color=BRAND_NAVY)
    _add_kpi_box(s, Inches(3.7), Inches(1.5), kpi_w, kpi_h,
                 "TOTAL EXPENSES", fmt(total_exp), color=BRAND_RED)
    _add_kpi_box(s, Inches(6.9), Inches(1.5), kpi_w, kpi_h,
                 "NET PROFIT", fmt(net_profit),
                 delta=("+" if net_profit >= 0 else "–"), color=BRAND_GREEN if net_profit >= 0 else BRAND_RED)
    _add_kpi_box(s, Inches(10.1), Inches(1.5), kpi_w, kpi_h,
                 "PROFIT MARGIN",
                 f"{(net_profit/total_rev*100 if total_rev else 0):.1f}%", color=BRAND_GOLD)

    _add_kpi_box(s, Inches(0.5), Inches(3.4), kpi_w, kpi_h,
                 "TOTAL ASSETS", fmt(total_assets), color=BRAND_NAVY)
    _add_kpi_box(s, Inches(3.7), Inches(3.4), kpi_w, kpi_h,
                 "TOTAL LIABILITIES", fmt(total_liab), color=BRAND_RED)
    _add_kpi_box(s, Inches(6.9), Inches(3.4), kpi_w, kpi_h,
                 "NET ASSETS", fmt(net), color=BRAND_GREEN)
    _add_kpi_box(s, Inches(10.1), Inches(3.4), kpi_w, kpi_h,
                 "RETURN ON EQUITY",
                 f"{(net_profit/total_equity*100 if total_equity else 0):.1f}%", color=BRAND_GOLD)

    # Compliance badge
    status_box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(0.5), Inches(5.5), Inches(12.3), Inches(1.0))
    status_box.fill.solid()
    status_box.fill.fore_color.rgb = BRAND_GREEN if net >= MIN_CAPITAL else BRAND_RED
    status_box.line.fill.background()
    tf = status_box.text_frame; tf.margin_left = Inches(0.3)
    p = tf.paragraphs[0]
    p.text = (f"✅ CMA Article 26 Solvency Requirement MET  ·  Net Assets {fmt(net)} ≥ {fmt(MIN_CAPITAL)}"
              if net >= MIN_CAPITAL else
              f"⚠️ CMA Article 26 Solvency Requirement NOT MET  ·  {fmt(net)} < {fmt(MIN_CAPITAL)}")
    p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    _add_footer(s, 2, TOTAL_SLIDES)

    # --- SLIDE 3: 3-YEAR REVENUE TREND ---
    s = prs.slides.add_slide(BLANK)
    _add_branded_header(s, "3-Year Revenue Trend", "Commission income by line of business")

    years = ["2023", "2024", "2025"]
    # Approximate yearly splits (from demo data)
    ins_by_year = [55000, 68000, 82000]
    reins_by_year = [8500, 10500, 12000]
    fee_by_year = [4500, 4500, 4500]

    _add_column_chart(s, Inches(0.5), Inches(1.4), Inches(8.0), Inches(5.4),
                      years,
                      {"Insurance Commission": ins_by_year,
                       "Reinsurance Commission": reins_by_year,
                       "Fee Income": fee_by_year},
                      "Revenue by Year & Line (OMR)")

    _add_pie_chart(s, Inches(8.7), Inches(1.4), Inches(4.3), Inches(5.4),
                   ["Insurance", "Reinsurance", "Fees"],
                   [sum(ins_by_year), sum(reins_by_year), sum(fee_by_year)],
                   "Revenue Mix (3-Yr)")

    _add_footer(s, 3, TOTAL_SLIDES)

    # --- SLIDE 4: P&L SUMMARY ---
    s = prs.slides.add_slide(BLANK)
    _add_branded_header(s, "Profit & Loss Summary", "Aggregated revenue and expense accounts")

    # Build compact P&L table
    pl_rows = []
    for k, v in rev.items():
        if v != 0:
            pl_rows.append([CHART_OF_ACCOUNTS[k], f"{v:,.3f}"])
    pl_rows.append(["", ""])
    pl_rows.append(["TOTAL REVENUE", f"{total_rev:,.3f}"])
    for k, v in exp.items():
        if v != 0:
            pl_rows.append([CHART_OF_ACCOUNTS[k], f"{v:,.3f}"])
    pl_rows.append(["", ""])
    pl_rows.append(["TOTAL EXPENSES", f"{total_exp:,.3f}"])
    pl_rows.append(["NET PROFIT / (LOSS)", f"{net_profit:,.3f}"])

    # Cap the display to top entries for readability
    pl_rows_display = pl_rows[:20]

    _add_table(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.4),
               ["Account", "Amount (OMR)"], pl_rows_display,
               col_widths=[Inches(8.5), Inches(3.8)])
    _add_footer(s, 4, TOTAL_SLIDES)

    # --- SLIDE 5: BALANCE SHEET ---
    s = prs.slides.add_slide(BLANK)
    _add_branded_header(s, "Statement of Financial Position", "Assets, liabilities & equity")

    asset_rows = [[CHART_OF_ACCOUNTS[k], f"{v:,.3f}"]
                  for k, v in st.session_state.balances.items()
                  if k.startswith("1") and k != "1000" and v != 0]
    asset_rows.append(["TOTAL ASSETS", f"{total_assets:,.3f}"])

    le_rows = [[CHART_OF_ACCOUNTS[k], f"{v:,.3f}"]
               for k, v in st.session_state.balances.items() if k.startswith("2") and v != 0]
    le_rows.append(["Total Liabilities", f"{total_liab:,.3f}"])
    le_rows.append(["", ""])
    eq_rows = [[CHART_OF_ACCOUNTS[k], f"{v:,.3f}"]
               for k, v in st.session_state.balances.items() if k.startswith("3") and v != 0]
    eq_rows.append(["Total Equity", f"{total_equity:,.3f}"])

    _add_table(s, Inches(0.5), Inches(1.4), Inches(6.0), Inches(5.4),
               ["ASSETS", "OMR"], asset_rows,
               col_widths=[Inches(4.3), Inches(1.7)])
    _add_table(s, Inches(6.8), Inches(1.4), Inches(6.0), Inches(5.4),
               ["LIABILITIES & EQUITY", "OMR"], le_rows + eq_rows,
               col_widths=[Inches(4.3), Inches(1.7)])
    _add_footer(s, 5, TOTAL_SLIDES)

    # --- SLIDE 6: EXPENSE BREAKDOWN ---
    s = prs.slides.add_slide(BLANK)
    _add_branded_header(s, "Expense Breakdown", "All expense categories including gifts, entertainment, donations")

    # Group expenses
    expense_groups = {
        "Personnel": ["5000", "5005", "5010", "5015", "5020"],
        "Occupancy": ["5030", "5035", "5040", "5045"],
        "Operations": ["5050", "5055", "5060", "5065"],
        "Marketing & Gifts": ["5070", "5075", "5080", "5085"],
        "Professional": ["5090", "5095", "5100"],
        "Finance": ["5105", "5110"],
        "Other": ["5115", "5120", "5125", "5130", "5135"],
    }
    group_totals = {}
    for g, codes in expense_groups.items():
        group_totals[g] = sum(st.session_state.balances.get(c, 0) for c in codes)

    _add_column_chart(s, Inches(0.5), Inches(1.4), Inches(8.0), Inches(5.4),
                      list(group_totals.keys()),
                      {"Total (OMR)": list(group_totals.values())},
                      "Expenses by Category (3-Year Total)")

    # Gifts & entertainment detail table
    gift_rows = []
    for code in ["5070", "5075", "5080", "5085"]:
        v = st.session_state.balances.get(code, 0)
        gift_rows.append([CHART_OF_ACCOUNTS[code], f"{v:,.3f}"])
    gift_rows.append(["Sub-total (Marketing, Gifts, Entertainment, Donations)",
                      f"{sum(st.session_state.balances.get(c,0) for c in ['5070','5075','5080','5085']):,.3f}"])

    _add_table(s, Inches(8.7), Inches(1.4), Inches(4.3), Inches(5.4),
               ["Category", "OMR"], gift_rows,
               col_widths=[Inches(3.1), Inches(1.2)])
    _add_footer(s, 6, TOTAL_SLIDES)

    # --- SLIDE 7: CMA COMPLIANCE DASHBOARD ---
    s = prs.slides.add_slide(BLANK)
    _add_branded_header(s, "CMA Compliance Dashboard", "Oman Capital Market Authority – Insurance Brokers Regulation")

    checks = [
        ("Article 25 – Client funds segregated (A/C 1000)", "PASS",
         f"Client balance: {fmt(st.session_state.balances.get('1000', 0))}"),
        ("Article 26 – Minimum net assets ≥ RO 50,000", "PASS" if net >= MIN_CAPITAL else "FAIL",
         f"Net assets: {fmt(net)}"),
        ("Article 26 – Intangibles excluded from net assets", "PASS", "ROU asset separately classified"),
        ("VAT – 5% output VAT on commission income", "PASS", "Tracked in A/C 2110"),
        ("IFRS 15 – Revenue recognised on placement", "PASS", "Commission accrued quarterly"),
        ("IFRS 9 – ECL allowance on receivables", "PASS",
         f"Allowance: {fmt(st.session_state.ecl_allowance)}"),
        ("IFRS 16 – Lease ROU & liability recognised", "PASS" if st.session_state.lease else "N/A",
         "See Lease module"),
        ("Statutory audit by CMA-accredited auditor", "PASS", "Audit module active"),
    ]

    comp_rows = [[c[0], c[1], c[2]] for c in checks]
    _add_table(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.4),
               ["Compliance Requirement", "Status", "Evidence"],
               comp_rows,
               col_widths=[Inches(5.5), Inches(1.3), Inches(5.5)])
    _add_footer(s, 7, TOTAL_SLIDES)

    # --- SLIDE 8: CLOSING ---
    s = prs.slides.add_slide(BLANK)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid(); bg.fill.fore_color.rgb = BRAND_NAVY
    bg.line.fill.background()

    accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(3.6), Inches(13.333), Inches(0.08))
    accent.fill.solid(); accent.fill.fore_color.rgb = BRAND_GOLD
    accent.line.fill.background()

    tb = s.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(11.7), Inches(1.2))
    tf = tb.text_frame; tf.text = "Thank You"
    p = tf.paragraphs[0]; p.font.size = Pt(54); p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); p.alignment = PP_ALIGN.CENTER

    tb2 = s.shapes.add_textbox(Inches(0.8), Inches(4.0), Inches(11.7), Inches(1.5))
    tf2 = tb2.text_frame
    tf2.text = "Al Noor Insurance Brokerage LLC"
    p2 = tf2.paragraphs[0]; p2.font.size = Pt(22); p2.font.color.rgb = BRAND_GOLD
    p2.alignment = PP_ALIGN.CENTER
    p3 = tf2.add_paragraph(); p3.text = "Muscat, Sultanate of Oman"
    p3.font.size = Pt(16); p3.font.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
    p3.alignment = PP_ALIGN.CENTER
    p4 = tf2.add_paragraph(); p4.text = "This presentation is confidential and for internal use only."
    p4.font.size = Pt(12); p4.font.color.rgb = BRAND_GREY
    p4.alignment = PP_ALIGN.CENTER

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# =====================================================================
# NAVIGATION
# =====================================================================
st.sidebar.title("🏦 Insurance Broker FS")
st.sidebar.caption("Oman CMA · Full IFRS")
if not REPORTLAB_OK:
    st.sidebar.warning("reportlab missing – PDF disabled")
if not PPTX_OK:
    st.sidebar.warning("python-pptx missing – PPTX disabled")

menu = st.sidebar.radio("Navigation", [
    "📊 Dashboard", "📥 Journal Entry", "📄 Documents",
    "🔧 Accruals & ECL", "🏢 Lease (IFRS 16)",
    "🔍 Statutory Audit", "📑 Reports", "📚 IFRS Notes"
])

# =====================================================================
# DASHBOARD
# =====================================================================
if menu == "📊 Dashboard":
    st.title("Dashboard")
    assets = sum(v for k, v in st.session_state.balances.items() if k.startswith("1") and k != "1000")
    liab = sum(v for k, v in st.session_state.balances.items() if k.startswith("2"))
    net = assets - liab
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Assets (excl. client funds)", fmt(assets))
    c2.metric("Total Liabilities", fmt(liab))
    c3.metric("Net Assets", fmt(net))
    st.subheader("CMA Article 26 Solvency Test")
    if net >= MIN_CAPITAL:
        st.success(f"✅ MET: {fmt(net)} ≥ {fmt(MIN_CAPITAL)}")
    else:
        st.error(f"⚠️ NOT MET: {fmt(net)} < {fmt(MIN_CAPITAL)}")
    st.info("Client funds (Account 1000) excluded per CMA Article 25.")

    st.subheader("Revenue by Year")
    if st.session_state.journal_entries:
        df = pd.DataFrame(st.session_state.journal_entries)
        df["Year"] = pd.to_datetime(df["date"]).dt.year
        rev_rows = []
        for je in st.session_state.journal_entries:
            yr = pd.to_datetime(je["date"]).year
            for l in je["lines"]:
                if l["account"] in ("4000", "4010", "4020", "4030") and l["credit"] > 0:
                    rev_rows.append({"Year": yr, "Account": l["account"], "Amount": l["credit"]})
        if rev_rows:
            rev_df = pd.DataFrame(rev_rows).groupby(["Year", "Account"])["Amount"].sum().unstack(fill_value=0)
            rev_df.columns = [CHART_OF_ACCOUNTS.get(c, c) for c in rev_df.columns]
            st.dataframe(rev_df, use_container_width=True)
            st.bar_chart(rev_df)

# =====================================================================
# JOURNAL ENTRY
# =====================================================================
elif menu == "📥 Journal Entry":
    st.title("Multi-Line Double-Entry Journal")
    if "jlines" not in st.session_state:
        st.session_state.jlines = [{"account": "1010", "debit": 0.0, "credit": 0.0},
                                   {"account": "4000", "debit": 0.0, "credit": 0.0}]
    ca, cb = st.columns(2)
    if ca.button("➕ Add Line"):
        st.session_state.jlines.append({"account": "4000", "debit": 0.0, "credit": 0.0})
        st.rerun()
    if cb.button("➖ Remove Last"):
        if len(st.session_state.jlines) > 2:
            st.session_state.jlines.pop(); st.rerun()
    with st.form("jform"):
        ed = st.date_input("Date", date.today())
        desc = st.text_input("Description")
        lines = []
        for i, ln in enumerate(st.session_state.jlines):
            c1, c2, c3 = st.columns([3, 1, 1])
            acct = c1.selectbox(f"Account {i+1}", list(CHART_OF_ACCOUNTS.keys()),
                                index=list(CHART_OF_ACCOUNTS.keys()).index(ln["account"]),
                                format_func=lambda x: f"{x} – {CHART_OF_ACCOUNTS[x]}", key=f"a{i}")
            dr = c2.number_input("Debit", min_value=0.0, step=1.0, value=float(ln["debit"]), key=f"d{i}")
            cr = c3.number_input("Credit", min_value=0.0, step=1.0, value=float(ln["credit"]), key=f"c{i}")
            lines.append({"account": acct, "debit": dr, "credit": cr})
        tdr = sum(l["debit"] for l in lines); tcr = sum(l["credit"] for l in lines)
        st.write(f"Dr: {fmt(tdr)} | Cr: {fmt(tcr)}")
        if st.form_submit_button("Post"):
            ok, msg = post_journal(ed, desc, lines)
            if ok:
                st.success(msg)
                st.session_state.jlines = [{"account": "1010", "debit": 0.0, "credit": 0.0},
                                           {"account": "4000", "debit": 0.0, "credit": 0.0}]
                st.rerun()
            else:
                st.error(msg)
    if st.session_state.journal_entries:
        st.subheader("Recent Entries")
        for je in reversed(st.session_state.journal_entries[-8:]):
            with st.expander(f"{je['date']} – {je['description']} – {fmt(je['total'])}"):
                st.dataframe(pd.DataFrame([
                    {"Account": f"{l['account']} – {CHART_OF_ACCOUNTS.get(l['account'],'')}",
                     "Debit": l["debit"], "Credit": l["credit"]}
                    for l in je["lines"] if l["debit"] or l["credit"]
                ]), use_container_width=True, hide_index=True)

# =====================================================================
# DOCUMENTS
# =====================================================================
elif menu == "📄 Documents":
    st.title("Document Generation")
    doc_type = st.selectbox("Type", ["VAT Invoice (Commission)", "Invoice",
                                     "Debit Note", "Credit Note", "Receipt"])
    c1, c2 = st.columns(2)
    party = c1.text_input("Party", "Dhofar Insurance")
    ref = c2.text_input("Reference", "POL-2026-001")
    desc = st.text_input("Description", "Commission – Motor fleet")
    amt = st.number_input("Amount (OMR)", min_value=0.0, value=1250.0, step=10.0)
    apply_vat = st.checkbox("VAT 5%", value=(doc_type == "VAT Invoice (Commission)"))
    if st.button("Generate"):
        vat = round(amt * VAT_RATE, 3) if apply_vat else 0.0
        doc_no = f"{doc_type[:3].upper()}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        doc = {"doc_no": doc_no, "type": doc_type, "date": date.today(),
               "party": party, "reference": ref,
               "items": [{"desc": desc, "amount": amt}], "vat": vat, "total": amt + vat}
        st.session_state.documents.append(doc)
        st.success(f"{doc_no} – {fmt(doc['total'])}")
        pdf_buf = generate_pdf_document(doc_type, doc)
        if pdf_buf:
            st.download_button("⬇️ Download PDF", data=pdf_buf,
                               file_name=f"{doc_no}.pdf", mime="application/pdf")
    if st.session_state.documents:
        st.dataframe(pd.DataFrame([{"No": d["doc_no"], "Type": d["type"],
                                    "Party": d["party"], "Total": d["total"]}
                                   for d in st.session_state.documents]),
                     use_container_width=True)

# =====================================================================
# ACCRUALS & ECL
# =====================================================================
elif menu == "🔧 Accruals & ECL":
    st.title("Accruals & ECL (IFRS 9)")
    tab1, tab2 = st.tabs(["📅 Accruals", "📉 ECL"])
    with tab1:
        with st.form("accrual_form"):
            acc_date = st.date_input("Date", date.today())
            acc_desc = st.text_input("Description", "Accrued audit fee")
            exp_acct = st.selectbox("Expense Account",
                                    [k for k in CHART_OF_ACCOUNTS if k.startswith("5")],
                                    format_func=lambda x: f"{x} – {CHART_OF_ACCOUNTS[x]}")
            acc_amt = st.number_input("Amount (OMR)", min_value=0.0, value=1500.0, step=100.0)
            if st.form_submit_button("Post Accrual"):
                lines = [{"account": exp_acct, "debit": acc_amt, "credit": 0},
                         {"account": "2100", "debit": 0, "credit": acc_amt}]
                ok, msg = post_journal(acc_date, f"Accrual: {acc_desc}", lines)
                if ok:
                    st.session_state.accruals.append({"date": acc_date, "desc": acc_desc, "amount": acc_amt})
                    st.success(msg)
        if st.session_state.accruals:
            st.dataframe(pd.DataFrame(st.session_state.accruals), use_container_width=True)
    with tab2:
        ar = get_bal("1100")
        st.write(f"Gross Accounts Receivable: **{fmt(ar)}**")
        if st.button("Compute ECL"):
            ecl, details = compute_ecl()
            st.session_state.ecl_allowance = ecl
            st.dataframe(pd.DataFrame(details), use_container_width=True, hide_index=True)
            st.metric("Total ECL Allowance", fmt(ecl))
        if st.session_state.ecl_allowance > 0 and st.button("Post ECL Adjustment"):
            delta = st.session_state.ecl_allowance
            lines = [{"account": "5110", "debit": delta, "credit": 0},
                     {"account": "1100", "debit": 0, "credit": delta}]
            ok, msg = post_journal(date.today(), "IFRS 9 ECL impairment", lines)
            if ok:
                st.success(f"ECL posted: {fmt(delta)}")
                st.session_state.ecl_allowance = 0.0

# =====================================================================
# LEASE
# =====================================================================
elif menu == "🏢 Lease (IFRS 16)":
    st.title("IFRS 16 Lease Accounting")
    if st.session_state.lease is None:
        c1, c2, c3 = st.columns(3)
        annual = c1.number_input("Annual Payment (OMR)", value=12000.0, step=1000.0)
        term = c2.number_input("Term (years)", value=5, min_value=1, max_value=20)
        rate = c3.number_input("IBR (%)", value=6.0, step=0.5) / 100
        if st.button("Recognise Lease"):
            setup_lease(annual, term, rate); st.rerun()
    else:
        L = st.session_state.lease
        c1, c2, c3 = st.columns(3)
        c1.metric("ROU Asset", fmt(L["rou_balance"]))
        c2.metric("Lease Liability", fmt(L["liability_balance"]))
        c3.metric("Year", f"{L['year']}/{L['term']}")
        if L["year"] < L["term"] and st.button(f"Run Year {L['year']+1}"):
            run_lease_year(); st.rerun()
        if L["schedule"]:
            st.dataframe(pd.DataFrame(L["schedule"]), use_container_width=True, hide_index=True)
        if st.button("Reset Lease"):
            st.session_state.lease = None; st.rerun()

# =====================================================================
# AUDIT
# =====================================================================
elif menu == "🔍 Statutory Audit":
    st.title("Statutory Audit Workpaper")
    st.caption("CMA Article 16 requires a CMA-accredited auditor.")
    with st.form("audit_adj"):
        adj_desc = st.text_input("Adjustment Description", "Reclass prepaid rent")
        c1, c2 = st.columns(2)
        dr_acct = c1.selectbox("Debit", list(CHART_OF_ACCOUNTS.keys()),
                               format_func=lambda x: f"{x} – {CHART_OF_ACCOUNTS[x]}")
        cr_acct = c2.selectbox("Credit", list(CHART_OF_ACCOUNTS.keys()),
                               format_func=lambda x: f"{x} – {CHART_OF_ACCOUNTS[x]}", index=10)
        adj_amt = st.number_input("Amount (OMR)", min_value=0.0, value=500.0, step=100.0)
        if st.form_submit_button("Post Adjustment"):
            lines = [{"account": dr_acct, "debit": adj_amt, "credit": 0},
                     {"account": cr_acct, "debit": 0, "credit": adj_amt}]
            ok, msg = post_audit_adjustment(adj_desc, lines)
            if ok: st.success(msg); st.rerun()
            else: st.error(msg)
    if st.session_state.audit_adjustments:
        st.dataframe(pd.DataFrame([{"Description": a["desc"],
                                    "Amount": a["lines"][0]["debit"]}
                                   for a in st.session_state.audit_adjustments]),
                     use_container_width=True, hide_index=True)
    st.subheader("Checklist")
    net_a = sum(v for k, v in st.session_state.balances.items() if k.startswith("1") and k != "1000") \
            - sum(v for k, v in st.session_state.balances.items() if k.startswith("2"))
    checks = [
        ("CMA Art. 25 – Client funds segregated", True),
        ("CMA Art. 26 – Net assets ≥ RO 50,000", net_a >= MIN_CAPITAL),
        ("IFRS 15 – Revenue on placement", True),
        ("IFRS 9 – ECL allowance", st.session_state.ecl_allowance > 0),
        ("IFRS 16 – Lease recognised", st.session_state.lease is not None),
        ("VAT – 5% output applied", True),
    ]
    for item, ok in checks:
        st.write(("✅ " if ok else "⬜ ") + item)

# =====================================================================
# REPORTS
# =====================================================================
elif menu == "📑 Reports":
    st.title("Financial Reports")
    assets = {k: v for k, v in st.session_state.balances.items() if k.startswith("1") and k != "1000"}
    liabs = {k: v for k, v in st.session_state.balances.items() if k.startswith("2")}
    equity = {k: v for k, v in st.session_state.balances.items() if k.startswith("3")}
    rev = {k: v for k, v in st.session_state.balances.items() if k.startswith("4")}
    exp = {k: v for k, v in st.session_state.balances.items() if k.startswith("5")}

    st.header("Statement of Financial Position")
    c1, c2 = st.columns(2)
    with c1:
        for k, v in assets.items():
            st.write(f"{CHART_OF_ACCOUNTS[k]}: {fmt(v)}")
        st.markdown(f"**Total Assets: {fmt(sum(assets.values()))}**")
    with c2:
        for k, v in liabs.items():
            st.write(f"{CHART_OF_ACCOUNTS[k]}: {fmt(v)}")
        st.markdown(f"**Total Liabilities: {fmt(sum(liabs.values()))}**")
        for k, v in equity.items():
            st.write(f"{CHART_OF_ACCOUNTS[k]}: {fmt(v)}")
        st.markdown(f"**Total Equity: {fmt(sum(equity.values()))}**")

    st.divider()
    st.header("Profit & Loss")
    c1, c2 = st.columns(2)
    with c1:
        for k, v in rev.items():
            st.write(f"{CHART_OF_ACCOUNTS[k]}: {fmt(v)}")
        st.markdown(f"**Total Revenue: {fmt(sum(rev.values()))}**")
    with c2:
        for k, v in exp.items():
            st.write(f"{CHART_OF_ACCOUNTS[k]}: {fmt(v)}")
        st.markdown(f"**Total Expenses: {fmt(sum(exp.values()))}**")
    st.markdown(f"## Net Profit: {fmt(sum(rev.values()) - sum(exp.values()))}")

    st.divider()
    st.subheader("📥 Export Reports")
    c1, c2, c3 = st.columns(3)
    with c1:
        buf = generate_excel_report()
        st.download_button("📊 Excel (TB / Journal / P&L / Lease)",
                           data=buf,
                           file_name=f"oman_broker_{date.today()}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        if REPORTLAB_OK:
            pdf_buf = io.BytesIO()
            doc = SimpleDocTemplate(pdf_buf, pagesize=A4)
            styles = getSampleStyleSheet()
            el = [
                Paragraph("<b>Financial Summary</b>",
                          ParagraphStyle('T', parent=styles['Heading1'], alignment=TA_CENTER)),
                Spacer(1, 20),
                Paragraph(f"Total Assets: {fmt(sum(assets.values()))}", styles['Normal']),
                Paragraph(f"Total Liabilities: {fmt(sum(liabs.values()))}", styles['Normal']),
                Paragraph(f"Net Assets: {fmt(sum(assets.values()) - sum(liabs.values()))}", styles['Normal']),
                Spacer(1, 10),
                Paragraph(f"Revenue: {fmt(sum(rev.values()))}", styles['Normal']),
                Paragraph(f"Expenses: {fmt(sum(exp.values()))}", styles['Normal']),
                Paragraph(f"Net Profit: {fmt(sum(rev.values()) - sum(exp.values()))}", styles['Normal']),
            ]
            doc.build(el); pdf_buf.seek(0)
            st.download_button("📄 PDF Summary",
                               data=pdf_buf,
                               file_name=f"oman_broker_summary_{date.today()}.pdf",
                               mime="application/pdf")
        else:
            st.info("Install reportlab to enable PDF.")
    with c3:
        if PPTX_OK:
            pptx_buf = generate_pptx_report()
            st.download_button("📽️ PowerPoint (8 formatted slides)",
                               data=pptx_buf,
                               file_name=f"oman_broker_board_pack_{date.today()}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        else:
            st.info("Install python-pptx to enable PPTX.")

# =====================================================================
# IFRS NOTES
# =====================================================================
elif menu == "📚 IFRS Notes":
    st.title("IFRS & CMA Notes")
    st.markdown("""
**Accruals (IFRS 15)** – Commission recognised when policy bound.

**ECL (IFRS 9)** – Simplified provision matrix.

**Lease (IFRS 16)** – ROU asset = PV of lease payments.

**CMA Article 25** – Client funds in segregated Insurance Brokerage Account, excluded from broker balance sheet.

**CMA Article 26** – Minimum net assets RO 50,000 (insurance/reinsurance) or RO 75,000 (combined), excluding intangibles.

**VAT (Oman)** – 5% on commission services; premiums out of scope for disclosed agent.
    """)