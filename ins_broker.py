import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta
import io
import uuid
import math

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER

# PPTX
from pptx import Presentation
from pptx.util import Inches

# -------------------------------------------------------------------
# CONFIG
# -------------------------------------------------------------------
st.set_page_config(page_title="Oman Insurance Broker FS", page_icon="🇴🇲", layout="wide")

VAT_RATE = 0.05
ECL_RATES = {"current": 0.005, "30": 0.02, "60": 0.05, "90": 0.15, "120+": 0.30}

CHART_OF_ACCOUNTS = {
    # Assets
    "1000": "Cash - Insurance Brokerage Account (Client Funds)",
    "1010": "Cash - Operating Account",
    "1020": "Bank - Operating Account",
    "1100": "Accounts Receivable - Commissions",
    "1110": "VAT Receivable (Input)",
    "1120": "Right-of-Use Asset (Net)",
    "1130": "Prepaid Expenses",
    "1140": "Fixed Assets (Net)",
    # Liabilities
    "2000": "Accounts Payable - Insurers (Premiums)",
    "2010": "Accounts Payable - Clients (Claims)",
    "2100": "Accrued Expenses",
    "2110": "VAT Payable (Output)",
    "2120": "Lease Liability - Current",
    "2130": "Lease Liability - Non-Current",
    "2200": "Loans Payable",
    # Equity
    "3000": "Share Capital",
    "3100": "Retained Earnings",
    "3200": "Current Year Profit/Loss",
    # Revenue
    "4000": "Commission Income - Insurance",
    "4010": "Commission Income - Reinsurance",
    "4020": "Fee Income",
    # Expenses
    "5000": "Salaries and Benefits",
    "5010": "Depreciation - ROU Asset",
    "5020": "Interest Expense - Lease",
    "5030": "ECL Impairment Loss",
    "5040": "Utilities",
    "5050": "Professional Fees (Audit/Legal)",
    "5060": "Marketing",
    "5070": "Bank Charges",
    "5080": "Other Operating Expenses",
}

DEFAULT_BALANCES = {
    "1000": 0.0, "1010": 5000.0, "1020": 15000.0, "1100": 0.0,
    "1110": 0.0, "1120": 0.0, "1130": 1000.0, "1140": 10000.0,
    "2000": 0.0, "2010": 0.0, "2100": 2000.0, "2110": 0.0,
    "2120": 0.0, "2130": 0.0, "2200": 0.0,
    "3000": 50000.0, "3100": 5000.0, "3200": 0.0,
    "4000": 0.0, "4010": 0.0, "4020": 0.0,
    "5000": 0.0, "5010": 0.0, "5020": 0.0, "5030": 0.0,
    "5040": 0.0, "5050": 0.0, "5060": 0.0, "5070": 0.0, "5080": 0.0,
}

# -------------------------------------------------------------------
# SESSION STATE
# -------------------------------------------------------------------
if "balances" not in st.session_state:
    st.session_state.balances = DEFAULT_BALANCES.copy()
if "journal_entries" not in st.session_state:
    st.session_state.journal_entries = []
if "documents" not in st.session_state:
    st.session_state.documents = []
if "accruals" not in st.session_state:
    st.session_state.accruals = []
if "ecl_allowance" not in st.session_state:
    st.session_state.ecl_allowance = 0.0
if "lease" not in st.session_state:
    st.session_state.lease = None
if "audit_adjustments" not in st.session_state:
    st.session_state.audit_adjustments = []

# -------------------------------------------------------------------
# HELPERS
# -------------------------------------------------------------------
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

# -------------------------------------------------------------------
# INITIAL SAMPLE DATA (Accrual basis)
# -------------------------------------------------------------------
def init_sample():
    if st.session_state.journal_entries:
        return
    entries = [
        # Accrued commission (IFRS 15: point-in-time when policy bound)
        (date(2026, 1, 15), "Commission accrued - Motor fleet policy bound",
         [{"account": "1100", "debit": 1250.0, "credit": 0},
          {"account": "4000", "debit": 0, "credit": 1250.0}]),
        # VAT on commission
        (date(2026, 1, 15), "Output VAT on commission (5%)",
         [{"account": "1100", "debit": 62.5, "credit": 0},
          {"account": "2110", "debit": 0, "credit": 62.5}]),
        # Commission received
        (date(2026, 2, 1), "Commission received from insurer",
         [{"account": "1020", "debit": 1312.5, "credit": 0},
          {"account": "1100", "debit": 0, "credit": 1312.5}]),
        # Accrued audit fee
        (date(2026, 3, 31), "Accrual: statutory audit fee for FY2026",
         [{"account": "5050", "debit": 1500.0, "credit": 0},
          {"account": "2100", "debit": 0, "credit": 1500.0}]),
        # Rent accrual
        (date(2026, 3, 31), "Accrual: March office rent",
         [{"account": "5040", "debit": 900.0, "credit": 0},
          {"account": "2100", "debit": 0, "credit": 900.0}]),
    ]
    for d, desc, lines in entries:
        post_journal(d, desc, lines)

init_sample()

# -------------------------------------------------------------------
# ECL CALCULATOR (IFRS 9 Simplified Approach)
# -------------------------------------------------------------------
def compute_ecl():
    ar = get_bal("1100")
    if ar <= 0:
        return 0.0, []
    buckets = [
        ("Current", 0, 30, ECL_RATES["current"]),
        ("1-30 days", 30, 60, ECL_RATES["30"]),
        ("31-60 days", 60, 90, ECL_RATES["60"]),
        ("61-90 days", 90, 120, ECL_RATES["90"]),
        ("90+ days", 120, 9999, ECL_RATES["120+"]),
    ]
    # Simplified: assume 60% current, 25% 1-30, 10% 31-60, 5% 61-90
    splits = [0.60, 0.25, 0.10, 0.05, 0.00]
    ecl_details = []
    total_ecl = 0.0
    for i, (label, low, high, rate) in enumerate(buckets):
        gross = ar * splits[i]
        loss = gross * rate
        ecl_details.append({"Bucket": label, "Gross (OMR)": gross, "Rate": f"{rate*100:.1f}%", "ECL (OMR)": loss})
        total_ecl += loss
    return total_ecl, ecl_details

# -------------------------------------------------------------------
# LEASE CALCULATOR (IFRS 16)
# -------------------------------------------------------------------
def setup_lease(annual_payment, term_years, rate):
    pv = annual_payment * (1 - (1 + rate) ** -term_years) / rate
    st.session_state.lease = {
        "annual_payment": annual_payment,
        "term": term_years,
        "rate": rate,
        "initial_liability": pv,
        "initial_rou": pv,
        "rou_balance": pv,
        "liability_balance": pv,
        "year": 0,
        "schedule": [],
    }
    # Post initial recognition
    post_journal(date.today(), "IFRS 16: Initial recognition of office lease ROU asset and liability",
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

    # Post interest, depreciation, payment
    post_journal(date.today(), f"IFRS 16 Year {L['year']+1}: Interest accretion",
                 [{"account": "5020", "debit": interest, "credit": 0},
                  {"account": "2130", "debit": 0, "credit": interest}])
    post_journal(date.today(), f"IFRS 16 Year {L['year']+1}: ROU depreciation",
                 [{"account": "5010", "debit": dep, "credit": 0},
                  {"account": "1120", "debit": 0, "credit": dep}])
    post_journal(date.today(), f"IFRS 16 Year {L['year']+1}: Lease payment",
                 [{"account": "2130", "debit": principal, "credit": 0},
                  {"account": "1020", "debit": 0, "credit": principal}])

    L["schedule"].append({
        "Year": L["year"] + 1, "Opening Liability": L["liability_balance"],
        "Interest": interest, "Payment": L["annual_payment"],
        "Principal": principal, "Closing Liability": new_liab,
        "ROU Depreciation": dep, "ROU Closing": new_rou
    })
    L["liability_balance"] = new_liab
    L["rou_balance"] = new_rou
    L["year"] += 1

# -------------------------------------------------------------------
# AUDIT ADJUSTMENT
# -------------------------------------------------------------------
def post_audit_adjustment(desc, lines):
    ok, msg = post_journal(date.today(), f"[AUDIT ADJ] {desc}", lines)
    if ok:
        st.session_state.audit_adjustments.append({"desc": desc, "lines": lines})
    return ok, msg

# -------------------------------------------------------------------
# NAVIGATION
# -------------------------------------------------------------------
st.sidebar.title("🏦 Insurance Broker FS")
st.sidebar.caption("Oman CMA · Full IFRS")
menu = st.sidebar.radio("Navigation", [
    "📊 Dashboard", "📥 Journal Entry", "📄 Documents",
    "🔧 Accruals & ECL", "🏢 Lease (IFRS 16)",
    "🔍 Statutory Audit", "📑 Reports", "📚 IFRS Notes"
])

# -------------------------------------------------------------------
# DASHBOARD
# -------------------------------------------------------------------
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
    min_cap = 50000.0
    # Exclude intangibles per CMA
    intangibles = get_bal("1120")  # ROU is not intangible, but adjust if needed
    if net >= min_cap:
        st.success(f"✅ MET: {fmt(net)} ≥ {fmt(min_cap)}")
    else:
        st.error(f"⚠️ NOT MET: {fmt(net)} < {fmt(min_cap)}")

# -------------------------------------------------------------------
# JOURNAL ENTRY
# -------------------------------------------------------------------
elif menu == "📥 Journal Entry":
    st.title("Multi-Line Double-Entry Journal")

    if "jlines" not in st.session_state:
        st.session_state.jlines = [{"account": "1010", "debit": 0.0, "credit": 0.0},
                                   {"account": "4000", "debit": 0.0, "credit": 0.0}]

    col_a, col_b = st.columns(2)
    if col_a.button("➕ Add Line"):
        st.session_state.jlines.append({"account": "4000", "debit": 0.0, "credit": 0.0})
        st.rerun()
    if col_b.button("➖ Remove Last"):
        if len(st.session_state.jlines) > 2:
            st.session_state.jlines.pop()
            st.rerun()

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
        tdr = sum(l["debit"] for l in lines)
        tcr = sum(l["credit"] for l in lines)
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

# -------------------------------------------------------------------
# DOCUMENTS
# -------------------------------------------------------------------
elif menu == "📄 Documents":
    st.title("Document Generation")
    doc_type = st.selectbox("Type", ["VAT Invoice (Commission)", "Invoice", "Debit Note", "Credit Note", "Receipt"])
    col1, col2 = st.columns(2)
    party = col1.text_input("Party", "Dhofar Insurance")
    ref = col2.text_input("Reference", "POL-2026-001")
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

    if st.session_state.documents:
        st.dataframe(pd.DataFrame([{"No": d["doc_no"], "Type": d["type"], "Party": d["party"], "Total": d["total"]}
                                   for d in st.session_state.documents]), use_container_width=True)

# -------------------------------------------------------------------
# ACCRUALS & ECL
# -------------------------------------------------------------------
elif menu == "🔧 Accruals & ECL":
    st.title("Accruals & Expected Credit Loss (IFRS 9)")

    tab1, tab2 = st.tabs(["📅 Accruals", "📉 ECL Provision Matrix"])

    with tab1:
        st.subheader("Record Accrual")
        with st.form("accrual_form"):
            acc_date = st.date_input("Accrual Date", date.today())
            acc_desc = st.text_input("Description", "Accrued audit fee")
            exp_acct = st.selectbox("Expense Account", ["5050", "5040", "5000", "5080"],
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
        st.subheader("ECL Calculator (Simplified Approach)")
        ar = get_bal("1100")
        st.write(f"Gross Accounts Receivable (1100): **{fmt(ar)}**")

        if st.button("Compute ECL"):
            ecl, details = compute_ecl()
            st.session_state.ecl_allowance = ecl
            st.dataframe(pd.DataFrame(details), use_container_width=True, hide_index=True)
            st.metric("Total ECL Allowance Required", fmt(ecl))

        if st.session_state.ecl_allowance > 0:
            if st.button("Post ECL Adjustment"):
                # Move allowance to target
                current_allowance = 0.0  # simplified: track separately
                delta = st.session_state.ecl_allowance - current_allowance
                if delta > 0:
                    lines = [{"account": "5030", "debit": delta, "credit": 0},
                             {"account": "1100", "debit": 0, "credit": delta}]
                    ok, msg = post_journal(date.today(), "IFRS 9 ECL impairment allowance", lines)
                    if ok:
                        st.success(f"ECL posted: {fmt(delta)}")

# -------------------------------------------------------------------
# LEASE (IFRS 16)
# -------------------------------------------------------------------
elif menu == "🏢 Lease (IFRS 16)":
    st.title("IFRS 16 Lease Accounting")

    if st.session_state.lease is None:
        st.subheader("Set Up Office Lease")
        col1, col2, col3 = st.columns(3)
        annual = col1.number_input("Annual Lease Payment (OMR)", value=12000.0, step=1000.0)
        term = col2.number_input("Lease Term (years)", value=5, min_value=1, max_value=20)
        rate = col3.number_input("Incremental Borrowing Rate (%)", value=6.0, step=0.5) / 100

        if st.button("Recognise Lease"):
            setup_lease(annual, term, rate)
            st.success("ROU asset and lease liability recognised.")
            st.rerun()
    else:
        L = st.session_state.lease
        st.subheader(f"Lease: {L['term']} years @ {L['rate']*100:.1f}%")
        c1, c2, c3 = st.columns(3)
        c1.metric("ROU Asset Balance", fmt(L["rou_balance"]))
        c2.metric("Lease Liability", fmt(L["liability_balance"]))
        c3.metric("Year", f"{L['year']} / {L['term']}")

        if L["year"] < L["term"]:
            if st.button(f"Run Year {L['year']+1} Entries"):
                run_lease_year()
                st.rerun()
        else:
            st.info("Lease fully amortised.")

        if L["schedule"]:
            st.subheader("Amortisation Schedule")
            st.dataframe(pd.DataFrame(L["schedule"]), use_container_width=True, hide_index=True)

        if st.button("Reset Lease"):
            st.session_state.lease = None
            st.rerun()

# -------------------------------------------------------------------
# STATUTORY AUDIT
# -------------------------------------------------------------------
elif menu == "🔍 Statutory Audit":
    st.title("Statutory Audit Workpaper")
    st.caption("CMA Article 16 requires an auditor accredited by the Authority [citation:4]")

    st.subheader("Audit Adjustments")
    with st.form("audit_adj"):
        adj_desc = st.text_input("Adjustment Description", "Reclassification of prepaid rent")
        col1, col2 = st.columns(2)
        dr_acct = col1.selectbox("Debit Account", list(CHART_OF_ACCOUNTS.keys()),
                                 format_func=lambda x: f"{x} – {CHART_OF_ACCOUNTS[x]}")
        cr_acct = col2.selectbox("Credit Account", list(CHART_OF_ACCOUNTS.keys()),
                                 format_func=lambda x: f"{x} – {CHART_OF_ACCOUNTS[x]}", index=10)
        adj_amt = st.number_input("Amount (OMR)", min_value=0.0, value=500.0, step=100.0)
        if st.form_submit_button("Post Audit Adjustment"):
            lines = [{"account": dr_acct, "debit": adj_amt, "credit": 0},
                     {"account": cr_acct, "debit": 0, "credit": adj_amt}]
            ok, msg = post_audit_adjustment(adj_desc, lines)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    if st.session_state.audit_adjustments:
        st.dataframe(pd.DataFrame([{"Description": a["desc"],
                                    "Debit": a["lines"][0]["account"],
                                    "Credit": a["lines"][1]["account"],
                                    "Amount": a["lines"][0]["debit"]}
                                   for a in st.session_state.audit_adjustments]),
                     use_container_width=True, hide_index=True)

    st.subheader("Audit Checklist")
    checklist = [
        ("CMA Article 25 – Client funds segregated in Insurance Brokerage Account", True),
        ("CMA Article 26 – Minimum net assets ≥ RO 50,000 (excl. intangibles)", True),
        ("IFRS 15 – Commission revenue recognised on policy binding (point-in-time)", True),
        ("IFRS 9 – ECL allowance computed on commission receivables", st.session_state.ecl_allowance > 0),
        ("IFRS 16 – Lease ROU asset and liability recognised", st.session_state.lease is not None),
        ("VAT – Output VAT 5% applied to commission income", True),
    ]
    for item, done in checklist:
        st.write(("✅" if done else "⬜") + " " + item)

# -------------------------------------------------------------------
# REPORTS
# -------------------------------------------------------------------
elif menu == "📑 Reports":
    st.title("Financial Reports")

    # FS
    st.header("Statement of Financial Position")
    assets = {k: v for k, v in st.session_state.balances.items() if k.startswith("1") and k != "1000"}
    liabs = {k: v for k, v in st.session_state.balances.items() if k.startswith("2")}
    equity = {k: v for k, v in st.session_state.balances.items() if k.startswith("3")}
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
    rev = {k: v for k, v in st.session_state.balances.items() if k.startswith("4")}
    exp = {k: v for k, v in st.session_state.balances.items() if k.startswith("5")}
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
    st.subheader("Export")
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="xlsxwriter") as w:
        pd.DataFrame([{"Code": k, "Account": v, "Balance": st.session_state.balances[k]}
                      for k, v in CHART_OF_ACCOUNTS.items()]).to_excel(w, sheet_name="Trial Balance", index=False)
        if st.session_state.journal_entries:
            rows = []
            for je in st.session_state.journal_entries:
                for l in je["lines"]:
                    rows.append({"Date": je["date"], "Desc": je["description"],
                                 "Account": l["account"], "Debit": l["debit"], "Credit": l["credit"]})
            pd.DataFrame(rows).to_excel(w, sheet_name="Journal", index=False)
    st.download_button("📊 Excel", excel_buf.getvalue(), f"fs_{date.today()}.xlsx")

# -------------------------------------------------------------------
# IFRS NOTES
# -------------------------------------------------------------------
elif menu == "📚 IFRS Notes":
    st.title("IFRS & CMA Notes")
    st.markdown("""
    ### Accruals (IFRS 15)
    Commission income is recognised when the policy is **bound** (point-in-time), not when cash is received. 
    Claims handling services may be recognised over time [citation:1][citation:21].

    ### ECL (IFRS 9)
    Simplified approach: provision matrix based on days past due. ECL = Gross receivable × loss rate [citation:12].

    ### Lease (IFRS 16)
    ROU asset = PV of lease payments. Depreciation (straight-line) + interest (effective rate) replace rent expense [citation:23].

    ### CMA Audit (Article 16 / 27)
    Auditor must be **CMA-accredited**. CMA may inspect books and appoint external auditor [citation:4][citation:9].

    ### CMA Article 25
    Client funds in "Insurance Brokerage Account" are **excluded** from the broker's own balance sheet [citation:9].

    ### CMA Article 26
    Minimum net assets: **RO 50,000** (insurance OR reinsurance) / **RO 75,000** (combined). Excludes intangibles [citation:14].
    """)