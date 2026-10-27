import streamlit as st
import pandas as pd
from datetime import date, datetime
import io
import uuid

# -------------------------------------------------------------------
# PDF Generation
# -------------------------------------------------------------------
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT

# -------------------------------------------------------------------
# PPTX Generation
# -------------------------------------------------------------------
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# -------------------------------------------------------------------
# 1. PAGE CONFIGURATION
# -------------------------------------------------------------------
st.set_page_config(
    page_title="Oman Insurance Broker Financials",
    page_icon="🇴🇲",
    layout="wide"
)

# -------------------------------------------------------------------
# 2. CHART OF ACCOUNTS (Oman Insurance Broker Regulations)
# -------------------------------------------------------------------
CHART_OF_ACCOUNTS = {
    # ASSETS
    "1000": "Cash - Insurance Brokerage Account (Client Funds)",
    "1010": "Cash - Operating Account (Broker Funds)",
    "1020": "Bank - Operating Account",
    "1100": "Accounts Receivable - Insurance Companies (Commissions)",
    "1110": "VAT Receivable (Input)",
    "1200": "Prepaid Expenses",
    "1300": "Fixed Assets (Net)",
    "1400": "Intangible Assets",
    # LIABILITIES
    "2000": "Accounts Payable - Insurance Companies (Premiums)",
    "2010": "Accounts Payable - Clients (Claims)",
    "2100": "Accrued Expenses",
    "2110": "VAT Payable (Output)",
    "2200": "Loans Payable",
    # EQUITY
    "3000": "Share Capital",
    "3100": "Retained Earnings",
    "3200": "Current Year Profit/Loss",
    # REVENUE
    "4000": "Commission Income - Insurance",
    "4010": "Commission Income - Reinsurance",
    "4020": "Fee Income",
    "4100": "Interest Income",
    # EXPENSES
    "5000": "Salaries and Benefits",
    "5010": "Rent Expense",
    "5020": "Utilities",
    "5030": "Professional Fees (Audit/Legal)",
    "5040": "Marketing and Advertising",
    "5050": "Depreciation",
    "5060": "Bank Charges",
    "5070": "Other Operating Expenses",
}

DEFAULT_BALANCES = {
    "1000": 0.0, "1010": 5000.0, "1020": 15000.0, "1100": 8000.0,
    "1110": 0.0, "1200": 1000.0, "1300": 10000.0, "1400": 0.0,
    "2000": 0.0, "2010": 0.0, "2100": 2000.0, "2110": 0.0, "2200": 0.0,
    "3000": 50000.0, "3100": 5000.0, "3200": 0.0,
    "4000": 0.0, "4010": 0.0, "4020": 0.0, "4100": 0.0,
    "5000": 0.0, "5010": 0.0, "5020": 0.0, "5030": 0.0,
    "5040": 0.0, "5050": 0.0, "5060": 0.0, "5070": 0.0,
}

VAT_RATE = 0.05  # Oman standard VAT rate

# -------------------------------------------------------------------
# 3. SESSION STATE INITIALIZATION
# -------------------------------------------------------------------
if "balances" not in st.session_state:
    st.session_state.balances = DEFAULT_BALANCES.copy()

if "journal_entries" not in st.session_state:
    st.session_state.journal_entries = []

if "documents" not in st.session_state:
    st.session_state.documents = []

# -------------------------------------------------------------------
# 4. HELPER FUNCTIONS
# -------------------------------------------------------------------
def get_balance(code):
    return st.session_state.balances.get(code, 0.0)

def format_omr(val):
    return f"{val:,.3f} OMR"

def post_journal_entry(entry_date, description, lines):
    """Post a double-entry journal (list of dicts with account, debit, credit)."""
    total_debit = sum(l["debit"] for l in lines)
    total_credit = sum(l["credit"] for l in lines)
    if abs(total_debit - total_credit) > 0.001:
        return False, "Debits and credits do not balance."
    if total_debit == 0:
        return False, "Amount must be greater than zero."

    # Apply to balances (simplified: debit increases asset/expense, credit increases liab/equity/rev)
    for l in lines:
        code = l["account"]
        net = l["debit"] - l["credit"]
        # Asset/Expense accounts (1xxx, 5xxx): debit positive
        # Liability/Equity/Revenue (2xxx, 3xxx, 4xxx): credit positive
        if code.startswith("1") or code.startswith("5"):
            st.session_state.balances[code] = st.session_state.balances.get(code, 0) + net
        else:
            st.session_state.balances[code] = st.session_state.balances.get(code, 0) - net

    st.session_state.journal_entries.append({
        "id": str(uuid.uuid4())[:8],
        "date": entry_date,
        "description": description,
        "lines": lines,
        "total": total_debit,
        "posted_at": datetime.now().isoformat()
    })
    return True, f"Journal posted: {total_debit:,.3f} OMR"

def load_sample_data():
    """Load illustrative Oman insurance broker transactions."""
    if st.session_state.journal_entries:
        return

    samples = [
        # 1. Commission earned (VAT invoice to insurer)
        {
            "date": date(2026, 1, 15),
            "desc": "Commission - Motor fleet (Dhofar Insurance) + 5% VAT",
            "lines": [
                {"account": "1100", "debit": 1312.50, "credit": 0},
                {"account": "4000", "debit": 0, "credit": 1250.00},
                {"account": "2110", "debit": 0, "credit": 62.50},
            ]
        },
        # 2. Premium received from client (client funds - out of scope for VAT)
        {
            "date": date(2026, 1, 20),
            "desc": "Premium received from client (client funds)",
            "lines": [
                {"account": "1000", "debit": 8500.00, "credit": 0},
                {"account": "2000", "debit": 0, "credit": 8500.00},
            ]
        },
        # 3. Premium remitted to insurer
        {
            "date": date(2026, 1, 22),
            "desc": "Premium remitted to insurer",
            "lines": [
                {"account": "2000", "debit": 8500.00, "credit": 0},
                {"account": "1000", "debit": 0, "credit": 8500.00},
            ]
        },
        # 4. Rent expense
        {
            "date": date(2026, 2, 1),
            "desc": "Office rent - Muscat",
            "lines": [
                {"account": "5010", "debit": 900.00, "credit": 0},
                {"account": "1010", "debit": 0, "credit": 900.00},
            ]
        },
        # 5. Commission - Reinsurance
        {
            "date": date(2026, 2, 10),
            "desc": "Commission - Marine cargo (reinsurance) + 5% VAT",
            "lines": [
                {"account": "1100", "debit": 2205.79, "credit": 0},
                {"account": "4010", "debit": 0, "credit": 2100.75},
                {"account": "2110", "debit": 0, "credit": 105.04},
            ]
        },
        # 6. Salaries
        {
            "date": date(2026, 2, 15),
            "desc": "Salaries paid - February",
            "lines": [
                {"account": "5000", "debit": 3200.00, "credit": 0},
                {"account": "1010", "debit": 0, "credit": 3200.00},
            ]
        },
        # 7. Bank charges
        {
            "date": date(2026, 3, 1),
            "desc": "Bank charges",
            "lines": [
                {"account": "5060", "debit": 45.50, "credit": 0},
                {"account": "1020", "debit": 0, "credit": 45.50},
            ]
        },
        # 8. Commission receivable
        {
            "date": date(2026, 3, 10),
            "desc": "Commission receivable from insurer",
            "lines": [
                {"account": "1100", "debit": 1890.00, "credit": 0},
                {"account": "4000", "debit": 0, "credit": 1800.00},
                {"account": "2110", "debit": 0, "credit": 90.00},
            ]
        },
        # 9. Audit fees
        {
            "date": date(2026, 3, 15),
            "desc": "Audit fees - annual statutory audit",
            "lines": [
                {"account": "5030", "debit": 1200.00, "credit": 0},
                {"account": "1010", "debit": 0, "credit": 1200.00},
            ]
        },
    ]

    for s in samples:
        post_journal_entry(s["date"], s["desc"], s["lines"])

load_sample_data()

# -------------------------------------------------------------------
# 5. DOCUMENT GENERATION FUNCTIONS
# -------------------------------------------------------------------
def generate_pdf_document(doc_type, doc_data):
    """Generate a PDF for a document (Invoice, DN, CN, Receipt)."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20*mm, leftMargin=20*mm,
                            topMargin=20*mm, bottomMargin=20*mm)
    styles = getSampleStyleSheet()
    elements = []

    # Header
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=18,
                                  alignment=TA_CENTER, spaceAfter=20)
    elements.append(Paragraph(f"<b>{doc_type.upper()}</b>", title_style))
    elements.append(Spacer(1, 10))

    # Company info (demo)
    company_info = [
        ["<b>Al Noor Insurance Brokerage LLC</b>", ""],
        ["Muscat, Sultanate of Oman", ""],
        ["VAT No: OM1100XXXXXX", ""],
    ]
    t = Table(company_info, colWidths=[100*mm, 60*mm])
    t.setStyle(TableStyle([('FONTSIZE', (0, 0), (-1, -1), 9),
                           ('ALIGN', (0, 0), (0, -1), 'LEFT')]))
    elements.append(t)
    elements.append(Spacer(1, 15))

    # Document details
    details = [
        ["<b>Document No:</b>", doc_data.get("doc_no", "N/A")],
        ["<b>Date:</b>", str(doc_data.get("date", date.today()))],
        ["<b>Party:</b>", doc_data.get("party", "N/A")],
        ["<b>Reference:</b>", doc_data.get("reference", "N/A")],
    ]
    t2 = Table(details, colWidths=[40*mm, 120*mm])
    t2.setStyle(TableStyle([('FONTSIZE', (0, 0), (-1, -1), 10),
                            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                            ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey)]))
    elements.append(t2)
    elements.append(Spacer(1, 15))

    # Line items
    items = doc_data.get("items", [])
    if items:
        header = ["Description", "Amount (OMR)"]
        data = [header] + [[i["desc"], f"{i['amount']:,.3f}"] for i in items]
        # Totals
        subtotal = sum(i["amount"] for i in items)
        vat = doc_data.get("vat", 0)
        total = subtotal + vat
        data.append(["", ""])
        data.append(["<b>Subtotal</b>", f"<b>{subtotal:,.3f}</b>"])
        if vat > 0:
            data.append(["<b>VAT (5%)</b>", f"<b>{vat:,.3f}</b>"])
        data.append(["<b>Total</b>", f"<b>{total:,.3f}</b>"])

        t3 = Table(data, colWidths=[120*mm, 40*mm])
        t3.setStyle(TableStyle([
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a5490')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ]))
        elements.append(t3)

    # Footer note
    elements.append(Spacer(1, 30))
    footer = Paragraph(
        "<i>This is a computer-generated document. No signature is required.</i>",
        ParagraphStyle('Footer', parent=styles['Normal'], fontSize=8,
                       alignment=TA_CENTER, textColor=colors.grey))
    elements.append(footer)

    doc.build(elements)
    buffer.seek(0)
    return buffer

def generate_excel_report():
    """Generate a multi-sheet Excel workbook with Trial Balance, Journal, and Reports."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Sheet 1: Trial Balance
        tb_data = [{"Code": k, "Account": CHART_OF_ACCOUNTS[k], "Balance (OMR)": v}
                   for k, v in st.session_state.balances.items()]
        pd.DataFrame(tb_data).to_excel(writer, sheet_name='Trial Balance', index=False)

        # Sheet 2: Journal Entries
        je_rows = []
        for je in st.session_state.journal_entries:
            for l in je["lines"]:
                if l["debit"] > 0 or l["credit"] > 0:
                    je_rows.append({
                        "Date": je["date"], "Journal ID": je["id"],
                        "Description": je["description"], "Account": l["account"],
                        "Account Name": CHART_OF_ACCOUNTS.get(l["account"], ""),
                        "Debit": l["debit"], "Credit": l["credit"]
                    })
        if je_rows:
            pd.DataFrame(je_rows).to_excel(writer, sheet_name='Journal Entries', index=False)

        # Sheet 3: P&L
        revenue = {k: v for k, v in st.session_state.balances.items() if k.startswith("4")}
        expenses = {k: v for k, v in st.session_state.balances.items() if k.startswith("5")}
        pl_data = [{"Account": CHART_OF_ACCOUNTS[k], "Amount": v} for k, v in revenue.items()]
        pl_data.append({"Account": "TOTAL REVENUE", "Amount": sum(revenue.values())})
        pl_data.append({"Account": "", "Amount": 0})
        pl_data += [{"Account": CHART_OF_ACCOUNTS[k], "Amount": v} for k, v in expenses.items()]
        pl_data.append({"Account": "TOTAL EXPENSES", "Amount": sum(expenses.values())})
        pl_data.append({"Account": "NET PROFIT / (LOSS)", "Amount": sum(revenue.values()) - sum(expenses.values())})
        pd.DataFrame(pl_data).to_excel(writer, sheet_name='P&L', index=False)

    output.seek(0)
    return output

def generate_pptx_report():
    """Generate a PowerPoint summary of financials."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1: Title
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Oman Insurance Broker"
    slide.placeholders[1].text = f"Financial Summary\n{date.today().strftime('%d %B %Y')}"

    # Slide 2: Key Metrics
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Key Financial Metrics"

    total_assets = sum(v for k, v in st.session_state.balances.items() if k.startswith("1"))
    total_liab = sum(v for k, v in st.session_state.balances.items() if k.startswith("2"))
    net_assets = total_assets - total_liab
    revenue = sum(v for k, v in st.session_state.balances.items() if k.startswith("4"))
    expenses = sum(v for k, v in st.session_state.balances.items() if k.startswith("5"))
    net_profit = revenue - expenses

    tf = slide.placeholders[1].text_frame
    tf.text = f"Total Assets: {format_omr(total_assets)}"
    p = tf.add_paragraph()
    p.text = f"Total Liabilities: {format_omr(total_liab)}"
    p = tf.add_paragraph()
    p.text = f"Net Assets: {format_omr(net_assets)}"
    p = tf.add_paragraph()
    p.text = f"Revenue: {format_omr(revenue)}"
    p = tf.add_paragraph()
    p.text = f"Expenses: {format_omr(expenses)}"
    p = tf.add_paragraph()
    p.text = f"Net Profit / (Loss): {format_omr(net_profit)}"

    # Slide 3: Compliance
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "CMA Compliance Check"
    tf = slide.placeholders[1].text_frame
    min_capital = 50000.0
    status = "✅ MET" if net_assets >= min_capital else "⚠️ NOT MET"
    tf.text = f"Minimum Net Assets Required: {format_omr(min_capital)}"
    p = tf.add_paragraph()
    p.text = f"Actual Net Assets: {format_omr(net_assets)}"
    p = tf.add_paragraph()
    p.text = f"Status: {status}"

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

def create_document_record(doc_type, party, reference, items, vat=0):
    """Store a document record in session state."""
    doc_no = f"{doc_type[:3].upper()}-{datetime.now().strftime('%Y%m%d')}-{len(st.session_state.documents)+1:04d}"
    doc = {
        "doc_no": doc_no,
        "type": doc_type,
        "date": date.today(),
        "party": party,
        "reference": reference,
        "items": items,
        "vat": vat,
        "total": sum(i["amount"] for i in items) + vat
    }
    st.session_state.documents.append(doc)
    return doc

# -------------------------------------------------------------------
# 6. APP LAYOUT & NAVIGATION
# -------------------------------------------------------------------
st.sidebar.title("🏦 Insurance Broker FS")
st.sidebar.caption("Oman CMA Compliant Template")

menu = st.sidebar.radio(
    "Navigation",
    ["📊 Dashboard", "📝 Chart of Accounts", "📥 Journal Entry",
     "📄 Documents", "📑 Reports", "📚 IFRS Notes"]
)

# -------------------------------------------------------------------
# 7. PAGE: DASHBOARD
# -------------------------------------------------------------------
if menu == "📊 Dashboard":
    st.title("Insurance Broker Dashboard")
    st.markdown("### Oman Regulatory Overview")

    col1, col2, col3 = st.columns(3)
    total_assets = sum(v for k, v in st.session_state.balances.items() if k.startswith("1"))
    total_liabilities = sum(v for k, v in st.session_state.balances.items() if k.startswith("2"))
    net_assets = total_assets - total_liabilities

    with col1:
        st.metric("Total Assets", format_omr(total_assets))
    with col2:
        st.metric("Total Liabilities", format_omr(total_liabilities))
    with col3:
        st.metric("Net Assets (Equity)", format_omr(net_assets))

    st.subheader("Compliance Check (CMA Article 26)")
    min_capital = 50000.0
    if net_assets >= min_capital:
        st.success(f"✅ Solvency Met: Net Assets ({format_omr(net_assets)}) ≥ {format_omr(min_capital)}")
    else:
        st.error(f"⚠️ Solvency Not Met: Net Assets ({format_omr(net_assets)}) < {format_omr(min_capital)}")

    st.info("Client funds (Account 1000) are excluded from broker net assets per CMA regulations.")

# -------------------------------------------------------------------
# 8. PAGE: CHART OF ACCOUNTS
# -------------------------------------------------------------------
elif menu == "📝 Chart of Accounts":
    st.title("Chart of Accounts")
    df = pd.DataFrame([
        {"Code": k, "Account": v, "Balance (OMR)": get_balance(k)}
        for k, v in CHART_OF_ACCOUNTS.items()
    ])
    st.dataframe(df, use_container_width=True)

# -------------------------------------------------------------------
# 9. PAGE: JOURNAL ENTRY (Multi-line Double Entry)
# -------------------------------------------------------------------
elif menu == "📥 Journal Entry":
    st.title("Multi-Line Double-Entry Journal")

    if "journal_lines" not in st.session_state:
        st.session_state.journal_lines = [
            {"account": "1010", "debit": 0.0, "credit": 0.0},
            {"account": "4000", "debit": 0.0, "credit": 0.0},
        ]

    # Add/remove lines
    col_a, col_b = st.columns([1, 1])
    with col_a:
        if st.button("➕ Add Line"):
            st.session_state.journal_lines.append(
                {"account": "4000", "debit": 0.0, "credit": 0.0}
            )
            st.rerun()
    with col_b:
        if st.button("➖ Remove Last Line") and len(st.session_state.journal_lines) > 2:
            st.session_state.journal_lines.pop()
            st.rerun()

    # Entry form
    with st.form("journal_form"):
        entry_date = st.date_input("Date", date.today())
        description = st.text_input("Description", placeholder="e.g., Commission earned - Motor policy")

        st.markdown("**Journal Lines**")
        lines_to_post = []

        for i, line in enumerate(st.session_state.journal_lines):
            c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
            with c1:
                acct = st.selectbox(
                    f"Account {i+1}", options=list(CHART_OF_ACCOUNTS.keys()),
                    index=list(CHART_OF_ACCOUNTS.keys()).index(line["account"]) if line["account"] in CHART_OF_ACCOUNTS else 0,
                    format_func=lambda x: f"{x} - {CHART_OF_ACCOUNTS[x]}",
                    key=f"acct_{i}"
                )
            with c2:
                debit = st.number_input("Debit", min_value=0.0, step=1.0, value=float(line["debit"]), key=f"dr_{i}")
            with c3:
                credit = st.number_input("Credit", min_value=0.0, step=1.0, value=float(line["credit"]), key=f"cr_{i}")
            with c4:
                st.write("")  # spacing

            lines_to_post.append({"account": acct, "debit": debit, "credit": credit})

        # Totals
        total_dr = sum(l["debit"] for l in lines_to_post)
        total_cr = sum(l["credit"] for l in lines_to_post)

        st.markdown(f"**Total Debit:** {format_omr(total_dr)} | **Total Credit:** {format_omr(total_cr)}")
        if abs(total_dr - total_cr) < 0.001 and total_dr > 0:
            st.success("✓ Balanced")
        else:
            st.warning("⚠ Not balanced")

        if st.form_submit_button("Post Journal Entry"):
            if not description:
                st.error("Description required.")
            else:
                ok, msg = post_journal_entry(entry_date, description, lines_to_post)
                if ok:
                    st.success(msg)
                    st.session_state.journal_lines = [
                        {"account": "1010", "debit": 0.0, "credit": 0.0},
                        {"account": "4000", "debit": 0.0, "credit": 0.0},
                    ]
                    st.rerun()
                else:
                    st.error(msg)

    # Recent entries
    if st.session_state.journal_entries:
        st.subheader("Recent Journal Entries")
        for je in reversed(st.session_state.journal_entries[-10:]):
            with st.expander(f"{je['date']} - {je['description']} - {format_omr(je['total'])}"):
                rows = []
                for l in je["lines"]:
                    if l["debit"] > 0 or l["credit"] > 0:
                        rows.append({
                            "Account": f"{l['account']} - {CHART_OF_ACCOUNTS.get(l['account'], '')}",
                            "Debit": l["debit"],
                            "Credit": l["credit"]
                        })
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# -------------------------------------------------------------------
# 10. PAGE: DOCUMENTS (DN / CN / Receipt / Invoice / VAT Invoice)
# -------------------------------------------------------------------
elif menu == "📄 Documents":
    st.title("Document Generation")
    st.caption("Generate Debit Notes, Credit Notes, Receipts, Invoices, and VAT Invoices for Oman insurance broker transactions.")

    tab1, tab2, tab3 = st.tabs(["📋 Create Document", "📜 Document Register", "⬇️ Download"])

    with tab1:
        doc_type = st.selectbox(
            "Document Type",
            ["VAT Invoice (Commission)", "Invoice", "Debit Note (DN)", "Credit Note (CN)", "Receipt"]
        )

        col1, col2 = st.columns(2)
        with col1:
            party = st.text_input("Party Name", value="Dhofar Insurance Company")
            reference = st.text_input("Reference (Policy No. / Claim No.)", value="POL-2026-001")
        with col2:
            item_desc = st.text_input("Item Description", value="Insurance commission - Motor fleet policy")
            amount = st.number_input("Amount (OMR)", min_value=0.0, value=1250.00, step=10.0)

        apply_vat = st.checkbox("Apply VAT (5%)", value=(doc_type == "VAT Invoice (Commission)"))

        if st.button("Generate Document"):
            vat = round(amount * VAT_RATE, 3) if apply_vat else 0.0
            items = [{"desc": item_desc, "amount": amount}]
            doc = create_document_record(doc_type, party, reference, items, vat)
            st.success(f"Document {doc['doc_no']} created. Total: {format_omr(doc['total'])}")

            # Preview
            pdf_buf = generate_pdf_document(doc_type, doc)
            st.download_button(
                "⬇️ Download PDF Preview",
                data=pdf_buf,
                file_name=f"{doc['doc_no']}.pdf",
                mime="application/pdf"
            )

    with tab2:
        if st.session_state.documents:
            df_docs = pd.DataFrame([
                {"Doc No": d["doc_no"], "Type": d["type"], "Party": d["party"],
                 "Date": d["date"], "Total (OMR)": d["total"]}
                for d in st.session_state.documents
            ])
            st.dataframe(df_docs, use_container_width=True)
        else:
            st.info("No documents generated yet.")

    with tab3:
        if st.session_state.documents:
            selected = st.selectbox(
                "Select Document",
                options=[d["doc_no"] for d in st.session_state.documents]
            )
            doc = next(d for d in st.session_state.documents if d["doc_no"] == selected)
            pdf_buf = generate_pdf_document(doc["type"], doc)
            st.download_button(
                f"⬇️ Download {selected}.pdf",
                data=pdf_buf,
                file_name=f"{selected}.pdf",
                mime="application/pdf"
            )

# -------------------------------------------------------------------
# 11. PAGE: REPORTS (FS / PL / Cash Flow + Exports)
# -------------------------------------------------------------------
elif menu == "📑 Reports":
    st.title("Financial Reports")

    # --- FS ---
    st.header("Statement of Financial Position")
    assets = {k: v for k, v in st.session_state.balances.items() if k.startswith("1")}
    liabilities = {k: v for k, v in st.session_state.balances.items() if k.startswith("2")}
    equity = {k: v for k, v in st.session_state.balances.items() if k.startswith("3")}
    total_assets = sum(assets.values())
    total_liab = sum(liabilities.values())
    total_equity = sum(equity.values())

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Assets")
        for code, val in assets.items():
            st.write(f"{CHART_OF_ACCOUNTS[code]}: **{format_omr(val)}**")
        st.markdown(f"### Total Assets: {format_omr(total_assets)}")
    with col2:
        st.subheader("Liabilities & Equity")
        st.markdown("**Liabilities**")
        for code, val in liabilities.items():
            st.write(f"{CHART_OF_ACCOUNTS[code]}: **{format_omr(val)}**")
        st.markdown(f"**Total Liabilities: {format_omr(total_liab)}**")
        st.markdown("**Equity**")
        for code, val in equity.items():
            st.write(f"{CHART_OF_ACCOUNTS[code]}: **{format_omr(val)}**")
        st.markdown(f"**Total Equity: {format_omr(total_equity)}**")
        st.markdown(f"### Total Liab. & Equity: {format_omr(total_liab + total_equity)}")

    st.divider()

    # --- P&L ---
    st.header("Profit & Loss Statement")
    revenue = {k: v for k, v in st.session_state.balances.items() if k.startswith("4")}
    expenses = {k: v for k, v in st.session_state.balances.items() if k.startswith("5")}
    total_rev = sum(revenue.values())
    total_exp = sum(expenses.values())
    net_profit = total_rev - total_exp

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Revenue")
        for code, val in revenue.items():
            st.write(f"{CHART_OF_ACCOUNTS[code]}: **{format_omr(val)}**")
        st.markdown(f"### Total Revenue: {format_omr(total_rev)}")
    with col2:
        st.subheader("Expenses")
        for code, val in expenses.items():
            st.write(f"{CHART_OF_ACCOUNTS[code]}: **{format_omr(val)}**")
        st.markdown(f"### Total Expenses: {format_omr(total_exp)}")
    st.markdown(f"## Net Profit / (Loss): {format_omr(net_profit)}")

    st.divider()

    # --- Exports ---
    st.subheader("📥 Export Reports")
    col_ex1, col_ex2, col_ex3 = st.columns(3)

    with col_ex1:
        excel_buf = generate_excel_report()
        st.download_button(
            "📊 Download Excel",
            data=excel_buf,
            file_name=f"oman_broker_reports_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_ex2:
        # PDF Summary
        pdf_buf = io.BytesIO()
        doc = SimpleDocTemplate(pdf_buf, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = [
            Paragraph("<b>Financial Summary</b>", ParagraphStyle('T', parent=styles['Heading1'], alignment=TA_CENTER)),
            Spacer(1, 20),
            Paragraph(f"Total Assets: {format_omr(total_assets)}", styles['Normal']),
            Paragraph(f"Total Liabilities: {format_omr(total_liab)}", styles['Normal']),
            Paragraph(f"Net Assets: {format_omr(total_assets - total_liab)}", styles['Normal']),
            Spacer(1, 10),
            Paragraph(f"Revenue: {format_omr(total_rev)}", styles['Normal']),
            Paragraph(f"Expenses: {format_omr(total_exp)}", styles['Normal']),
            Paragraph(f"Net Profit: {format_omr(net_profit)}", styles['Normal']),
        ]
        doc.build(elements)
        pdf_buf.seek(0)
        st.download_button(
            "📄 Download PDF Summary",
            data=pdf_buf,
            file_name=f"oman_broker_summary_{date.today()}.pdf",
            mime="application/pdf"
        )

    with col_ex3:
        pptx_buf = generate_pptx_report()
        st.download_button(
            "📽️ Download PowerPoint",
            data=pptx_buf,
            file_name=f"oman_broker_summary_{date.today()}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

# -------------------------------------------------------------------
# 12. PAGE: IFRS NOTES
# -------------------------------------------------------------------
elif menu == "📚 IFRS Notes":
    st.title("IFRS & Oman VAT Notes")

    st.subheader("IFRS Standards")
    ifrs_data = pd.DataFrame([
        {"Standard": "IFRS 15", "Relevance": "Commission & fee income (4000/4010/4020) recognised on placement."},
        {"Standard": "IFRS 9", "Relevance": "Cash, commissions receivable (1100), payables (2000/2010); ECL."},
        {"Standard": "IFRS 16", "Relevance": "Office rent (5010) as ROU asset + lease liability."},
        {"Standard": "IAS 1", "Relevance": "Presentation of SoFP and P&L."},
        {"Standard": "IAS 7", "Relevance": "Cash flow statement (operating/investing/financing)."},
        {"Standard": "IAS 21", "Relevance": "FX on reinsurance commissions in foreign currency."},
        {"Standard": "IFRS 17", "Relevance": "Applies to insurers, NOT broker's own commission revenue."},
    ])
    st.dataframe(ifrs_data, use_container_width=True, hide_index=True)

    st.subheader("Oman VAT Treatment")
    st.markdown("""
    - **Commission Income:** Subject to **5% VAT** in Oman as a service supply [citation:30].
    - **Premiums Collected (Client Funds):** Generally **out of scope** for the broker's VAT invoice when acting as a disclosed agent. The insurer is the principal supplier [citation:5][citation:22].
    - **VAT Invoice Requirement:** A VAT-registered broker must issue a tax invoice for commission services. E-invoicing requirements are being phased in (April/October 2027) [citation:29].
    """)

    st.warning("This is a demonstration tool only. Real financial statements require full accruals, ECL, lease accounting, and statutory audit.")