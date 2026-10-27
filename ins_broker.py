import streamlit as st
import pandas as pd
from datetime import date, datetime
import io

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
# Based on CMA Oman Insurance Broker Regulations (Articles 25/26)
# - Client funds must be segregated (Insurance Brokerage Account).
# - Minimum net assets requirement (RO 50,000 for insurance brokerage).

CHART_OF_ACCOUNTS = {
    # ASSETS
    "1000": "Cash - Insurance Brokerage Account (Client Funds)",
    "1010": "Cash - Operating Account (Broker Funds)",
    "1020": "Bank - Operating Account",
    "1100": "Accounts Receivable - Insurance Companies (Commissions)",
    "1200": "Prepaid Expenses",
    "1300": "Fixed Assets (Net)",
    "1400": "Intangible Assets",
    # LIABILITIES
    "2000": "Accounts Payable - Insurance Companies (Premiums)",
    "2010": "Accounts Payable - Clients (Claims)",
    "2100": "Accrued Expenses",
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

# Default opening balances (Omani Rials - OMR)
DEFAULT_BALANCES = {
    "1000": 0.0, "1010": 5000.0, "1020": 15000.0, "1100": 8000.0,
    "1200": 1000.0, "1300": 10000.0, "1400": 0.0,
    "2000": 0.0, "2010": 0.0, "2100": 2000.0, "2200": 0.0,
    "3000": 50000.0, "3100": 5000.0, "3200": 0.0,
    "4000": 0.0, "4010": 0.0, "4020": 0.0, "4100": 0.0,
    "5000": 0.0, "5010": 0.0, "5020": 0.0, "5030": 0.0,
    "5040": 0.0, "5050": 0.0, "5060": 0.0, "5070": 0.0,
}

# -------------------------------------------------------------------
# 3. SESSION STATE INITIALIZATION
# -------------------------------------------------------------------
if "balances" not in st.session_state:
    st.session_state.balances = DEFAULT_BALANCES.copy()

if "transactions" not in st.session_state:
    st.session_state.transactions = []

# -------------------------------------------------------------------
# 4. HELPER FUNCTIONS
# -------------------------------------------------------------------
def get_balance(code):
    return st.session_state.balances.get(code, 0.0)

def update_balance(code, amount):
    if code in st.session_state.balances:
        st.session_state.balances[code] += amount
    else:
        st.error(f"Account code {code} not found.")

def format_omr(val):
    return f"{val:,.3f} OMR"

def load_sample_transactions():
    """Populate session state with illustrative Oman insurance broker transactions."""
    if st.session_state.transactions:
        return  # Don't overwrite existing data

    samples = [
        {
            "Date": date(2026, 1, 15),
            "Description": "Commission earned – Motor fleet policy placed with Dhofar Insurance",
            "Account": "4000",
            "Amount": 1250.500,
            "Type": "Credit"
        },
        {
            "Date": date(2026, 1, 20),
            "Description": "Premium collected from client for property policy (client funds)",
            "Account": "1000",
            "Amount": 8500.000,
            "Type": "Debit"
        },
        {
            "Date": date(2026, 1, 22),
            "Description": "Premium remitted to insurer from brokerage account",
            "Account": "1000",
            "Amount": -8500.000,
            "Type": "Debit"
        },
        {
            "Date": date(2026, 2, 1),
            "Description": "Office rent paid – Muscat premises",
            "Account": "5010",
            "Amount": 900.000,
            "Type": "Debit"
        },
        {
            "Date": date(2026, 2, 10),
            "Description": "Commission earned – Marine cargo policy (reinsurance placed)",
            "Account": "4010",
            "Amount": 2100.750,
            "Type": "Credit"
        },
        {
            "Date": date(2026, 2, 15),
            "Description": "Salaries paid for February",
            "Account": "5000",
            "Amount": 3200.000,
            "Type": "Debit"
        },
        {
            "Date": date(2026, 3, 1),
            "Description": "Bank charges – operating account",
            "Account": "5060",
            "Amount": 45.500,
            "Type": "Debit"
        },
        {
            "Date": date(2026, 3, 10),
            "Description": "Commission receivable from insurer (motor portfolio)",
            "Account": "1100",
            "Amount": 1800.000,
            "Type": "Debit"
        },
        {
            "Date": date(2026, 3, 15),
            "Description": "Audit fees paid – annual statutory audit",
            "Account": "5030",
            "Amount": 1200.000,
            "Type": "Debit"
        },
    ]

    for txn in samples:
        code = txn["Account"]
        amt = txn["Amount"]
        if code in st.session_state.balances:
            st.session_state.balances[code] += amt
        st.session_state.transactions.append(txn)

# Load sample data once
load_sample_transactions()

# -------------------------------------------------------------------
# 5. APP LAYOUT & NAVIGATION
# -------------------------------------------------------------------
st.sidebar.title("🏦 Insurance Broker FS")
st.sidebar.caption("Oman CMA Compliant Template")

menu = st.sidebar.radio(
    "Navigation",
    ["📊 Dashboard", "📝 Chart of Accounts", "📥 Data Entry", "📑 Reports", "📚 IFRS Notes"]
)

# -------------------------------------------------------------------
# 6. PAGE: DASHBOARD
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

    st.subheader("Compliance Check (Article 26)")
    min_capital = 50000.0
    if net_assets >= min_capital:
        st.success(
            f"✅ Solvency Requirement Met: Net Assets ({format_omr(net_assets)}) "
            f"exceed minimum {format_omr(min_capital)}"
        )
    else:
        st.error(
            f"⚠️ Solvency Requirement Not Met: Net Assets ({format_omr(net_assets)}) "
            f"are below minimum {format_omr(min_capital)}"
        )

    st.info(
        "Note: Client funds (Account 1000) are excluded from broker's net assets per CMA regulations."
    )

    st.subheader("Snapshot of Balances")
    snapshot = pd.DataFrame([
        {"Code": k, "Account": CHART_OF_ACCOUNTS[k], "Balance (OMR)": v}
        for k, v in st.session_state.balances.items()
    ])
    st.dataframe(snapshot, use_container_width=True, height=300)

# -------------------------------------------------------------------
# 7. PAGE: CHART OF ACCOUNTS
# -------------------------------------------------------------------
elif menu == "📝 Chart of Accounts":
    st.title("Chart of Accounts (Oman Context)")
    st.markdown("Standardized accounts for an Insurance Brokerage in the Sultanate of Oman.")

    df_accounts = pd.DataFrame([
        {"Code": code, "Account Name": name, "Current Balance": get_balance(code)}
        for code, name in CHART_OF_ACCOUNTS.items()
    ])

    st.dataframe(df_accounts, use_container_width=True)

# -------------------------------------------------------------------
# 8. PAGE: DATA ENTRY
# -------------------------------------------------------------------
elif menu == "📥 Data Entry":
    st.title("Record Transactions")

    with st.form("transaction_form"):
        col1, col2 = st.columns(2)
        with col1:
            txn_date = st.date_input("Date", date.today())
            description = st.text_input("Description")
        with col2:
            account_code = st.selectbox(
                "Account",
                options=list(CHART_OF_ACCOUNTS.keys()),
                format_func=lambda x: f"{x} - {CHART_OF_ACCOUNTS[x]}"
            )
            amount = st.number_input("Amount (OMR)", min_value=0.0, step=10.0)
            txn_type = st.radio(
                "Type",
                ["Debit (Increase Asset/Expense)", "Credit (Increase Liability/Revenue)"]
            )

        submitted = st.form_submit_button("Post Transaction")

        if submitted:
            if not description:
                st.error("Please enter a description.")
            elif amount <= 0:
                st.error("Amount must be positive.")
            else:
                signed_amount = amount if txn_type.startswith("Debit") else amount
                update_balance(account_code, signed_amount)

                st.session_state.transactions.append({
                    "Date": txn_date,
                    "Description": description,
                    "Account": account_code,
                    "Amount": signed_amount,
                    "Type": txn_type
                })
                st.success(f"Transaction posted to {account_code}.")

    if st.session_state.transactions:
        st.subheader("Recent Transactions")
        df_txns = pd.DataFrame(st.session_state.transactions)
        st.dataframe(df_txns, use_container_width=True)

# -------------------------------------------------------------------
# 9. PAGE: REPORTS (FS / PL / Cash Flow)
# -------------------------------------------------------------------
elif menu == "📑 Reports":
    st.title("Financial Reports")

    # --- 9a. Statement of Financial Position ---
    st.header("Statement of Financial Position")
    st.caption("As of " + str(date.today()))

    assets = {k: v for k, v in st.session_state.balances.items() if k.startswith("1")}
    total_assets = sum(assets.values())

    liabilities = {k: v for k, v in st.session_state.balances.items() if k.startswith("2")}
    total_liabilities = sum(liabilities.values())

    equity = {k: v for k, v in st.session_state.balances.items() if k.startswith("3")}
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
        st.markdown(f"Total Liabilities: {format_omr(total_liabilities)}")

        st.markdown("**Equity**")
        for code, val in equity.items():
            st.write(f"{CHART_OF_ACCOUNTS[code]}: **{format_omr(val)}**")
        st.markdown(f"Total Equity: {format_omr(total_equity)}")
        st.markdown(f"### Total Liab. & Equity: {format_omr(total_liabilities + total_equity)}")

    st.divider()

    # --- 9b. Profit & Loss ---
    st.header("Profit & Loss Statement")
    st.caption("For the period ending " + str(date.today()))

    revenue = {k: v for k, v in st.session_state.balances.items() if k.startswith("4")}
    expenses = {k: v for k, v in st.session_state.balances.items() if k.startswith("5")}

    total_revenue = sum(revenue.values())
    total_expenses = sum(expenses.values())
    net_profit = total_revenue - total_expenses

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Revenue")
        for code, val in revenue.items():
            st.write(f"{CHART_OF_ACCOUNTS[code]}: **{format_omr(val)}**")
        st.markdown(f"### Total Revenue: {format_omr(total_revenue)}")

    with col2:
        st.subheader("Expenses")
        for code, val in expenses.items():
            st.write(f"{CHART_OF_ACCOUNTS[code]}: **{format_omr(val)}**")
        st.markdown(f"### Total Expenses: {format_omr(total_expenses)}")

    st.markdown(f"## Net Profit / (Loss): {format_omr(net_profit)}")

    st.divider()

    # --- 9c. Cash Flow Statement (Simplified) ---
    st.header("Cash Flow Statement (Simplified)")
    st.caption("Indirect Method approximation")

    cash_accounts = ["1000", "1010", "1020"]
    total_cash = sum(get_balance(c) for c in cash_accounts)

    st.write(f"**Cash and Cash Equivalents (End of Period):** {format_omr(total_cash)}")
    st.info(
        "Note: In a production environment, this would reconcile with opening cash balances "
        "and detailed cash flow movements (Operating, Investing, Financing)."
    )

    # --- 9d. Export ---
    st.subheader("Export")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_bal = pd.DataFrame([
            {"Code": k, "Account": CHART_OF_ACCOUNTS[k], "Balance": v}
            for k, v in st.session_state.balances.items()
        ])
        df_bal.to_excel(writer, sheet_name='Trial Balance', index=False)

        if st.session_state.transactions:
            pd.DataFrame(st.session_state.transactions).to_excel(
                writer, sheet_name='Transactions', index=False
            )

    st.download_button(
        label="📥 Download Financials (Excel)",
        data=output.getvalue(),
        file_name=f"oman_broker_financials_{date.today()}.xlsx",
        mime="application/vnd.ms-excel"
    )

# -------------------------------------------------------------------
# 10. PAGE: IFRS NOTES
# -------------------------------------------------------------------
elif menu == "📚 IFRS Notes":
    st.title("IFRS Standards Applied to an Oman Insurance Broker")
    st.markdown(
        "An insurance **broker** is not an insurer. It does not issue insurance contracts, "
        "so it does **not** apply IFRS 17 to its own revenue. Below is the mapping of the "
        "most relevant IFRS standards to this demo."
    )

    ifrs_table = pd.DataFrame([
        {
            "IFRS / IAS": "IFRS 15",
            "Title": "Revenue from Contracts with Customers",
            "Relevance": "Commission & fee income (Accounts 4000/4010/4020) recognised on placement.",
            "Demo Account": "4000, 4010, 4020"
        },
        {
            "IFRS / IAS": "IFRS 9",
            "Title": "Financial Instruments",
            "Relevance": "Cash, commissions receivable (1100), payables (2000/2010); ECL on receivables.",
            "Demo Account": "1100, 2000, 2010"
        },
        {
            "IFRS / IAS": "IFRS 16",
            "Title": "Leases",
            "Relevance": "Office rent recognised as ROU asset + lease liability (simplified as expense in demo).",
            "Demo Account": "5010"
        },
        {
            "IFRS / IAS": "IAS 1",
            "Title": "Presentation of Financial Statements",
            "Relevance": "Structure of SoFP and P&L; current/non-current classification.",
            "Demo Account": "Reports tab"
        },
        {
            "IFRS / IAS": "IAS 7",
            "Title": "Statement of Cash Flows",
            "Relevance": "Operating / investing / financing classification; demo shows ending cash only.",
            "Demo Account": "1000, 1010, 1020"
        },
        {
            "IFRS / IAS": "IAS 21",
            "Title": "Effects of Changes in Foreign Exchange Rates",
            "Relevance": "FX on reinsurance commissions in USD / foreign currency.",
            "Demo Account": "4010, 1100"
        },
        {
            "IFRS / IAS": "IFRS 17",
            "Title": "Insurance Contracts",
            "Relevance": "Applies to insurers, NOT to the broker's own commission revenue.",
            "Demo Account": "—"
        },
    ])
    st.dataframe(ifrs_table, use_container_width=True, hide_index=True)

    st.subheader("Key Practical Points for Oman")
    st.markdown(
        """
        - **CMA Article 25 – Segregated client funds:** Premiums held for clients sit in the
          *Insurance Brokerage Account* (Account 1000). Under CMA rules these are **client
          monies**, not the broker's own cash, and should be disclosed separately rather than
          included in the broker's "Cash and cash equivalents".
        - **CMA Article 26 – Minimum net assets:** The broker must maintain minimum net assets
          of **RO 50,000** (insurance brokerage) or **RO 75,000** (combined insurance +
          reinsurance brokerage). This is a *regulatory* test, not an IFRS requirement.
        - **IFRS vs CMA:** IFRS governs the *presentation and measurement* of the financial
          statements; the CMA regulations govern *solvency, client-money segregation, and
          licensing*. Both must be satisfied simultaneously.
        - **Audit:** Statutory annual audit is required by the CMA; the auditor will test
          compliance with both IFRS and CMA rules.
        """
    )

    st.subheader("Disclaimer")
    st.warning(
        "This is a **demonstration / training tool** only. It is not a substitute for a "
        "full accounting system, statutory audit, or professional advice. Real IFRS financial "
        "statements require opening balances, accruals, ECL, lease accounting, FX revaluation, "
        "deferred tax (where applicable), and detailed note disclosures."
    )