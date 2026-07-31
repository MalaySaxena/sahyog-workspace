import streamlit as st
import pandas as pd
import datetime
import os
import openpyxl
import plotly.express as px

# -----------------------------------------------------------------------------
# 1. APP CONFIGURATION & STYLING
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Sahyog Financials Console", layout="wide", page_icon="🏦")

st.markdown("""
    <style>
    .section-header {
        background-color: #1F4E79;
        color: white;
        padding: 10px;
        border-radius: 5px;
        font-weight: bold;
        margin-bottom: 15px;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #1F4E79;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
        margin-bottom: 15px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #f1f3f1;
        border-radius: 4px 4px 0px 0px;
        padding: 8px 16px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1F4E79 !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. AUTO-PATH RESOLUTION & FILE-BASED DATABASE ENGINE
# -----------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else '.'
EXCEL_FILE = os.path.join(SCRIPT_DIR, "Live_Wealth_Manager_Book.xlsx")

SCHEMA_COLUMNS = [
    "Date", "Name", "Amount", "Ucc / APP./FOLIO NO.", "SCHEME", "TYPE", 
    "Top Up", "FROM", "TO", "PAYMENT DETAIL", "BANK DETAIL", "REMARK", 
    "Code", "STNAME", "EXIT LOAD", "Other Remarks", "Entry checked by"
]

def init_excel_ledger():
    """Ensures the permanent Excel file has the Transaction Ledger sheet."""
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        if "4. Transaction Ledger" not in wb.sheetnames:
            ws = wb.create_sheet("4. Transaction Ledger")
            ws.append(SCHEMA_COLUMNS)
            wb.save(EXCEL_FILE)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "4. Transaction Ledger"
        ws.append(SCHEMA_COLUMNS)
        wb.save(EXCEL_FILE)

init_excel_ledger()

@st.cache_data(ttl=2)
def load_all_data():
    """Reads masters and transaction ledger directly from the persistent Excel file."""
    try:
        df_c = pd.read_excel(EXCEL_FILE, sheet_name="2. Client AUM Master")
        df_s = pd.read_excel(EXCEL_FILE, sheet_name="3. Scheme Allocation Master")
        df_l = pd.read_excel(EXCEL_FILE, sheet_name="4. Transaction Ledger")
        
        # Safely strip summary footers
        df_c = df_c[df_c['Code'].notna() & (df_c['Group Name'] != 'Total Book')]
        df_s = df_s[df_s['Code'].notna() & (df_s['Mutual Fund Scheme Name'] != 'Total Book')]
        
        return df_c, df_s, df_l
    except Exception as e:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(columns=SCHEMA_COLUMNS)

df_master_clients, df_master_schemes, df_ledger = load_all_data()

# Dynamic drop-down variables
client_options = sorted(df_master_clients['Investor Name'].dropna().unique().tolist()) if not df_master_clients.empty else []
scheme_options = sorted(df_master_schemes['Mutual Fund Scheme Name'].dropna().unique().tolist()) if not df_master_schemes.empty else []

def append_txn_to_excel(record):
    """Appends a single verified transaction row dictionary back to disk."""
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["4. Transaction Ledger"]
    row_values = [record.get(col, "") for col in SCHEMA_COLUMNS]
    ws.append(row_values)
    wb.save(EXCEL_FILE)
    st.cache_data.clear()

# -----------------------------------------------------------------------------
# 3. INTERFACE NAVIGATION PANEL
# -----------------------------------------------------------------------------
st.sidebar.title("🛠️ Sahyog Financials Workspace")
st.sidebar.caption("Enterprise Consolidated Architecture v7.0")
st.sidebar.markdown("---")
workspace = st.sidebar.radio("Navigate Module Panel:", [
    "📝 Log Transaction",
    "🔍 Advanced Demographics Desk",
    "🗃️ View Saved Master Ledger",
    "📊 Book Summary & Analytics"
])

# -----------------------------------------------------------------------------
# MODULE 1: LOG TRANSACTION FORM DESK
# -----------------------------------------------------------------------------
if workspace == "📝 Log Transaction":
    st.title("Log Transaction")
    st.caption("Transactions entered here are written permanently back to 'Live_Wealth_Manager_Book.xlsx'.")
    st.markdown("---")

    with st.form("persistent_form", clear_on_submit=True):
        st.markdown("<div class='section-header'>1. Core Transaction Details</div>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            txn_date = st.date_input("Date", datetime.date.today())
            client_select = st.selectbox("Name (Select Existing Investor)", ["-- Enter Custom / New Client --"] + client_options)
            client_custom = st.text_input("Name (Type custom name if selected above)")
            final_name = client_custom if client_select == "-- Enter Custom / New Client --" else client_select
        with c2:
            txn_amount = st.number_input("Amount (₹)", min_value=0.0, step=1000.0)
            txn_type = st.selectbox("TYPE", [
                "Purchase", "Redemption", "SIP", "SWP", "STP", "Switch",
                "PMS Subscription", "PMS Top-Up", "PMS Redemption"
            ])
        with c3:
            ucc_folio = st.text_input("Ucc / APP./FOLIO NO.")
            top_up = st.selectbox("Top Up", ["No", "Yes"])

        st.markdown("<div class='section-header'>2. Asset Class & Routing Channels</div>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            scheme_select = st.selectbox("SCHEME (Select Pre-loaded Product/Fund)", ["-- Custom / PMS Entry --"] + scheme_options)
            scheme_custom = st.text_input("SCHEME (Type custom product / strategy name)")
            final_scheme = scheme_custom if scheme_select == "-- Custom / PMS Entry --" else scheme_select
        with c2:
            from_target = st.text_input("FROM (Source scheme for Switch/STP)")
        with c3:
            to_target = st.text_input("TO (Destination target scheme for Switch/STP)")

        st.markdown("<div class='section-header'>3. Banking Rails & Compliance Audit Details</div>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            payment_detail = st.selectbox("PAYMENT DETAIL", ["ACH", "Netbanking", "Cheque", "UPI"])
            bank_detail = st.text_input("BANK DETAIL (Bank Name & A/c details)")
        with c2:
            prod_code = st.selectbox("Code", ["NJ-31241", "NJ-19221", "NJ-43854", "NJ-72021", "ARN-274672"])
            stname_subbroker = st.selectbox("STNAME (Sub Broker Name)", ["NJ", "MFU"])
            exit_load_val = st.text_input("EXIT LOAD (e.g. 1%)")
        with c3:
            remark = st.text_input("REMARK")
            other_remarks = st.text_area("Other Remarks", height=68)
            checker_name = st.text_input("Entry checked by (Auditor Verification Signature)")

        submit = st.form_submit_button("Save Permanently to Excel Database")

        if submit:
            if not final_name or final_name.strip() == "":
                st.error("❌ Save Refused: 'Name' is a required field.")
            elif txn_amount <= 0:
                st.error("❌ Save Refused: 'Amount' must be greater than zero.")
            elif "PMS" in txn_type and "Subscription" in txn_type and txn_amount < 5000000:
                st.error("❌ Regulatory Halt: Fresh PMS mandates cannot fall below the SEBI mandate of ₹50 Lakhs.")
            else:
                new_record = {
                    "Date": str(txn_date),
                    "Name": final_name,
                    "Amount": float(txn_amount),
                    "Ucc / APP./FOLIO NO.": ucc_folio,
                    "SCHEME": final_scheme,
                    "TYPE": txn_type,
                    "Top Up": top_up,
                    "FROM": from_target,
                    "TO": to_target,
                    "PAYMENT DETAIL": payment_detail,
                    "BANK DETAIL": bank_detail,
                    "REMARK": remark,
                    "Code": prod_code,
                    "STNAME": stname_subbroker,
                    "EXIT LOAD": exit_load_val,
                    "Other Remarks": other_remarks,
                    "Entry checked by": checker_name
                }
                append_txn_to_excel(new_record)
                st.success("💾 Transaction successfully saved permanently inside the central workbook!")
                st.rerun()

# -----------------------------------------------------------------------------
# MODULE 2: ADVANCED DEMOGRAPHICS LOOKUPS DESK
# -----------------------------------------------------------------------------
elif workspace == "🔍 Advanced Demographics Desk":
    st.title("Advanced Demographics & Investor Lookups Desk")
    st.caption("Query the fully enriched multi-branch master database using micro-demographic tracking keys.")
    st.markdown("---")
    
    if df_master_clients.empty:
        st.info("No master data loaded yet.")
    else:
        search_filter = st.radio("Select Lookup Key Method:", ["Search by PAN", "Search by UCC", "Search by Distribution Code", "Search by Client Group/Name"], horizontal=True)
        
        filtered_df = df_master_clients.copy()
        
        if search_filter == "Search by PAN":
            pan_input = st.text_input("Enter 10-Digit PAN:").strip().upper()
            if pan_input:
                filtered_df = filtered_df[filtered_df['PAN'].astype(str).str.contains(pan_input, na=False)]
        elif search_filter == "Search by UCC":
            ucc_input = st.text_input("Enter UCC Code:").strip()
            if ucc_input:
                filtered_df = filtered_df[filtered_df['UCC'].astype(str).str.contains(ucc_input, na=False)]
        elif search_filter == "Search by Distribution Code":
            code_select = st.selectbox("Select Branch License Code:", sorted(filtered_df['Code'].unique().tolist()))
            filtered_df = filtered_df[filtered_df['Code'] == code_select]
        elif search_filter == "Search by Client Group/Name":
            name_input = st.text_input("Enter Investor / Group Name keyword:").strip().lower()
            if name_input:
                filtered_df = filtered_df[
                    filtered_df['Investor Name'].astype(str).str.lower().str.contains(name_input, na=False) |
                    filtered_df['Group Name'].astype(str).str.lower().str.contains(name_input, na=False)
                ]
        
        st.markdown(f"#### Matching Accounts Records: ({len(filtered_df)} profiles found)")
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)
        
        if len(filtered_df) == 1:
            st.markdown("---")
            st.markdown("### 📈 Individual Portfolio Breakdown Snapshot")
            row = filtered_df.iloc[0]
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Investor Name", str(row['Investor Name']))
            c2.metric("Total Current AUM", f"₹{row['Current Total AUM']:,.0f}")
            c3.metric("Personal XIRR", f"{row['Total XIRR (%)']:.2f}%")
            c4.metric("Tax Status Slab", str(row['Tax Status']))
            
            # Draw single chart
            labels = ['Equity AUM', 'Debt AUM', 'Hybrid AUM']
            vals = [row['Current Equity AUM'], row['Current Debt AUM'], row['Current Hybrid AUM']]
            if sum(vals) > 0:
                fig_ind = px.pie(names=labels, values=vals, hole=0.4, color_discrete_sequence=px.colors.sequential.Plotly3)
                st.plotly_chart(fig_ind, use_container_width=True)

# -----------------------------------------------------------------------------
# MODULE 3: VIEW SAVED MASTER LEDGER
# -----------------------------------------------------------------------------
elif workspace == "🗃️ View Saved Master Ledger":
    st.title("Permanent Transaction Audit Ledger")
    st.caption("Direct streaming logs retrieved from '4. Transaction Ledger' tab.")
    st.markdown("---")

    if df_ledger.empty:
        st.info("No saved records found in the transaction log sheet yet. Use the input form to create rows.")
    else:
        total_volume = df_ledger["Amount"].sum()
        total_rows = len(df_ledger)
        checked_rows = df_ledger["Entry checked by"].notna().sum()
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Logged Entries", f"{total_rows} Rows")
        with col2:
            st.metric("Gross Logged Volume", f"₹{total_volume:,.2f}")
        with col3:
            st.metric("Audited & Verified", f"{checked_rows} Rows")

        st.markdown("### Saved Transaction Data View")
        st.dataframe(df_ledger, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# MODULE 4: ENRICHED BOOK SUMMARY & ANALYTICS
# -----------------------------------------------------------------------------
elif workspace == "📊 Book Summary & Analytics":
    st.title("Sahyog Financials Advanced Intelligence Desk")
    st.caption("Consolidated macro insights, sales leaderboard channels, and asset house weights.")
    st.markdown("---")
    
    if df_master_clients.empty or df_master_schemes.empty:
        st.error("Historical base sheets are missing from the workbook. Re-verify Excel file integrity.")
    else:
        tab1, tab2, tab3, tab4 = st.tabs([
            "🏠 Executive Overview", 
            "👥 Executive Leaderboard", 
            "💎 Client Segment Matrix", 
            "🏢 Asset House Concentration"
        ])
        
        total_aum_base = df_master_clients['Current Total AUM'].sum()
        equity_base = df_master_clients['Current Equity AUM'].sum()
        hybrid_base = df_master_clients['Current Hybrid AUM'].sum()
        debt_base = df_master_clients['Current Debt AUM'].sum()
        
        with tab1:
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f"<div class='metric-card'><h5>Consolidated Business AUA</h5><h2>₹{total_aum_base:,.2f}</h2></div>", unsafe_allow_html=True)
            with c2:
                st.markdown(f"<div class='metric-card'><h5>Total Monitored Investors</h5><h2>{len(df_master_clients)} Accounts</h2></div>", unsafe_allow_html=True)
            with c3:
                st.markdown(f"<div class='metric-card'><h5>Active Product Matrix</h5><h2>{len(df_master_schemes)} Unique Schemes</h2></div>", unsafe_allow_html=True)
                
            ch_col1, ch_col2 = st.columns(2)
            with ch_col1:
                st.markdown("#### Top 10 High Net Worth Investor Accounts")
                top10_c = df_master_clients.sort_values(by='Current Total AUM', ascending=False).head(10)
                fig1 = px.bar(top10_c, x='Current Total AUM', y='Investor Name', orientation='h', color_discrete_sequence=['#1F4E79'])
                fig1.update_layout(yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig1, use_container_width=True)
            with ch_col2:
                st.markdown("#### Consolidated Asset Class Weights Matrix")
                labels = ['Equity', 'Hybrid', 'Debt']
                values = [equity_base, hybrid_base, debt_base]
                fig2 = px.pie(names=labels, values=values, hole=0.45, color_discrete_sequence=px.colors.sequential.Blues_r)
                st.plotly_chart(fig2, use_container_width=True)
                
        with tab2:
            st.markdown("### Relationship Executive Performance Review")
            partner_perf = df_master_clients.groupby('Partner/Employee').agg(
                Total_Book_AUM=('Current Total AUM', 'sum'),
                Client_Count=('Investor Name', 'count'),
                Avg_Ticket_Size=('Current Total AUM', 'mean')
            ).reset_index().sort_values(by='Total_Book_AUM', ascending=False)
            
            st.dataframe(
                partner_perf.style.format({
                    "Total_Book_AUM": "₹{:,.2f}",
                    "Client_Count": "{:,.0f}",
                    "Avg_Ticket_Size": "₹{:,.2f}"
                }), use_container_width=True, hide_index=True
            )
            
        with tab3:
            st.markdown("### Asset Placement Tiers Matrix")
            def categorize_client(val):
                if val >= 50000000: return "💎 Ultra HNI (> 5 Cr)"
                elif val >= 10000000: return "🥇 HNI (1 Cr - 5 Cr)"
                elif val >= 2500000: return "🥈 Affluent (25L - 1 Cr)"
                else: return "🥉 Retail (< 25L)"
                
            df_master_clients['Client Tier'] = df_master_clients['Current Total AUM'].apply(categorize_client)
            tier_perf = df_master_clients.groupby('Client Tier').agg(
                Combined_AUM=('Current Total AUM', 'sum'),
                Accounts_Count=('Investor Name', 'count')
            ).reset_index()
            tier_perf['AUM Contribution %'] = (tier_perf['Combined_AUM'] / total_aum_base) * 100
            
            st.dataframe(
                tier_perf.style.format({
                    "Combined_AUM": "₹{:,.0f}",
                    "Accounts_Count": "{:,.0f}",
                    "AUM Contribution %": "{:.2f}%"
                }), use_container_width=True, hide_index=True
            )
            
        with tab4:
            st.markdown("### AMC Portfolio Placement Weights")
            amc_grouped = df_master_schemes.groupby('AMC Name')['Current AUM Total (₹)'].sum().reset_index().sort_values(by='Current AUM Total (₹)', ascending=False).head(10)
            amc_grouped['Share %'] = (amc_grouped['Current AUM Total (₹)'] / total_aum_base) * 100
            st.dataframe(
                amc_grouped.style.format({
                    "Current AUM Total (₹)": "₹{:,.0f}",
                    "Share %": "{:.2f}%"
                }), use_container_width=True, hide_index=True
            )
