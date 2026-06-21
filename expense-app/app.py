import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import json

# ─── Page Config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AB Group – Expense Analyser",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── Category Mapping ─────────────────────────────────────────────────────────
CATEGORY_MAP = {
    # Food & Kitchen
    "Grocery": "Food & Kitchen",
    "Vegetable": "Food & Kitchen",
    "Fruit": "Food & Kitchen",
    "Nonveg": "Food & Kitchen",
    "Milk & Products": "Food & Kitchen",
    "Colddrink": "Food & Kitchen",
    "Icecream": "Food & Kitchen",
    "Gas": "Food & Kitchen",
    "Packing": "Food & Kitchen",
    # Staff & HR
    "Staff Salary": "Staff & HR",
    "PSTL Salary": "Staff & HR",
    "Salary Advance": "Staff & HR",
    "Professional Fees": "Staff & HR",
    "Uniform & Shoes Expenses": "Staff & HR",
    "Medical": "Staff & HR",
    "Chetan": "Staff & HR",
    "Paresh": "Staff & HR",
    "Prashant": "Staff & HR",
    "Vikram Sir": "Staff & HR",
    "Child": "Staff & HR",
    # Utilities
    "Electricity": "Utilities",
    "Internet": "Utilities",
    "Mobile Exps": "Utilities",
    "Generator Rent": "Utilities",
    "Fuel": "Utilities",
    "PSTL-Fuel": "Utilities",
    "Fuel Used for PSTL ": "Utilities",
    "Transport": "Utilities",
    "Conveyence": "Utilities",
    # Maintenance & Repairs
    "Repairs Civil": "Maintenance & Repairs",
    "Repairs Furniture": "Maintenance & Repairs",
    "Repairs Computer": "Maintenance & Repairs",
    "Hardware": "Maintenance & Repairs",
    "Electrical": "Maintenance & Repairs",
    "Civil": "Maintenance & Repairs",
    "Paint": "Maintenance & Repairs",
    "Material": "Maintenance & Repairs",
    "Labour Charges": "Maintenance & Repairs",
    "Fabrication": "Maintenance & Repairs",
    # Assets & Equipment
    "Furniture": "Assets & Equipment",
    "PSTL-Furniture": "Assets & Equipment",
    "Machinery": "Assets & Equipment",
    "Utensils": "Assets & Equipment",
    "Cloth & Mattress": "Assets & Equipment",
    "Sports Material": "Assets & Equipment",
    "PSTL-Sports Material": "Assets & Equipment",
    # Guest & Operations
    "VIP Guest Expenses": "Guest & Operations",
    "Decoration": "Guest & Operations",
    "Anciliary Services": "Guest & Operations",
    "Housekeeping": "Guest & Operations",
    "Gardening": "Guest & Operations",
    # Marketing & Admin
    "Advertising Expenses": "Marketing & Admin",
    "PSTL-Advertising Expenses": "Marketing & Admin",
    "Printing and Stationery": "Marketing & Admin",
    "Comission Paid": "Marketing & Admin",
}

MAIN_CATEGORIES = list(dict.fromkeys(CATEGORY_MAP.values())) + ["Other"]
ALL_EXPENSE_HEADS = sorted(CATEGORY_MAP.keys())

# ─── Helpers ──────────────────────────────────────────────────────────────────
def get_main_category(expense_head):
    if not expense_head or str(expense_head).strip() in ("", "nan", "0", "1", "2"):
        return None
    return CATEGORY_MAP.get(str(expense_head).strip(), None)

def fmt_inr(amount):
    """Format number as Indian Rupee with lakhs formatting."""
    try:
        amount = float(amount)
        if amount >= 100000:
            return f"₹{amount/100000:.2f}L"
        elif amount >= 1000:
            return f"₹{amount/1000:.1f}K"
        else:
            return f"₹{amount:,.0f}"
    except:
        return "₹0"

def is_upi_sheet(df_raw):
    """Check if a sheet looks like a UPI bank statement."""
    for i in range(min(5, len(df_raw))):
        row_vals = [str(v) for v in df_raw.iloc[i].values]
        row_str = " ".join(row_vals)
        if "Expense Head" in row_str and "Debit" in row_str:
            return True
    return False

def parse_upi_sheet(df_raw, sheet_name):
    """Parse a UPI statement sheet into clean records."""
    header_idx = -1
    for i in range(min(5, len(df_raw))):
        row_vals = [str(v) for v in df_raw.iloc[i].values]
        row_str = " ".join(row_vals)
        if "Expense Head" in row_str and "Debit" in row_str:
            header_idx = i
            break

    if header_idx == -1:
        return pd.DataFrame(), []

    headers = [str(v).strip() for v in df_raw.iloc[header_idx].values]
    data = df_raw.iloc[header_idx + 1:].copy()
    data.columns = headers

    # Find key columns flexibly
    date_col   = next((c for c in headers if "date" in c.lower()), None)
    debit_col  = next((c for c in headers if c.lower() == "debit"), None)
    desc_col   = next((c for c in headers if "transaction" in c.lower() or "details" in c.lower()), None)
    head_col   = next((c for c in headers if "expense head" in c.lower()), None)
    acct_col   = next((c for c in headers if "statement" in c.lower()), None)

    if not debit_col or not head_col:
        return pd.DataFrame(), []

    records = []
    uncategorized = []

    for _, row in data.iterrows():
        raw_debit = str(row.get(debit_col, "")).replace(",", "").strip()
        try:
            debit = float(raw_debit)
        except ValueError:
            continue
        if debit <= 0:
            continue

        expense_head = str(row.get(head_col, "")).strip()
        desc         = str(row.get(desc_col, "")).strip() if desc_col else ""
        date         = str(row.get(date_col, "")).strip() if date_col else ""
        account      = str(row.get(acct_col, "")).strip() if acct_col else ""

        # Clean date if it's a pandas Timestamp artifact
        if "00:00:00" in date:
            date = date.split(" ")[0]

        main_cat = get_main_category(expense_head)
        is_missing = (not expense_head or expense_head in ("nan", "0", "1", "2", ""))

        rec = {
            "Date": date,
            "Description": desc[:80],
            "Amount": debit,
            "Expense Head": expense_head if not is_missing else "",
            "Main Category": main_cat or "Other",
            "Source": sheet_name,
            "Type": "UPI/Bank",
            "Needs Category": is_missing,
        }
        records.append(rec)
        if is_missing:
            uncategorized.append(rec.copy())

    return pd.DataFrame(records), uncategorized

# ─── Session State Init ───────────────────────────────────────────────────────
if "all_records" not in st.session_state:
    st.session_state.all_records = pd.DataFrame()
if "cash_entries" not in st.session_state:
    st.session_state.cash_entries = []
if "uncat_queue" not in st.session_state:
    st.session_state.uncat_queue = []
if "sheets_loaded" not in st.session_state:
    st.session_state.sheets_loaded = []

# ─── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1E3A5F 0%, #2563EB 100%);
        padding: 1.2rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 1rem 1.2rem;
        border-left: 4px solid;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
    .uncat-banner {
        background: #FFFBEB;
        border: 1px solid #FCD34D;
        border-radius: 10px;
        padding: 1rem 1.2rem;
        margin-bottom: 1rem;
    }
    div[data-testid="stMetric"] {
        background: white;
        border-radius: 10px;
        padding: 0.8rem 1rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.07);
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 6px 18px;
        font-weight: 600;
    }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h2 style="margin:0; font-size:1.5rem;">📊 AB Group — Expense Analyser</h2>
    <p style="margin:0.2rem 0 0; opacity:0.75; font-size:0.85rem;">FY 2025-26 · Internal Finance Tool</p>
</div>
""", unsafe_allow_html=True)

# ─── File Upload ──────────────────────────────────────────────────────────────
with st.container():
    uploaded_files = st.file_uploader(
        "📂 Upload UPI Bank Statement(s)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="Upload one or more Excel files. Each sheet named like 'Oct 25', 'Nov 25', etc. will be detected automatically."
    )

    if uploaded_files:
        new_sheets = []
        all_records_list = []
        all_uncat = []

        for f in uploaded_files:
            try:
                xl = pd.ExcelFile(f)
                for sheet in xl.sheet_names:
                    df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
                    if is_upi_sheet(df_raw):
                        df_parsed, uncat = parse_upi_sheet(df_raw, sheet)
                        if not df_parsed.empty:
                            new_sheets.append(sheet)
                            all_records_list.append(df_parsed)
                            all_uncat.extend(uncat)
            except Exception as e:
                st.error(f"Could not read {f.name}: {e}")

        if all_records_list:
            st.session_state.all_records = pd.concat(all_records_list, ignore_index=True)
            st.session_state.sheets_loaded = new_sheets
            # Only queue truly missing expense heads
            st.session_state.uncat_queue = [u for u in all_uncat if not u.get("Expense Head")]
            st.success(f"✅ Loaded **{len(new_sheets)} sheet(s)**: {', '.join(new_sheets)}  |  **{len(st.session_state.all_records)}** transactions")
        elif uploaded_files:
            st.warning("No UPI statement sheets found. Make sure your file has columns: Date, Transaction Details, Debit, Expense Head.")

# ─── Uncategorized Queue ──────────────────────────────────────────────────────
if st.session_state.uncat_queue:
    remaining = st.session_state.uncat_queue
    st.markdown(f"""
    <div class="uncat-banner">
        ⚠️ <strong>{len(remaining)} transaction(s)</strong> have no Expense Head — please categorise them below.
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"🏷️ Categorise {len(remaining)} transaction(s)", expanded=True):
        resolved_indices = []
        for idx, item in enumerate(remaining[:10]):
            col1, col2, col3, col4 = st.columns([3, 2, 2, 1])
            with col1:
                st.caption(f"**{item['Description'][:60]}**  \n{item['Date']} · ₹{item['Amount']:,.0f}")
            with col2:
                head = st.selectbox("Expense Head", ["— Select —"] + ALL_EXPENSE_HEADS, key=f"uq_head_{idx}")
            with col3:
                auto_cat = get_main_category(head) if head != "— Select —" else None
                cat = st.selectbox("Main Category", ["— Auto —"] + MAIN_CATEGORIES,
                                   index=0 if not auto_cat else MAIN_CATEGORIES.index(auto_cat) + 1,
                                   key=f"uq_cat_{idx}")
            with col4:
                st.write("")
                if st.button("Save", key=f"uq_save_{idx}"):
                    if head != "— Select —":
                        final_cat = cat if cat != "— Auto —" else (auto_cat or "Other")
                        # Update the main records
                        mask = (
                            (st.session_state.all_records["Description"] == item["Description"]) &
                            (st.session_state.all_records["Date"] == item["Date"]) &
                            (st.session_state.all_records["Amount"] == item["Amount"])
                        )
                        st.session_state.all_records.loc[mask, "Expense Head"] = head
                        st.session_state.all_records.loc[mask, "Main Category"] = final_cat
                        st.session_state.all_records.loc[mask, "Needs Category"] = False
                        resolved_indices.append(idx)
                        st.rerun()

        if len(remaining) > 10:
            st.caption(f"…and {len(remaining)-10} more. Save the above to proceed.")

        # Remove resolved
        st.session_state.uncat_queue = [u for i, u in enumerate(remaining) if i not in resolved_indices]

# ─── Main Content ─────────────────────────────────────────────────────────────
df = st.session_state.all_records.copy()

# Add cash entries
cash_df = pd.DataFrame(st.session_state.cash_entries) if st.session_state.cash_entries else pd.DataFrame()
if not cash_df.empty:
    combined = pd.concat([df, cash_df], ignore_index=True)
else:
    combined = df.copy()

tab1, tab2, tab3 = st.tabs(["📊 Analysis", "📋 All Records", "💵 Cash Entry"])

# ══════════════════════════════════════════════════════════
# TAB 1 — ANALYSIS
# ══════════════════════════════════════════════════════════
with tab1:
    if combined.empty:
        st.info("👆 Upload a UPI bank statement above to see spending analysis.")
    else:
        total     = combined["Amount"].sum()
        upi_total = df["Amount"].sum() if not df.empty else 0
        cash_total = cash_df["Amount"].sum() if not cash_df.empty else 0
        num_txns  = len(combined)

        # ── Month filter
        months = ["All Months"] + (
            sorted(combined["Source"].unique().tolist()) if "Source" in combined.columns else []
        )
        selected_month = st.selectbox("Filter by Month", months, key="month_filter")
        view_df = combined if selected_month == "All Months" else combined[combined["Source"] == selected_month]

        # ── KPI row
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("💰 Total Expenses", fmt_inr(view_df["Amount"].sum()))
        c2.metric("🏦 UPI / Bank", fmt_inr(view_df[view_df["Type"] == "UPI/Bank"]["Amount"].sum()))
        c3.metric("💵 Cash", fmt_inr(view_df[view_df["Type"] == "Cash"]["Amount"].sum()) if "Type" in view_df else "₹0")
        c4.metric("🧾 Transactions", len(view_df))

        st.divider()

        # ── Aggregations
        by_main = view_df.groupby("Main Category")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
        by_head = view_df.groupby("Expense Head")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
        by_head = by_head[by_head["Expense Head"].str.strip() != ""]

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("By Main Category")
            fig_pie = px.pie(
                by_main, values="Amount", names="Main Category",
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Bold,
            )
            fig_pie.update_traces(textposition="inside", textinfo="percent+label")
            fig_pie.update_layout(margin=dict(t=10, b=10, l=10, r=10), showlegend=False, height=320)
            st.plotly_chart(fig_pie, use_container_width=True)

        with col2:
            st.subheader("Top Expense Heads")
            fig_bar = px.bar(
                by_head.head(12), x="Amount", y="Expense Head",
                orientation="h",
                color="Amount",
                color_continuous_scale="Blues",
                text=by_head.head(12)["Amount"].apply(fmt_inr),
            )
            fig_bar.update_traces(textposition="outside")
            fig_bar.update_layout(
                margin=dict(t=10, b=10, l=10, r=60),
                yaxis=dict(autorange="reversed"),
                coloraxis_showscale=False,
                height=320,
                xaxis_title="",
                yaxis_title="",
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        # ── Category Breakdown Table
        st.subheader("📋 Category Breakdown")
        total_amt = view_df["Amount"].sum()

        breakdown_rows = []
        for _, row in by_main.iterrows():
            cat = row["Main Category"]
            amt = row["Amount"]
            sub_heads = view_df[view_df["Main Category"] == cat]["Expense Head"].value_counts()
            sub_heads = [h for h in sub_heads.index if h.strip()]
            breakdown_rows.append({
                "Main Category": cat,
                "Amount (₹)": f"₹{amt:,.0f}",
                "% of Total": f"{(amt/total_amt*100):.1f}%",
                "Sub-categories": ", ".join(sub_heads[:6]) + ("…" if len(sub_heads) > 6 else ""),
            })

        st.dataframe(
            pd.DataFrame(breakdown_rows),
            use_container_width=True,
            hide_index=True,
        )

        # ── Monthly trend (if multiple months)
        if len(combined["Source"].unique()) > 1:
            st.subheader("📅 Month-wise Trend")
            monthly = combined.groupby(["Source", "Main Category"])["Amount"].sum().reset_index()
            fig_trend = px.bar(
                monthly, x="Source", y="Amount", color="Main Category",
                barmode="stack",
                color_discrete_sequence=px.colors.qualitative.Bold,
            )
            fig_trend.update_layout(
                margin=dict(t=10, b=10),
                xaxis_title="Month",
                yaxis_title="Amount (₹)",
                legend_title="Category",
                height=350,
            )
            st.plotly_chart(fig_trend, use_container_width=True)

# ══════════════════════════════════════════════════════════
# TAB 2 — ALL RECORDS
# ══════════════════════════════════════════════════════════
with tab2:
    if combined.empty:
        st.info("No records yet. Upload a file or add cash expenses.")
    else:
        search = st.text_input("🔍 Search description", placeholder="e.g. salary, fuel…")
        cat_filter = st.multiselect("Filter by Main Category", options=sorted(combined["Main Category"].unique()))

        filtered = combined.copy()
        if search:
            filtered = filtered[filtered["Description"].str.contains(search, case=False, na=False)]
        if cat_filter:
            filtered = filtered[filtered["Main Category"].isin(cat_filter)]

        st.caption(f"Showing **{len(filtered)}** of {len(combined)} records · Total: ₹{filtered['Amount'].sum():,.0f}")

        show_cols = ["Date", "Description", "Amount", "Expense Head", "Main Category", "Source", "Type"]
        show_cols = [c for c in show_cols if c in filtered.columns]
        st.dataframe(
            filtered[show_cols].sort_values("Amount", ascending=False),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Amount": st.column_config.NumberColumn("Amount (₹)", format="₹%d"),
            }
        )

# ══════════════════════════════════════════════════════════
# TAB 3 — CASH ENTRY
# ══════════════════════════════════════════════════════════
with tab3:
    col1, col2 = st.columns([1, 1], gap="large")

    with col1:
        st.subheader("➕ Add Cash Expense")
        with st.form("cash_form", clear_on_submit=True):
            title   = st.text_input("Title / Description *", placeholder="e.g. Vegetables from market")
            amount  = st.number_input("Amount (₹) *", min_value=0.0, step=10.0, format="%.2f")
            head    = st.selectbox("Expense Head *", ["— Select —"] + ALL_EXPENSE_HEADS)
            auto_mc = get_main_category(head) if head != "— Select —" else None
            mc      = st.selectbox(
                "Main Category",
                MAIN_CATEGORIES,
                index=MAIN_CATEGORIES.index(auto_mc) if auto_mc and auto_mc in MAIN_CATEGORIES else len(MAIN_CATEGORIES)-1,
            )
            custom_head = st.text_input("Or enter custom expense head", placeholder="Leave blank if selected above")
            submitted = st.form_submit_button("✅ Add Cash Expense", use_container_width=True, type="primary")

            if submitted:
                final_head = custom_head.strip() if custom_head.strip() else head
                if not title:
                    st.error("Title is required.")
                elif amount <= 0:
                    st.error("Amount must be greater than 0.")
                elif final_head == "— Select —":
                    st.error("Please select or enter an expense head.")
                else:
                    final_mc = mc if mc else (get_main_category(final_head) or "Other")
                    st.session_state.cash_entries.append({
                        "Date": pd.Timestamp.today().strftime("%d/%m/%y"),
                        "Description": title,
                        "Amount": float(amount),
                        "Expense Head": final_head,
                        "Main Category": final_mc,
                        "Source": "Cash",
                        "Type": "Cash",
                        "Needs Category": False,
                    })
                    st.success(f"✅ Added: {title} — ₹{amount:,.0f} under {final_head}")
                    st.rerun()

    with col2:
        st.subheader(f"💵 Cash Entries ({len(st.session_state.cash_entries)})")
        if not st.session_state.cash_entries:
            st.info("No cash entries yet.")
        else:
            cash_view = pd.DataFrame(st.session_state.cash_entries)[
                ["Date", "Description", "Amount", "Expense Head", "Main Category"]
            ]
            st.dataframe(cash_view, use_container_width=True, hide_index=True,
                         column_config={"Amount": st.column_config.NumberColumn("Amount (₹)", format="₹%d")})
            st.metric("Total Cash Expenses", f"₹{cash_view['Amount'].sum():,.0f}")

            if st.button("🗑️ Clear all cash entries"):
                st.session_state.cash_entries = []
                st.rerun()

# ─── Footer ───────────────────────────────────────────────────────────────────
st.divider()
st.caption("AB Group Internal Finance Tool · Phase 1 · Built with Streamlit")
