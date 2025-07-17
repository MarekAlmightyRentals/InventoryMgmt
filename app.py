import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =====================
# PAGE CONFIG
# =====================
st.set_page_config(
    page_title="📦 Almighty Rentals | Inventory Optimizer",
    layout="centered"
)

st.markdown("<h1 style='text-align: center;'>📦 Almighty Rentals Inventory Optimizer</h1>", unsafe_allow_html=True)
st.markdown("---")

# =====================
# VENDOR SELECTOR
# =====================
vendor = st.selectbox(
    "Select Vendor",
    ["Crader Dist. (STIHL)", "Vendor B", "Vendor C"],
    index=0
)

st.markdown(f"📝 Selected vendor: **{vendor}**")
st.markdown("---")

# =====================
# FILE UPLOADS WITH VISIBLE DESCRIPTIONS
# =====================
st.subheader("🔗 Upload Excel Files:")

activity_file = st.file_uploader("Upload Merchandise Activity.xlsx", type="xlsx")
st.caption("🔹 **This file should contain 1 year of item movement data.**")

history_file = st.file_uploader("Upload Merchandise History.xlsx", type="xlsx")
st.caption("🔹 **This file should contain 2 years of historical movement data for ABC classification.**")

list_file = st.file_uploader("Upload Merchandise List.xlsx", type="xlsx")
st.caption("🔹 **This file is the master inventory list to be updated.**")

st.markdown("---")

# =====================
# MAIN PROCESSING LOGIC
# =====================
if st.button("🚀 Run Optimization") and all([activity_file, history_file, list_file]):
    st.info(f"Processing inventory for **{vendor}**... Please wait.")

    # Load and normalize Activity file
    df_activity = pd.read_excel(activity_file)
    df_activity.columns = df_activity.columns.str.strip().str.lower()

    # Rename inconsistent column names for Crader STIHL files
    column_rename_map = {
        'qty_expense': 'qty_expensed',
        'wo_qty used': 'wo_qty_used',
        'part': 'partno'
    }
    df_activity.rename(columns=column_rename_map, inplace=True)

    st.write("🔍 Columns in Merchandise Activity.xlsx after normalization/rename:", df_activity.columns.tolist())

    # Load and normalize List file
    df_list = pd.read_excel(list_file)
    df_list.columns = df_list.columns.str.strip().str.lower()

    st.write("🔍 Columns in Merchandise List.xlsx BEFORE rename:", df_list.columns.tolist())

    # Ensure consistent 'partno' key by checking both 'part' and 'part no'
    if 'part' in df_list.columns:
        df_list.rename(columns={'part': 'partno'}, inplace=True)
    elif 'part no' in df_list.columns:
        df_list.rename(columns={'part no': 'partno'}, inplace=True)

    st.write("🔍 Columns in Merchandise List.xlsx AFTER rename:", df_list.columns.tolist())

    # Validate that 'partno' exists before merge
    if 'partno' not in df_list.columns:
        st.error("❌ Cannot proceed: 'partno' column missing in Merchandise List.xlsx after normalization and rename.")
        st.stop()

    df = None

    if vendor == "Crader Dist. (STIHL)":
        required_cols = ['qty_sold', 'qty_expensed', 'wo_qty_used', 'partno']
        missing_cols = [col for col in required_cols if col not in df_activity.columns]

        if missing_cols:
            st.error(f"❌ Missing columns in Activity file for Crader Dist. (STIHL): {missing_cols}")
            st.stop()

        df_activity['qty_sold_calc'] = df_activity[['qty_sold', 'qty_expensed', 'wo_qty_used']].sum(axis=1)
        df_activity['max_qty'] = df_activity['qty_sold_calc'] * 0.5
        df_activity['min_qty'] = df_activity['max_qty'] * 0.25
        df_activity['re_order_point'] = df_activity['max_qty'] * 0.25
        df_activity['re_order_qty'] = df_activity['max_qty'] - df_activity['min_qty']

        for col in ['qty_sold_calc', 'max_qty', 'min_qty', 're_order_point', 're_order_qty']:
            df_activity[col] = df_activity[col].round(0).astype(int)

        df = pd.merge(
            df_list,
            df_activity[['partno', 'qty_sold_calc', 'max_qty', 'min_qty', 're_order_point', 're_order_qty']],
            on='partno',
            how='left'
        )

        df['vendor'] = "Crader Dist. (STIHL)"

    elif vendor == "Vendor B":
        df = df_list.copy()
        df['vendor'] = "Vendor B"
        st.warning("⚠️ Vendor B logic not yet implemented.")

    elif vendor == "Vendor C":
        df = df_list.copy()
        df['vendor'] = "Vendor C"
        st.warning("⚠️ Vendor C logic not yet implemented.")

    else:
        st.error("Unknown vendor selected.")
        st.stop()

    # =====================
    # FINAL OUTPUT
    # =====================
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    st.success(f"✅ Optimization complete for **{vendor}**.")
    st.download_button(
        label="📥 Download optimized inventory file",
        data=output,
        file_name=f"Optimized_Inventory_{vendor.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Please upload all three required files and select a vendor.")
