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
    ["Crader Dist. (STIHL)"],
    index=0
)

st.markdown(f"📝 Selected vendor: **{vendor}**")
st.markdown("---")

# =====================
# UNIFIED FILE UPLOADER
# =====================
st.subheader("📂 Upload All Excel Files (Drag & Drop Supported):")

uploaded_files = st.file_uploader(
    "Upload all 3 Excel files here",
    type="xlsx",
    accept_multiple_files=True
)

st.caption("✅ Expected files:")
st.caption("- Merchandise Activity.xlsx (1 year of movement data)")
st.caption("- Merchandise History.xlsx (2 years for ABC classification)")
st.caption("- Merchandise List.xlsx (master inventory list)")

activity_file, history_file, list_file = None, None, None

if uploaded_files:
    for file in uploaded_files:
        fname = file.name.lower()
        if 'activity' in fname:
            activity_file = file
        elif 'history' in fname:
            history_file = file
        elif 'list' in fname:
            list_file = file

    st.markdown(f"🔍 Files detected:")
    st.markdown(f"- {'✅' if activity_file else '❌'} Merchandise Activity file")
    st.markdown(f"- {'✅' if history_file else '❌'} Merchandise History file")
    st.markdown(f"- {'✅' if list_file else '❌'} Merchandise List file")

st.markdown("---")

# =====================
# MAIN PROCESSING LOGIC
# =====================
if st.button("🚀 Run Optimization") and all([activity_file, history_file, list_file]):
    st.info(f"Processing inventory for **{vendor}**... Please wait.")

    # Load and normalize Activity file
    df_activity = pd.read_excel(activity_file)
    df_activity.columns = df_activity.columns.str.strip().str.lower()

    rename_map = {
        'qty_expense': 'qty_expensed',
        'wo_qty used': 'wo_qty_used',
        'part': 'partno'
    }
    df_activity.rename(columns=rename_map, inplace=True)

    # Load and normalize List file
    df_list = pd.read_excel(list_file)
    df_list.columns = df_list.columns.str.strip().str.lower()

    if 'part' in df_list.columns:
        df_list.rename(columns={'part': 'partno'}, inplace=True)
    elif 'part no' in df_list.columns:
        df_list.rename(columns={'part no': 'partno'}, inplace=True)

    if 'partno' not in df_list.columns:
        st.error("❌ Cannot proceed: 'partno' missing in Merchandise List.xlsx after normalization.")
        st.stop()

    # =====================
    # CALCULATIONS
    # =====================
    required_cols = ['qty_sold', 'qty_expensed', 'wo_qty_used', 'partno']
    missing_cols = [col for col in required_cols if col not in df_activity.columns]
    if missing_cols:
        st.error(f"❌ Missing columns in Activity file: {missing_cols}")
        st.stop()

    df_activity['qty_sold_calc'] = df_activity[['qty_sold', 'qty_expensed', 'wo_qty_used']].sum(axis=1)
    df_activity['max_qty'] = df_activity['qty_sold_calc'] * 0.5
    df_activity['min_qty'] = df_activity['max_qty'] * 0.25
    df_activity['re_order_point'] = df_activity['max_qty'] * 0.25
    df_activity['re_order_qty'] = df_activity['max_qty'] - df_activity['min_qty']

    for col in ['qty_sold_calc', 'max_qty', 'min_qty', 're_order_point', 're_order_qty']:
        df_activity[col] = df_activity[col].round(0).astype(int)

    df_merge = pd.merge(
        df_list,
        df_activity[['partno', 'qty_sold_calc', 'max_qty', 'min_qty', 're_order_point', 're_order_qty']],
        on='partno',
        how='left'
    )

    # =====================
    # SKU LOGIC
    # =====================
    def generate_sku(row):
        partno_clean = str(row['partno']).replace(' ', '')
        if row['qty_sold_calc'] > 0:
            return f"S-{row['max_qty']}-{partno_clean}"
        else:
            return f"NS-{partno_clean}"

    df_merge['sku'] = df_merge.apply(generate_sku, axis=1)

    # =====================
    # Vendor column
    # =====================
    df_merge['vendor'] = "Crader Dist. (STIHL)"

    # =====================
    # Remove unwanted columns
    # =====================
    drop_cols = ['upc code', 'last purchase date', 'last count date', 'dated added']
    df_merge = df_merge.drop(columns=[col for col in drop_cols if col in df_merge.columns])

    # Rename Qty_Sold_Calc to Qty_Sold (final output label)
    df_merge.rename(columns={'qty_sold_calc': 'qty_sold'}, inplace=True)

    # =====================
    # OUTPUT TO EXCEL WITH FORMATTING
    # =====================
    output = BytesIO()
    df_merge.to_excel(output, index=False)
    output.seek(0)

    # Apply formatting
    wb = load_workbook(output)
    ws = wb.active

    yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    red = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    grey = PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid')
    green = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')

    # Column index map
    col_idx = {cell.value.lower(): cell.column for cell in ws[1]}
    qty_cols = ['qty_sold', 'max_qty', 'min_qty', 're_order_point', 're_order_qty']

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        sku = row[col_idx['sku'] - 1].value
        qty_sold_val = row[col_idx['qty_sold'] - 1].value

        # Highlight qty fields in yellow
        for qc in qty_cols:
            if qc in col_idx:
                row[col_idx[qc] - 1].fill = yellow

        # Row color logic
        if sku.startswith("NS") and qty_sold_val > 0:
            fill = red
        elif sku.startswith("NS") and qty_sold_val <= 0:
            fill = grey
        elif sku.startswith("S-") and qty_sold_val > 0:
            fill = green
        else:
            fill = grey

        for cell in row:
            cell.fill = fill

    output_final = BytesIO()
    wb.save(output_final)
    output_final.seek(0)

    st.success(f"✅ Optimization complete for **{vendor}**.")
    st.download_button(
        label="📥 Download optimized inventory file",
        data=output_final,
        file_name=f"Optimized_Inventory_{vendor.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Please upload all three required files and select a vendor.")
