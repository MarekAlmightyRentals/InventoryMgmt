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
    layout="wide"
)

st.markdown("<h1 style='text-align: center;'>📦 Almighty Rentals Inventory Optimizer</h1>", unsafe_allow_html=True)
st.markdown("---")

# =====================
# VENDOR SELECTOR
# =====================
vendor = st.selectbox("Select Vendor", ["Crader Dist. (STIHL)"], index=0)
st.markdown(f"📝 Selected vendor: **{vendor}**")
st.markdown("---")

# =====================
# STOCKING RULE SLIDERS
# =====================
st.subheader("⚙️ Custom Stocking Parameters")
max_factor = st.slider("Max stock coverage (% of annual sales)", 10, 100, 50) / 100
min_factor = st.slider("Min stock threshold (% of max)", 5, 50, 25) / 100
reorder_point_factor = st.slider("Reorder point (% of max)", 5, 50, 25) / 100

st.markdown("---")

# =====================
# UNIFIED FILE UPLOADER
# =====================
st.subheader("📂 Upload All Excel Files (Drag & Drop Supported):")
uploaded_files = st.file_uploader("Upload all 3 Excel files here", type="xlsx", accept_multiple_files=True)

st.caption("✅ Expected files:\n- Merchandise Activity.xlsx (1 year)\n- Merchandise History.xlsx (2 years)\n- Merchandise List.xlsx (master list)")

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

    st.markdown("🔍 **Files detected:**")
    st.markdown(f"- {'✅' if activity_file else '❌'} Merchandise Activity")
    st.markdown(f"- {'✅' if history_file else '❌'} Merchandise History")
    st.markdown(f"- {'✅' if list_file else '❌'} Merchandise List")

st.markdown("---")

if st.button("🚀 Run Optimization") and all([activity_file, history_file, list_file]):
    st.info(f"Processing inventory for **{vendor}**... Please wait.")

    # Load files
    df_activity = pd.read_excel(activity_file).rename(columns=lambda x: x.strip().lower())
    df_history = pd.read_excel(history_file).rename(columns=lambda x: x.strip().lower())
    df_list = pd.read_excel(list_file).rename(columns=lambda x: x.strip().lower())

    # Standardize part number
    df_activity.rename(columns={'qty_expense': 'qty_expensed', 'wo_qty used': 'wo_qty_used', 'part': 'partno'}, inplace=True)
    if 'part' in df_list.columns:
        df_list.rename(columns={'part': 'partno'}, inplace=True)
    elif 'part no' in df_list.columns:
        df_list.rename(columns={'part no': 'partno'}, inplace=True)

    if 'partno' not in df_list.columns:
        st.error("❌ 'partno' column missing in Merchandise List.")
        st.stop()

    # Calculate annual sales
    df_activity['qty_sold_calc'] = df_activity[['qty_sold', 'qty_expensed', 'wo_qty_used']].sum(axis=1)
    df_activity['max_qty'] = df_activity['qty_sold_calc'] * max_factor
    df_activity['min_qty'] = df_activity['max_qty'] * min_factor
    df_activity['re_order_point'] = df_activity['max_qty'] * reorder_point_factor
    df_activity['re_order_qty'] = df_activity['max_qty'] - df_activity['min_qty']
    for col in ['qty_sold_calc', 'max_qty', 'min_qty', 're_order_point', 're_order_qty']:
        df_activity[col] = df_activity[col].round(0).astype(int)

    # ABC Classification
    df_history_grouped = df_history.groupby('partno')['qty_sold'].sum().reset_index()
    df_history_grouped['rank'] = df_history_grouped['qty_sold'].rank(ascending=False, method='first')
    df_history_grouped['percentile'] = df_history_grouped['rank'] / len(df_history_grouped)

    def assign_abc(p):
        if p <= 0.2:
            return 'A'
        elif p <= 0.5:
            return 'B'
        else:
            return 'C'

    df_history_grouped['abc_class'] = df_history_grouped['percentile'].apply(assign_abc)

    # Merge all
    df_merge = df_list.merge(df_activity[['partno', 'qty_sold_calc', 'max_qty', 'min_qty', 're_order_point', 're_order_qty']], on='partno', how='left')
    df_merge = df_merge.merge(df_history_grouped[['partno', 'abc_class']], on='partno', how='left')
    df_merge.rename(columns={'qty_sold_calc': 'qty_sold'}, inplace=True)

    # SKU logic (no S-0 safeguard)
    def generate_sku(row):
        partno_clean = str(row['partno']).replace(' ', '')
        if pd.notnull(row['qty_sold']) and row['qty_sold'] > 0:
            if pd.notnull(row['max_qty']) and int(row['max_qty']) > 0:
                return f"S-{int(row['max_qty'])}-{partno_clean}"
            else:
                return f"NS-{partno_clean}"
        else:
            return f"NS-{partno_clean}"

    df_merge['sku'] = df_merge.apply(generate_sku, axis=1)
    df_merge['vendor'] = "Crader Dist. (STIHL)"
    drop_cols = ['upc code', 'last purchase date', 'last count date', 'dated added']
    df_merge.drop(columns=[col for col in drop_cols if col in df_merge.columns], inplace=True)

    # Output to Excel with formatting
    output = BytesIO()
    df_merge.to_excel(output, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    red = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    grey = PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid')
    green = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')

    col_idx = {cell.value.lower(): cell.column for cell in ws[1]}
    qty_cols = ['qty_sold', 'max_qty', 'min_qty', 're_order_point', 're_order_qty']

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        sku = row[col_idx['sku'] - 1].value
        qty_sold_val = row[col_idx['qty_sold'] - 1].value

        # Highlight qty fields in yellow
        for qc in qty_cols:
            if qc in col_idx:
                row[col_idx[qc] - 1].fill = yellow

        # Row coloring
        fill = grey
        if sku:
            if sku.startswith("NS"):
                if qty_sold_val is not None and qty_sold_val > 0:
                    fill = red
            elif sku.startswith("S-"):
                if qty_sold_val is not None and qty_sold_val > 0:
                    fill = green

        for cell in row:
            cell.fill = fill

    output_final = BytesIO()
    wb.save(output_final)
    output_final.seek(0)

    st.success("✅ Optimization complete!")
    st.download_button(
        label="📥 Download optimized inventory file",
        data=output_final,
        file_name=f"Optimized_Inventory_{vendor.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Please upload all three required files and select a vendor.")
