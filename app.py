import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="📦 Inventory Optimizer", layout="centered")
st.title("📦 Inventory Optimizer")
st.markdown("Upload your Excel files below:")

activity_file = st.file_uploader("Merchandise Activity.xlsx", type="xlsx")
history_file = st.file_uploader("Merchandise History.xlsx", type="xlsx")
list_file = st.file_uploader("Merchandise List.xlsx", type="xlsx")

if st.button("🔧 Process Inventory") and all([activity_file, history_file, list_file]):
    df_activity = pd.read_excel(activity_file)
    df_list = pd.read_excel(list_file)

    df_activity['Qty_Sold_Calc'] = df_activity[['Qty_Sold', 'Qty_Expensed', 'WO_Qty_Used']].sum(axis=1)
    df_activity['MAX_Qty'] = df_activity['Qty_Sold_Calc'] * 0.5
    df_activity['MIN_Qty'] = df_activity['MAX_Qty'] * 0.25
    df_activity['Re_Order_Point'] = df_activity['MAX_Qty'] * 0.25
    df_activity['Re_Order_Qty'] = df_activity['MAX_Qty'] - df_activity['MIN_Qty']

    for col in ['Qty_Sold_Calc', 'MAX_Qty', 'MIN_Qty', 'Re_Order_Point', 'Re_Order_Qty']:
        df_activity[col] = df_activity[col].round(0).astype(int)

    df = pd.merge(df_list, df_activity[['PartNo', 'Qty_Sold_Calc', 'MAX_Qty', 'MIN_Qty', 'Re_Order_Point', 'Re_Order_Qty']], on='PartNo', how='left')

    def format_sku(row):
        part = str(row['PartNo']).replace(" ", "")
        if row['Qty_Sold_Calc'] > 0:
            return f"S-{row['MAX_Qty']}-{part}"
        else:
            return f"NS-{part}"

    df['SKU'] = df.apply(format_sku, axis=1)
    df['Vendor'] = "Crader Dist. (STIHL)"
    df.drop(columns=[c for c in ['UPC Code', 'Last Purchase Date', 'Last Count Date', 'Date Added'] if c in df.columns], inplace=True)
    df.rename(columns={'Qty_Sold_Calc': 'Qty_Sold'}, inplace=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    wb = load_workbook(output)
    ws = wb.active

    yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    red = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    grey = PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid')
    green = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')

    col_idx = {cell.value: cell.column for cell in ws[1]}
    qty_fields = ['Qty_Sold', 'MAX_Qty', 'MIN_Qty', 'Re_Order_Point', 'Re_Order_Qty']

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        sku_val = row[col_idx['SKU'] - 1].value
        qty_val = row[col_idx['Qty_Sold'] - 1].value

        for field in qty_fields:
            row[col_idx[field] - 1].fill = yellow

        if sku_val.startswith("NS") and qty_val > 0:
            fill = red
        elif sku_val.startswith("NS") and qty_val <= 0:
            fill = grey
        elif sku_val.startswith("S-") and qty_val > 0:
            fill = green
        else:
            fill = grey

        for cell in row:
            cell.fill = fill

    output_final = BytesIO()
    wb.save(output_final)
    output_final.seek(0)

    st.success("✅ Processing complete!")
    st.download_button(
        label="📥 Download optimized inventory",
        data=output_final,
        file_name="Optimized_Inventory.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Upload all 3 Excel files and press 'Process Inventory'")
