import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="📦 Almighty Rentals | Inventory Optimizer", layout="centered")

st.markdown("<h1 style='text-align: center;'>📦 Almighty Rentals Inventory Optimizer</h1>", unsafe_allow_html=True)
st.markdown("---")

# 🔹 Vendor dropdown selector
vendor = st.selectbox(
    "Select Vendor",
    ["Crader Dist. (STIHL)", "Vendor B", "Vendor C"],
    index=0
)

st.markdown(f"📝 Selected vendor: **{vendor}**")

st.markdown("---")

# 🔹 File upload section with descriptions
activity_file = st.file_uploader(
    "Upload Merchandise Activity.xlsx",
    type="xlsx",
    help="This file should contain 1 year of item movement data."
)

history_file = st.file_uploader(
    "Upload Merchandise History.xlsx",
    type="xlsx",
    help="This file should contain 2 years of historical movement data for ABC classification."
)

list_file = st.file_uploader(
    "Upload Merchandise List.xlsx",
    type="xlsx",
    help="This file is the master inventory list to be updated."
)

st.markdown("---")

if st.button("🚀 Run Optimization") and all([activity_file, history_file, list_file]):
    st.info(f"Processing inventory for **{vendor}**... Please wait.")

    # 🔹 Load common files
    df_activity = pd.read_excel(activity_file)
    df_list = pd.read_excel(list_file)

    # 🔧 Branching: handle different vendor logic
    if vendor == "Crader Dist. (STIHL)":
        # 📝 Crader-specific logic (your original implementation)
        df_activity['Qty_Sold_Calc'] = df_activity[['Qty_Sold', 'Qty_Expensed', 'WO_Qty_Used']].sum(axis=1)
        df_activity['MAX_Qty'] = df_activity['Qty_Sold_Calc'] * 0.5
        df_activity['MIN_Qty'] = df_activity['MAX_Qty'] * 0.25
        df_activity['Re_Order_Point'] = df_activity['MAX_Qty'] * 0.25
        df_activity['Re_Order_Qty'] = df_activity['MAX_Qty'] - df_activity['MIN_Qty']
        for col in ['Qty_Sold_Calc', 'MAX_Qty', 'MIN_Qty', 'Re_Order_Point', 'Re_Order_Qty']:
            df_activity[col] = df_activity[col].round(0).astype(int)
        df = pd.merge(df_list, df_activity[['PartNo', 'Qty_Sold_Calc', 'MAX_Qty', 'MIN_Qty', 'Re_Order_Point', 'Re_Order_Qty']], on='PartNo', how='left')
        df['Vendor'] = "Crader Dist. (STIHL)"

    elif vendor == "Vendor B":
        # Example: Different logic for Vendor B
        df = df_list.copy()
        df['Vendor'] = "Vendor B"
        # 📝 Add Vendor B specific calculations here...

    elif vendor == "Vendor C":
        # Example: Different logic for Vendor C
        df = df_list.copy()
        df['Vendor'] = "Vendor C"
        # 📝 Add Vendor C specific calculations here...

    # 🖼️ Output logic below remains the same (formatting, download)
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

