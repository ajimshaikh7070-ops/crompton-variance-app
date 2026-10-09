import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Crompton BOM Variance Portal", layout="wide")

st.title("⚡ Crompton Material & BOM Variance Portal")
st.write("Monthly Inward, Outward aur BOM rates compare karke instant Variance Report download karein.")

st.sidebar.header("1. Master BOM Setup (Quarterly)")
bom_file = st.sidebar.file_uploader("Upload Crompton Master BOM Excel", type=["xlsx", "xls"])

st.header("2. Monthly Processing")
col1, col2 = st.columns(2)

with col1:
    inward_file = st.file_uploader("Upload Monthly Inward Excel", type=["xlsx", "xls"])

with col2:
    outward_file = st.file_uploader("Upload Monthly Outward Excel (Optional)", type=["xlsx", "xls"])

def find_column(df, possible_names):
    for col in df.columns:
        clean_col = str(col).strip().lower()
        for name in possible_names:
            if name.lower() in clean_col:
                return col
    return None

def load_excel_smart(file_obj, sheet_name=0):
    # Reads Excel by auto-detecting the real header row if top rows contain titles/blank spaces
    xls = pd.ExcelFile(file_obj)
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = row.astype(str).str.lower().to_list()
        # Look for typical header indicators
        if any(keyword in ' '.join(row_str) for keyword in ['material', 'item code', 'part', 'code', 'rate', 'price', 'description']):
            header_idx = idx
            break
            
    df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_idx)
    return df

if bom_file and inward_file:
    try:
        df_bom = load_excel_smart(bom_file)
        df_inward = load_excel_smart(inward_file)

        # Smart column detection for BOM
        bom_mat_col = find_column(df_bom, ['Material Code', 'Material', 'Item Code', 'Item code', 'Part No', 'Material_Code', 'Code'])
        bom_rate_col = find_column(df_bom, ['BOM Rate', 'Rate', 'BOM_Rate', 'Unit Rate', 'Price', 'Basic Rate', 'Amount'])

        # Smart column detection for Inward Sheet
        inward_mat_col = find_column(df_inward, ['Item code', 'Item Code', 'Material Code', 'Material', 'Part No', 'Code'])
        inward_qty_col = find_column(df_inward, ['Received qty', 'Inward Qty', 'Qty', 'Quantity', 'Inward_Qty', 'Received Qty'])
        inward_rate_col = find_column(df_inward, ['Purchase Rate', 'Actual Rate', 'Rate', 'Unit Rate', 'Price'])
        inward_amount_col = find_column(df_inward, ['Total Invoice Amount', 'Total Amount', 'Amount', 'Value'])

        if not bom_mat_col or not bom_rate_col:
            st.error(f"⚠️ BOM Excel mein Material Code ya Rate column nahi mila. Detected headers: {list(df_bom.columns)}")
        elif not inward_mat_col or not inward_qty_col:
            st.error(f"⚠️ Inward Excel mein 'Item code' ya 'Received qty' column nahi mila. Detected headers: {list(df_inward.columns)}")
        else:
            # Clean BOM DataFrame
            df_bom_clean = df_bom.rename(columns={bom_mat_col: 'Material Code', bom_rate_col: 'BOM Rate'})
            df_bom_clean['Material Code'] = df_bom_clean['Material Code'].astype(str).str.strip()

            # Clean Inward DataFrame
            df_inward_clean = df_inward.copy()
            df_inward_clean['Material Code'] = df_inward_clean[inward_mat_col].astype(str).str.strip()
            df_inward_clean['Inward Qty'] = pd.to_numeric(df_inward_clean[inward_qty_col], errors='coerce').fillna(0)

            # Handle Purchase Rate / Unit Price
            if inward_rate_col:
                df_inward_clean['Purchase Rate'] = pd.to_numeric(df_inward_clean[inward_rate_col], errors='coerce').fillna(0)
            elif inward_amount_col:
                df_inward_clean['Total Amount'] = pd.to_numeric(df_inward_clean[inward_amount_col], errors='coerce').fillna(0)
                df_inward_clean['Purchase Rate'] = df_inward_clean.apply(
                    lambda row: (row['Total Amount'] / row['Inward Qty']) if row['Inward Qty'] > 0 and row['Total Amount'] > 0 else 0,
                    axis=1
                )
            else:
                df_inward_clean['Purchase Rate'] = 0

            # Merge Inward with BOM
            merged_df = pd.merge(df_inward_clean, df_bom_clean[['Material Code', 'BOM Rate']], on='Material Code', how='left')

            merged_df['BOM Rate'] = pd.to_numeric(merged_df['BOM Rate'], errors='coerce').fillna(0)
            merged_df['Price Difference (Per Unit)'] = merged_df['Purchase Rate'] - merged_df['BOM Rate']
            merged_df['Total Variance Amount'] = merged_df['Price Difference (Per Unit)'] * merged_df['Inward Qty']

            total_inward_cost = (merged_df['Inward Qty'] * merged_df['Purchase Rate']).sum()
            total_bom_cost = (merged_df['Inward Qty'] * merged_df['BOM Rate']).sum()
            net_variance = merged_df['Total Variance Amount'].sum()

            st.markdown("---")
            st.subheader("📊 Summary Overview")
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Inward Purchase Cost", f"₹ {total_inward_cost:,.2f}")
            m2.metric("Total BOM Allowed Cost", f"₹ {total_bom_cost:,.2f}")
            m3.metric("Net Variance Amount", f"₹ {net_variance:,.2f}")

            st.markdown("---")
            st.subheader("📋 Detailed Material Variance Table")
            st.dataframe(merged_df)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                merged_df.to_excel(writer, index=False, sheet_name='Variance Report')
                if outward_file:
                    df_outward = pd.read_excel(outward_file)
                    df_outward.to_excel(writer, index=False, sheet_name='Outward Summary')

            st.download_button(
                label="📥 Download Final Report Excel (for Crompton)",
                data=buffer.getvalue(),
                file_name="Crompton_BOM_Variance_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Processing Error: {e}")
