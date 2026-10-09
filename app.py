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

def process_file_with_skip(file_obj, header_row=0, sheet_name=0):
    xls = pd.ExcelFile(file_obj)
    sheet = sheet_name if sheet_name in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet, skiprows=header_row)
    return df, xls.sheet_names, sheet

if bom_file and inward_file:
    try:
        # Side controls for header rows if needed
        st.sidebar.markdown("---")
        st.sidebar.subheader("⚙️ Sheet & Header Settings")
        
        bom_xls = pd.ExcelFile(bom_file)
        selected_bom_sheet = st.sidebar.selectbox("Select BOM Sheet", bom_xls.sheet_names)
        bom_header_row = st.sidebar.number_input("BOM Header Row Index (Skip Top Rows)", min_value=0, max_value=20, value=2, step=1)
        
        df_bom, _, _ = process_file_with_skip(bom_file, header_row=bom_header_row, sheet_name=selected_bom_sheet)
        df_inward, _, _ = process_file_with_skip(inward_file, header_row=0)

        # Smart column detection
        bom_mat_col = find_column(df_bom, ['Material Code', 'Material', 'Item Code', 'Part No', 'Code', 'Item'])
        bom_rate_col = find_column(df_bom, ['BOM Rate', 'Rate', 'Unit Rate', 'Price', 'Basic Rate', 'Amount'])

        inward_mat_col = find_column(df_inward, ['Item code', 'Item Code', 'Material Code', 'Material', 'Part No', 'Code'])
        inward_qty_col = find_column(df_inward, ['Received qty', 'Inward Qty', 'Qty', 'Quantity', 'Inward_Qty', 'Received Qty'])
        inward_rate_col = find_column(df_inward, ['Purchase Rate', 'Actual Rate', 'Rate', 'Unit Rate', 'Price'])
        inward_amount_col = find_column(df_inward, ['Total Invoice Amount', 'Total Amount', 'Amount', 'Value'])

        if not bom_mat_col or not bom_rate_col:
            st.error(f"⚠️ BOM Excel ({selected_bom_sheet}) mein Material/Rate column nahi mila.")
            st.info("👉 Sidebar me *BOM Header Row Index* ko change karke (1, 2, ya 3) check karein.")
            st.write("BOM Sheet Preview:")
            st.dataframe(df_bom.head(5))
        elif not inward_mat_col or not inward_qty_col:
            st.error(f"⚠️ Inward Excel mein 'Item code' ya 'Received qty' column nahi mila.")
            st.write("Inward Sheet Preview:")
            st.dataframe(df_inward.head(5))
        else:
            # Clean BOM DataFrame
            df_bom_clean = df_bom.rename(columns={bom_mat_col: 'Material Code', bom_rate_col: 'BOM Rate'})
            df_bom_clean['Material Code'] = df_bom_clean['Material Code'].astype(str).str.strip()

            # Clean Inward DataFrame
            df_inward_clean = df_inward.copy()
            df_inward_clean['Material Code'] = df_inward_clean[inward_mat_col].astype(str).str.strip()
            df_inward_clean['Inward Qty'] = pd.to_numeric(df_inward_clean[inward_qty_col], errors='coerce').fillna(0)

            # Purchase Rate handling
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

            # Merge
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
