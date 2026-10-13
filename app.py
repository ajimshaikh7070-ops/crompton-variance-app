import streamlit as st
import pandas as pd
import sqlite3
import io
from datetime import date

st.set_page_config(page_title="Crompton Portal & Invoice Generator", layout="wide")

# Initialize SQLite Database
def init_db():
    conn = sqlite3.connect('invoices.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_no TEXT,
            invoice_date TEXT,
            vehicle_no TEXT,
            transport_name TEXT,
            po_no TEXT,
            description TEXT,
            material_code TEXT,
            qty REAL,
            rate REAL,
            basic_amount REAL,
            cgst REAL,
            sgst REAL,
            grand_total REAL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    return conn

conn = init_db()

# Navigation Tabs
tab1, tab2, tab3 = st.tabs(["⚡ BOM Variance Portal", "📄 Tax Invoice Generator", "📂 Saved Invoices History"])

# ==========================================
# TAB 1: BOM VARIANCE PORTAL
# ==========================================
with tab1:
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

    if bom_file and inward_file:
        try:
            st.sidebar.markdown("---")
            st.sidebar.subheader("⚙️ BOM Sheet Settings")
            
            bom_xls = pd.ExcelFile(bom_file)
            default_sheet_index = 1 if len(bom_xls.sheet_names) > 1 else 0
            selected_bom_sheet = st.sidebar.selectbox("Select BOM Sheet", bom_xls.sheet_names, index=default_sheet_index)
            bom_header_row = st.sidebar.number_input("BOM Header Row Index (Skip Top Rows)", min_value=0, max_value=20, value=0, step=1)
            
            df_bom = pd.read_excel(bom_xls, sheet_name=selected_bom_sheet, skiprows=bom_header_row)
            df_inward = pd.read_excel(inward_file, sheet_name=0)

            bom_mat_col = find_column(df_bom, ['Material Code', 'Material', 'Item Code', 'Item code', 'Part No', 'Code', 'Item', 'Frame Ratings'])
            bom_rate_col = find_column(df_bom, ['BOM Rate', 'Rate', 'Unit Rate', 'Price', 'Basic Rate', 'Amount', 'Value', 'PO Rate', 'BOM'])

            inward_mat_col = find_column(df_inward, ['Item code', 'Item Code', 'Material Code', 'Material', 'Part No', 'Code'])
            inward_qty_col = find_column(df_inward, ['Received qty', 'Inward Qty', 'Qty', 'Quantity', 'Inward_Qty', 'Received Qty'])
            inward_rate_col = find_column(df_inward, ['Purchase Rate', 'Actual Rate', 'Rate', 'Unit Rate', 'Price'])
            inward_amount_col = find_column(df_inward, ['Total Invoice Amount', 'Total Amount', 'Amount', 'Value'])

            if not bom_mat_col or not bom_rate_col:
                st.error(f"⚠️ Selected BOM Sheet ({selected_bom_sheet}) me Material Code ya Rate column nahi mila.")
            elif not inward_mat_col or not inward_qty_col:
                st.error(f"⚠️ Inward Excel me 'Item code' ya 'Received qty' column nahi mila.")
            else:
                df_bom_clean = df_bom.copy()
                df_bom_clean['Material Code'] = df_bom_clean[bom_mat_col].astype(str).str.strip().str.upper()
                df_bom_clean['BOM Rate'] = pd.to_numeric(df_bom_clean[bom_rate_col], errors='coerce').fillna(0)
                df_bom_clean = df_bom_clean.drop_duplicates(subset=['Material Code'], keep='first')

                df_inward_clean = df_inward.copy()
                df_inward_clean['Material Code'] = df_inward_clean[inward_mat_col].astype(str).str.strip().str.upper()
                df_inward_clean['Inward Qty'] = pd.to_numeric(df_inward_clean[inward_qty_col], errors='coerce').fillna(0)

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

                merged_df = pd.merge(df_inward_clean, df_bom_clean[['Material Code', 'BOM Rate']], on='Material Code', how='left')

                merged_df['BOM Rate'] = merged_df['BOM Rate'].fillna(0)
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
                
                variance_label = "Loss (Higher Purchase Rate)" if net_variance > 0 else "Gain / Savings"
                m3.metric(f"Net Variance Amount ({variance_label})", f"₹ {abs(net_variance):,.2f}")

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
                    label="📥 Download Final Variance Report Excel",
                    data=buffer.getvalue(),
                    file_name="Crompton_BOM_Variance_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Processing Error: {e}")

# ==========================================
# TAB 2: TAX INVOICE GENERATOR
# ==========================================
with tab2:
    st.title("📄 SHRAVANI ELECTRICALS - Tax Invoice Generator")
    st.write("Neeche diye gaye details fill karein, database me save karein aur instant download karein.")

    st.markdown("---")
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("📋 Invoice & Transport Details")
        inv_no = st.text_input("Invoice No.", value="1128")
        inv_date = st.date_input("Invoice Date", value=date.today())
        vehicle_no = st.text_input("Vehicle No.", value="MH17AG6469")
        transport_name = st.text_input("Transport Name", value="EXCELLENT LOGISTIC")
        po_no = st.text_input("Purchase Order No.", value="5500130000")

    with c2:
        st.subheader("📦 Item Details")
        desc = st.text_input("Description", value="STATOR WINDING FOR NADK205_4_70E3")
        mat_code = st.text_input("Material Code", value="NADK205/4-70E3")
        qty = st.number_input("Quantity", min_value=1.0, value=3.0, step=1.0)
        rate = st.number_input("Rate per unit (₹)", min_value=0.0, value=421.0, step=1.0)

    # Calculations
    basic_amount = qty * rate
    cgst_amount = round(basic_amount * 0.09, 2)
    sgst_amount = round(basic_amount * 0.09, 2)
    grand_total = round(basic_amount + cgst_amount + sgst_amount, 2)

    st.markdown("---")
    st.subheader("💰 Calculated Summary")
    
    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.metric("Basic Amount", f"₹ {basic_amount:,.2f}")
    sc2.metric("CGST (9%)", f"₹ {cgst_amount:,.2f}")
    sc3.metric("SGST (9%)", f"₹ {sgst_amount:,.2f}")
    sc4.metric("Grand Total", f"₹ {grand_total:,.2f}")

    b_col1, b_col2 = st.columns(2)

    with b_col1:
        if st.button("💾 Save Invoice to Database"):
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO invoices (invoice_no, invoice_date, vehicle_no, transport_name, po_no, description, material_code, qty, rate, basic_amount, cgst, sgst, grand_total)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (inv_no, str(inv_date), vehicle_no, transport_name, po_no, desc, mat_code, qty, rate, basic_amount, cgst_amount, sgst_amount, grand_total))
            conn.commit()
            st.success(f"✅ Invoice #{inv_no} successfully database me save ho gaya!")

    with b_col2:
        # Prepare single invoice excel download
        single_inv_df = pd.DataFrame([{
            "Invoice No": inv_no,
            "Invoice Date": str(inv_date),
            "Vehicle No": vehicle_no,
            "Transport": transport_name,
            "PO No": po_no,
            "Description": desc,
            "Material Code": mat_code,
            "Quantity": qty,
            "Rate": rate,
            "Basic Amount": basic_amount,
            "CGST (9%)": cgst_amount,
            "SGST (9%)": sgst_amount,
            "Grand Total": grand_total
        }])
        
        single_buffer = io.BytesIO()
        with pd.ExcelWriter(single_buffer, engine='openpyxl') as writer:
            single_inv_df.to_excel(writer, index=False, sheet_name='Tax Invoice')

        st.download_button(
            label="📥 Download This Invoice as Excel",
            data=single_buffer.getvalue(),
            file_name=f"Tax_Invoice_{inv_no}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Live Preview
    st.markdown("---")
    st.subheader("👁️ Live Invoice Format Preview")

    invoice_html = f"""
    <div style="border: 2px solid #000; padding: 20px; font-family: Arial, sans-serif; background-color: #fff; color: #000;">
        <h2 style="text-align: center; margin-bottom: 5px;">Tax Invoice</h2>
        <h3 style="text-align: center; margin-top: 0;">Supplier : SHRAVANI ELECTRICALS</h3>
        <p style="text-align: center; font-size: 13px; margin: 2px;">Vender code : 7900393 | Address: W-2 MIDC, AHMEDNAGAR</p>
        <p style="text-align: center; font-size: 13px; margin: 2px;">State: Maharashtra (Code: 27) | GSTIN: 27AEJFS5536G1ZQ</p>
        <hr style="border-top: 1px solid #000;">
        <table style="width: 100%; font-size: 13px;">
            <tr>
                <td><strong>Bill TO : CG POWER AND INDUSTRIAL</strong><br>A-6/2 MIDC, AHMEDNAGAR</td>
                <td><strong>TRANSPORT: {transport_name}</strong><br><strong>Vehicle:</strong> {vehicle_no} | <strong>Inv No:</strong> {inv_no}</td>
            </tr>
        </table>
        <table style="width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px;" border="1">
            <tr style="background-color: #f2f2f2;">
                <th>Description</th><th>Material Code</th><th>Qty</th><th>Rate</th><th>Amount</th>
            </tr>
            <tr>
                <td style="padding: 6px;">{desc}</td>
                <td style="padding: 6px; text-align: center;">{mat_code}</td>
                <td style="padding: 6px; text-align: center;">{qty}</td>
                <td style="padding: 6px; text-align: right;">₹ {rate:,.2f}</td>
                <td style="padding: 6px; text-align: right;">₹ {basic_amount:,.2f}</td>
            </tr>
            <tr><td colspan="4" style="text-align: right;"><strong>Grand Total:</strong></td><td style="text-align: right;"><strong>₹ {grand_total:,.2f}</strong></td></tr>
        </table>
    </div>
    """
    st.components.v1.html(invoice_html, height=400, scrolling=True)

# ==========================================
# TAB 3: SAVED INVOICES HISTORY
# ==========================================
with tab3:
    st.title("📂 Saved Invoices History (SQLite Database)")
    st.write("Yahan aap apne saare saved invoices ka record dekh sakte hain.")

    df_history = pd.read_sql("SELECT * FROM invoices ORDER BY id DESC", conn)
    
    if not df_history.empty:
        st.dataframe(df_history)
        
        hist_buffer = io.BytesIO()
        with pd.ExcelWriter(hist_buffer, engine='openpyxl') as writer:
            df_history.to_excel(writer, index=False, sheet_name='Invoice_History')
            
        st.download_button(
            label="📥 Download All History as Excel",
            data=hist_buffer.getvalue(),
            file_name="All_Saved_Invoices_History.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Abhi tak koi invoice database me save nahi kiya gaya hai.")
