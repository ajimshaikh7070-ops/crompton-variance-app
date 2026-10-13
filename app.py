import streamlit as st
import pandas as pd
import sqlite3
import io
from datetime import date

st.set_page_config(page_title="Crompton ERP & Portal", layout="wide")

# Initialize SQLite Database
def init_db():
    conn = sqlite3.connect('crompton_erp.db', check_same_thread=False)
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

# Navigation Tabs / Modules
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📄 Tax Invoice", 
    "📂 Invoice History", 
    "📥 Monthly Inward", 
    "📤 Monthly Outward", 
    "📊 Monthly Plan & BOM", 
    "⚡ Variance Report"
])

# ==========================================
# TAB 1: TAX INVOICE GENERATOR
# ==========================================
with tab1:
    st.title("📄 SHRAVANI ELECTRICALS - Tax Invoice Generator")
    st.write("Neeche diye gaye details fill karein, database me save karein aur instant Excel download karein.")

    st.markdown("---")
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("📋 Invoice & Transport Details")
        inv_no = st.text_input("Invoice No.", value="1128", key="inv_no_field")
        inv_date = st.date_input("Invoice Date", value=date.today(), key="inv_date_field")
        vehicle_no = st.text_input("Vehicle No.", value="MH17AG6469", key="veh_field")
        transport_name = st.text_input("Transport Name", value="EXCELLENT LOGISTIC", key="trans_field")
        po_no = st.text_input("Purchase Order No.", value="5500130000", key="po_field")

    with c2:
        st.subheader("📦 Item Details")
        desc = st.text_input("Description", value="STATOR WINDING FOR NADK205_4_70E3", key="desc_field")
        mat_code = st.text_input("Material Code", value="NADK205/4-70E3", key="mat_field")
        qty = st.number_input("Quantity", min_value=1.0, value=3.0, step=1.0, key="qty_field")
        rate = st.number_input("Rate per unit (₹)", min_value=0.0, value=421.0, step=1.0, key="rate_field")

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
# TAB 2: INVOICE HISTORY
# ==========================================
with tab2:
    st.title("📂 Saved Invoices History")
    st.write("Database me save kiye gaye saare invoices ka record yahan dikhega.")

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

# ==========================================
# TAB 3: MONTHLY INWARD
# ==========================================
with tab3:
    st.title("📥 Monthly Inward Module")
    st.write("Apni monthly Inward Excel file yahan upload karke data preview karein.")
    
    inward_file_tab = st.file_uploader("Upload Monthly Inward Excel", type=["xlsx", "xls"], key="inward_uploader_tab")
    
    if inward_file_tab:
        df_inward_preview = pd.read_excel(inward_file_tab)
        st.success("✅ Inward file successfully loaded!")
        st.dataframe(df_inward_preview)

# ==========================================
# TAB 4: MONTHLY OUTWARD
# ==========================================
with tab4:
    st.title("📤 Monthly Outward Module")
    st.write("Apni monthly Outward Excel file yahan upload karke review karein.")
    
    outward_file_tab = st.file_uploader("Upload Monthly Outward Excel", type=["xlsx", "xls"], key="outward_uploader_tab")
    
    if outward_file_tab:
        df_outward_preview = pd.read_excel(outward_file_tab)
        st.success("✅ Outward file successfully loaded!")
        st.dataframe(df_outward_preview)

# ==========================================
# TAB 5: MONTHLY PLAN & BOM SETUP
# ==========================================
with tab5:
    st.title("📊 Monthly Plan & Master BOM Setup")
    st.write("Quarterly Master BOM aur Frame Ratings/Plans wali file yahan upload karein.")
    
    bom_file_tab = st.file_uploader("Upload Crompton Master BOM Excel", type=["xlsx", "xls"], key="bom_uploader_tab")
    
    if bom_file_tab:
        bom_xls_tab = pd.ExcelFile(bom_file_tab)
        selected_tab_sheet = st.selectbox("Select BOM Sheet", bom_xls_tab.sheet_names, key="bom_sheet_select_tab")
        skip_r = st.number_input("Header Row Index (Skip Top Rows)", min_value=0, max_value=20, value=0, step=1, key="skip_row_tab")
        
        df_bom_preview = pd.read_excel(bom_xls_tab, sheet_name=selected_tab_sheet, skiprows=skip_r)
        st.success(f"✅ Sheet '{selected_tab_sheet}' successfully loaded!")
        st.dataframe(df_bom_preview)

# ==========================================
# TAB 6: VARIANCE REPORT
# ==========================================
with tab6:
    st.title("⚡ Material Variance Report (Inward vs BOM)")
    st.write("Inward purchase rates ko Master BOM rates ke sath compare karke Loss/Gain check karein.")

    col_v1, col_v2 = st.columns(2)
    with col_v1:
        v_bom_file = st.file_uploader("Upload Master BOM for Variance", type=["xlsx", "xls"], key="v_bom")
    with col_v2:
        v_inward_file = st.file_uploader("Upload Monthly Inward for Variance", type=["xlsx", "xls"], key="v_inward")

    def find_col_safe(df, possible_names):
        for col in df.columns:
            clean_col = str(col).strip().lower()
            for name in possible_names:
                if name.lower() in clean_col:
                    return col
        return None

    if v_bom_file and v_inward_file:
        try:
            v_xls = pd.ExcelFile(v_bom_file)
            v_sheet = st.selectbox("Select BOM Sheet for Variance", v_xls.sheet_names, key="v_sheet_sel")
            v_skip = st.number_input("BOM Header Row Index", min_value=0, max_value=20, value=0, step=1, key="v_skip_input")
            
            df_v_bom = pd.read_excel(v_xls, sheet_name=v_sheet, skiprows=v_skip)
            df_v_inward = pd.read_excel(v_inward_file, sheet_name=0)

            b_mat = find_col_safe(df_v_bom, ['Material Code', 'Material', 'Item Code', 'Part No', 'Code'])
            b_rat = find_col_safe(df_v_bom, ['BOM Rate', 'Rate', 'Unit Rate', 'Price', 'Basic Rate', 'Amount'])

            i_mat = find_col_safe(df_v_inward, ['Item code', 'Item Code', 'Material Code', 'Material', 'Part No'])
            i_qty = find_col_safe(df_v_inward, ['Received qty', 'Inward Qty', 'Qty', 'Quantity', 'Received Qty'])
            i_rat = find_col_safe(df_v_inward, ['Purchase Rate', 'Actual Rate', 'Rate', 'Unit Rate'])
            i_amt = find_col_safe(df_v_inward, ['Total Invoice Amount', 'Total Amount', 'Amount'])

            if not b_mat or not b_rat:
                st.error("⚠️ BOM sheet me Material Code ya Rate column nahi mila. Header row index adjust karein.")
            elif not i_mat or not i_qty:
                st.error("⚠️ Inward sheet me Item code ya Qty column nahi mila.")
            else:
                df_vb_clean = df_v_bom.copy()
                df_vb_clean['Material Code'] = df_vb_clean[b_mat].astype(str).str.strip().str.upper()
                df_vb_clean['BOM Rate'] = pd.to_numeric(df_vb_clean[b_rat], errors='coerce').fillna(0)
                df_vb_clean = df_vb_clean.drop_duplicates(subset=['Material Code'], keep='first')

                df_vi_clean = df_v_inward.copy()
                df_vi_clean['Material Code'] = df_vi_clean[i_mat].astype(str).str.strip().str.upper()
                df_vi_clean['Inward Qty'] = pd.to_numeric(df_vi_clean[i_qty], errors='coerce').fillna(0)

                if i_rat:
                    df_vi_clean['Purchase Rate'] = pd.to_numeric(df_vi_clean[i_rat], errors='coerce').fillna(0)
                elif i_amt:
                    df_vi_clean['Total Amount'] = pd.to_numeric(df_vi_clean[i_amt], errors='coerce').fillna(0)
                    df_vi_clean['Purchase Rate'] = df_vi_clean.apply(
                        lambda row: (row['Total Amount'] / row['Inward Qty']) if row['Inward Qty'] > 0 and row['Total Amount'] > 0 else 0, axis=1
                    )
                else:
                    df_vi_clean['Purchase Rate'] = 0

                merged_v = pd.merge(df_vi_clean, df_vb_clean[['Material Code', 'BOM Rate']], on='Material Code', how='left')
                merged_v['BOM Rate'] = merged_v['BOM Rate'].fillna(0)
                merged_v['Price Difference (Per Unit)'] = merged_v['Purchase Rate'] - merged_v['BOM Rate']
                merged_v['Total Variance Amount'] = merged_v['Price Difference (Per Unit)'] * merged_v['Inward Qty']

                tot_inward = (merged_v['Inward Qty'] * merged_v['Purchase Rate']).sum()
                tot_bom = (merged_v['Inward Qty'] * merged_v['BOM Rate']).sum()
                net_var = merged_v['Total Variance Amount'].sum()

                st.markdown("---")
                v1, v2, v3 = st.columns(3)
                v1.metric("Total Inward Cost", f"₹ {tot_inward:,.2f}")
                v2.metric("Total BOM Allowed Cost", f"₹ {tot_bom:,.2f}")
                v3.metric("Net Variance (Loss/Gain)", f"₹ {abs(net_var):,.2f}", delta_color="inverse" if net_var > 0 else "normal")

                st.markdown("---")
                st.dataframe(merged_v)

                v_buffer = io.BytesIO()
                with pd.ExcelWriter(v_buffer, engine='openpyxl') as writer:
                    merged_v.to_excel(writer, index=False, sheet_name='Variance Report')

                st.download_button(
                    label="📥 Download Variance Report Excel",
                    data=v_buffer.getvalue(),
                    file_name="Crompton_Variance_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Variance Calculation Error: {e}")
