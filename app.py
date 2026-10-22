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
            po_date TEXT,
            description TEXT,
            material_code TEXT,
            sac_code TEXT,
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
# TAB 1: TAX INVOICE GENERATOR & PRINT
# ==========================================
with tab1:
    st.title("📄 SHRAVANI ELECTRICALS - Tax Invoice Generator")
    st.write("Details fill karein, live preview dekhein, database me save karein ya direct PDF Print karein.")

    st.markdown("---")
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("📋 Invoice & Transport Details")
        inv_no = st.text_input("Invoice No.", value="1128", key="inv_no_field")
        inv_date = st.date_input("Invoice Date", value=date.today(), key="inv_date_field")
        vehicle_no = st.text_input("Vehicle No.", value="MH17AG6469", key="veh_field")
        transport_name = st.text_input("Transport Name", value="EXCELLENT LOGISTIC", key="trans_field")
        po_no = st.text_input("Purchase Order No.", value="5500130000", key="po_field")
        po_date = st.date_input("Purchase Order Date", value=date.today(), key="po_date_field")

    with c2:
        st.subheader("📦 Item Details")
        sno = st.text_input("S.No.", value="200", key="sno_field")
        desc = st.text_input("Description", value="STATOR WINDING FOR NADK205_4_70E3", key="desc_field")
        mat_code = st.text_input("Material Code", value="NADK205/4-70E3", key="mat_field")
        sac_code = st.text_input("SAC Code", value="9988", key="sac_field")
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

    b_col1, b_col2, b_col3 = st.columns(3)

    with b_col1:
        if st.button("💾 Save to Database"):
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO invoices (invoice_no, invoice_date, vehicle_no, transport_name, po_no, po_date, description, material_code, sac_code, qty, rate, basic_amount, cgst, sgst, grand_total)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (inv_no, str(inv_date), vehicle_no, transport_name, po_no, str(po_date), desc, mat_code, sac_code, qty, rate, basic_amount, cgst_amount, sgst_amount, grand_total))
            conn.commit()
            st.success(f"✅ Invoice #{inv_no} saved!")

    with b_col2:
        single_inv_df = pd.DataFrame([{
            "Invoice No": inv_no, "Invoice Date": str(inv_date), "Vehicle No": vehicle_no, "Transport": transport_name,
            "PO No": po_no, "PO Date": str(po_date), "Description": desc, "Material Code": mat_code, "SAC": sac_code,
            "Quantity": qty, "Rate": rate, "Basic Amount": basic_amount, "CGST (9%)": cgst_amount, "SGST (9%)": sgst_amount, "Grand Total": grand_total
        }])
        single_buffer = io.BytesIO()
        with pd.ExcelWriter(single_buffer, engine='openpyxl') as writer:
            single_inv_df.to_excel(writer, index=False, sheet_name='Tax Invoice')

        st.download_button(
            label="📥 Download Excel",
            data=single_buffer.getvalue(),
            file_name=f"Tax_Invoice_{inv_no}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with b_col3:
        st.markdown("""
            <style>
            .print-btn {
                background-color: #ff4b4b;
                color: white;
                padding: 9px 20px;
                text-align: center;
                text-decoration: none;
                display: inline-block;
                font-size: 14px;
                font-weight: bold;
                border-radius: 4px;
                cursor: pointer;
                border: none;
                width: 100%;
            }
            .print-btn:hover { background-color: #ff2121; }
            </style>
            <button onclick="window.print();" class="print-btn">🖨️ Print / Save as PDF</button>
        """, unsafe_allow_html=True)

    # Professional Tax Invoice HTML Preview Layout
    st.markdown("---")
    st.subheader("👁️ Formal Tax Invoice Layout Preview")

    invoice_html = f"""
    <div style="border: 2px solid #000; padding: 15px; font-family: Arial, sans-serif; background-color: #fff; color: #000; max-width: 800px; margin: auto;">
        <h2 style="text-align: center; margin: 0; font-size: 20px;">Tax Invoice</h2>
        <div style="text-align: center; font-size: 14px; font-weight: bold; margin-top: 5px;">Supplier : SHRAVANI ELECTRICALS</div>
        <div style="text-align: center; font-size: 12px; margin-top: 2px;">
            Vender code : 7900393 &nbsp;|&nbsp; Address: W-2 MIDC, AHMEDNAGAR<br>
            State : Maharashtra &nbsp;|&nbsp; State code : 27 &nbsp;|&nbsp; Country: India &nbsp;|&nbsp; Pin code: 414111<br>
            GSTIN : 27AEJFS5536G1ZQ &nbsp;|&nbsp; PAN : AEJFS5536G<br>
            Contact details : jaysing musale mo 9527463787 &nbsp;|&nbsp; Email: shravani.electrical7@gmail.com
        </div>
        
        <table style="width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px;" border="1">
            <tr>
                <td style="width: 50%; vertical-align: top; padding: 6px;">
                    <strong>Bill TO : CG POWER AND INDUSTRIAL</strong><br>
                    ADdress: A-6/2 MIDC, AHMEDNAGAR<br>
                    State : Maharashtra &nbsp;|&nbsp; State code : 27<br>
                    Place of supply : maharashtra &nbsp;|&nbsp; Country: India &nbsp;|&nbsp; Pin code: 414111<br>
                    GSTIN : 27AAACC3840K1ZP &nbsp;|&nbsp; PAN : AAACC3840K
                </td>
                <td style="width: 50%; vertical-align: top; padding: 6px;">
                    <strong>TRANSPORT: {transport_name}</strong><br>
                    <strong>Invoice No.:</strong> {inv_no} &nbsp;&nbsp;|&nbsp;&nbsp; <strong>VHEHICAL NO:</strong> {vehicle_no}<br>
                    <strong>Invoice Date:</strong> {inv_date}<br>
                    <strong>Purchase Order No.:</strong> {po_no}<br>
                    <strong>Purchase Order Date:</strong> {po_date}
                </td>
            </tr>
        </table>

        <table style="width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px;" border="1">
            <tr style="background-color: #f2f2f2; text-align: center;">
                <th style="padding: 5px;">S.No.</th>
                <th style="padding: 5px;">Description</th>
                <th style="padding: 5px;">MATERIAL CODE</th>
                <th style="padding: 5px;">SAC</th>
                <th style="padding: 5px;">Quantity</th>
                <th style="padding: 5px;">per</th>
                <th style="padding: 5px;">RATE</th>
                <th style="padding: 5px;">Amount</th>
            </tr>
            <tr>
                <td style="padding: 6px; text-align: center;">{sno}</td>
                <td style="padding: 6px;">{desc}</td>
                <td style="padding: 6px; text-align: center;">{mat_code}</td>
                <td style="padding: 6px; text-align: center;">{sac_code}</td>
                <td style="padding: 6px; text-align: center;">{qty}</td>
                <td style="padding: 6px; text-align: center;">Nos</td>
                <td style="padding: 6px; text-align: right;">{rate:,.2f}</td>
                <td style="padding: 6px; text-align: right;">{basic_amount:,.2f}</td>
            </tr>
            <tr>
                <td colspan="6" style="text-align: right; padding: 5px;"><strong>TOTAL</strong></td>
                <td style="text-align: center; padding: 5px;">{qty}</td>
                <td style="text-align: right; padding: 5px;"><strong>{basic_amount:,.2f}</strong></td>
            </tr>
            <tr>
                <td colspan="6" style="text-align: right; padding: 4px;">Basic Amount</td>
                <td colspan="2" style="text-align: right; padding: 4px;">{basic_amount:,.2f}</td>
            </tr>
            <tr>
                <td colspan="6" style="text-align: right; padding: 4px;">CGST @ 9%</td>
                <td colspan="2" style="text-align: right; padding: 4px;">{cgst_amount:,.2f}</td>
            </tr>
            <tr>
                <td colspan="6" style="text-align: right; padding: 4px;">SGST @ 9%</td>
                <td colspan="2" style="text-align: right; padding: 4px;">{sgst_amount:,.2f}</td>
            </tr>
            <tr style="background-color: #f9f9f9;">
                <td colspan="6" style="text-align: right; padding: 6px;"><strong>Grand Total</strong></td>
                <td colspan="2" style="text-align: right; padding: 6px;"><strong>₹ {grand_total:,.2f}</strong></td>
            </tr>
        </table>

        <table style="width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 11px;" border="1">
            <tr>
                <td style="width: 33%; padding: 6px; vertical-align: top;"><strong>SECURITY ENTRY</strong><br><br><br>Recieved by &Date</td>
                <td style="width: 33%; padding: 6px; vertical-align: top;"><strong>CCI NO</strong><br><br><br>D&R Checked by &date</td>
                <td style="width: 34%; padding: 6px; vertical-align: top;"><strong>FOR SHRAVANI ELECTRICALS</strong><br><br><br><div style="text-align: center;">Authorized signatory</div>Inspection BY& DATE &nbsp;&nbsp;|&nbsp;&nbsp; Stores In Charge</td>
            </tr>
        </table>
    </div>
    """
    st.components.v1.html(invoice_html, height=580, scrolling=True)

# ==========================================
# TAB 2: INVOICE HISTORY
# ==========================================
with tab2:
    st.title("📂 Saved Invoices History")
    df_history = pd.read_sql("SELECT * FROM invoices ORDER BY id DESC", conn)
    if not df_history.empty:
        st.dataframe(df_history)
        hist_buffer = io.BytesIO()
        with pd.ExcelWriter(hist_buffer, engine='openpyxl') as writer:
            df_history.to_excel(writer, index=False, sheet_name='History')
        st.download_button("📥 Download All History Excel", data=hist_buffer.getvalue(), file_name="Invoices_History.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("No saved invoices found.")

# ==========================================
# TAB 3: MONTHLY INWARD
# ==========================================
with tab3:
    st.title("📥 Monthly Inward Module")
    inward_file_tab = st.file_uploader("Upload Inward Excel", type=["xlsx", "xls"], key="in_up")
    if inward_file_tab:
        st.dataframe(pd.read_excel(inward_file_tab))

# ==========================================
# TAB 4: MONTHLY OUTWARD
# ==========================================
with tab4:
    st.title("📤 Monthly Outward Module")
    outward_file_tab = st.file_uploader("Upload Outward Excel", type=["xlsx", "xls"], key="out_up")
    if outward_file_tab:
        st.dataframe(pd.read_excel(outward_file_tab))

# ==========================================
# TAB 5: MONTHLY PLAN & BOM
# ==========================================
with tab5:
    st.title("📊 Monthly Plan & Master BOM Setup")
    bom_file_tab = st.file_uploader("Upload Master BOM Excel", type=["xlsx", "xls"], key="bom_up")
    if bom_file_tab:
        xls_b = pd.ExcelFile(bom_file_tab)
        sh = st.selectbox("Select Sheet", xls_b.sheet_names)
        st.dataframe(pd.read_excel(xls_b, sheet_name=sh))

# ==========================================
# TAB 6: VARIANCE REPORT
# ==========================================
with tab6:
    st.title("⚡ Material Variance Report (Inward vs BOM)")
    st.write("Inward purchase rates ko Master BOM rates ke sath compare karke Loss/Gain check karein.")

    col_v1, col_v2 = st.columns(2)
    with col_v1:
        v_bom = st.file_uploader("Upload BOM for Variance", type=["xlsx", "xls"], key="vb")
    with col_v2:
        v_in = st.file_uploader("Upload Inward for Variance", type=["xlsx", "xls"], key="vi")

    def find_col_safe(df, possible_names):
        for col in df.columns:
            clean_col = str(col).strip().lower()
            for name in possible_names:
                if name.lower() in clean_col:
                    return col
        return None

    if v_bom and v_in:
        try:
            v_xls = pd.ExcelFile(v_bom)
            v_sheet_names = v_xls.sheet_names
            v_sheet = st.selectbox("Select BOM Sheet for Variance", v_sheet_names, index=1 if len(v_sheet_names)>1 else 0, key="v_sheet_sel")
            v_skip = st.number_input("BOM Header Row Index (Skip Top Rows)", min_value=0, max_value=20, value=2, step=1, key="v_skip_input")
            
            df_v_bom = pd.read_excel(v_xls, sheet_name=v_sheet, skiprows=v_skip)
            df_v_inward = pd.read_excel(v_in, sheet_name=0)

            b_mat = find_col_safe(df_v_bom, ['Material Code', 'Material', 'Item Code', 'Part No', 'Code', 'Item'])
            b_rat = find_col_safe(df_v_bom, ['BOM Rate', 'Rate', 'Unit Rate', 'Price', 'Basic Rate', 'Amount', 'PO Rate'])

            i_mat = find_col_safe(df_v_inward, ['Item code', 'Item Code', 'Material Code', 'Material', 'Part No', 'Code'])
            i_qty = find_col_safe(df_v_inward, ['Received qty', 'Inward Qty', 'Qty', 'Quantity', 'Received Qty'])
            i_rat = find_col_safe(df_v_inward, ['Purchase Rate', 'Actual Rate', 'Rate', 'Unit Rate'])
            i_amt = find_col_safe(df_v_inward, ['Total Invoice Amount', 'Total Amount', 'Amount', 'Value'])

            if not b_mat or not b_rat:
                st.error(f"⚠️ BOM sheet ({v_sheet}) me Material Code ya Rate column nahi mila. Sidebar me 'BOM Header Row Index' (0, 1, 2, 3) change karke check karein.")
                st.write("BOM Preview:", df_v_bom.head(5))
            elif not i_mat or not i_qty:
                st.error("⚠️ Inward sheet me Item code ya Received qty column nahi mila.")
                st.write("Inward Preview:", df_v_inward.head(5))
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
                
                v_label = "Loss (Higher Purchase Rate)" if net_var > 0 else "Gain / Savings"
                v3.metric(f"Net Variance ({v_label})", f"₹ {abs(net_var):,.2f}")

                st.markdown("---")
                st.subheader("📋 Detailed Variance Table")
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
