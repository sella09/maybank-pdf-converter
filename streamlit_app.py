import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import dataframe_to_rows

st.set_page_config(page_title="Maybank RK PDF to Excel Converter", page_icon="📊", layout="wide")

st.title("📊 Converter Rekening Koran Maybank PDF ke Excel")
st.write("Aplikasi ini membantu mengekstrak data Rekening Koran Maybank (termasuk PDF ber-password) dan mengonversinya ke format Excel (.xlsx).")

# Widget input file dan password
uploaded_files = st.file_uploader("Unggah File PDF Rekening Koran Maybank", type=["pdf"], accept_multiple_files=True)
pdf_password = st.text_input("Masukkan Password PDF (Kosongkan jika tidak ada)", type="password")

def clean_number(val_str):
    if not val_str:
        return 0.0
    val = val_str.replace(".", "").replace(",", ".").strip()
    try:
        return float(val)
    except ValueError:
        return 0.0

def process_pdf(pdf_file, password):
    all_rows = []
    try:
        # Membuka PDF dengan password jika ada
        with pdfplumber.open(pdf_file, password=password if password else None) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        row_cleaned = [str(c).replace('\n', ' ').strip() if c is not None else '' for c in row]
                        row_str = " ".join(row_cleaned).upper()
                        
                        # Lewati baris header tabel
                        if "TRANSAKSI REKENING GIRO" in row_str or ("TGL" in row_str and "KETERANGAN" in row_str):
                            continue
                        
                        if len(row_cleaned) >= 6:
                            tgl, kode, ket, debit, kredit, saldo = row_cleaned[:6]
                            
                            # Deteksi format tanggal DD/MM/YYYY
                            if re.match(r'^\d{2}/\d{2}/\d{4}$', tgl):
                                float_debit = clean_number(debit)
                                float_kredit = clean_number(kredit)
                                float_saldo = clean_number(saldo)
                                
                                all_rows.append({
                                    "TGL": tgl,
                                    "KODE": kode,
                                    "KETERANGAN": ket,
                                    "DEBIT": float_debit if float_debit != 0 else None,
                                    "KREDIT": float_kredit if float_kredit != 0 else None,
                                    "SALDO": float_saldo
                                })
                            elif all_rows and ket:
                                # Baris lanjutan keterangan transaksi
                                all_rows[-1]["KETERANGAN"] += " " + ket
    except Exception as e:
        st.error(f"Gagal memproses file {pdf_file.name}: {str(e)}")
        return None
        
    return pd.DataFrame(all_rows)

def create_excel(df_dict):
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    data_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    for sheet_name, df in df_dict.items():
        ws = wb.create_sheet(title=sheet_name[:31])
        headers = list(df.columns)
        ws.append(headers)
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        for row in dataframe_to_rows(df, index=False, header=False):
            ws.append(row)
            
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.font = data_font
                cell.border = thin_border
                col_name = headers[cell.column - 1]
                if col_name in ["DEBIT", "KREDIT", "SALDO"]:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_name in ["TGL", "KODE"]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    wb.save(output)
    return output.getvalue()

if uploaded_files:
    if st.button("🚀 Konversi Ke Excel", type="primary"):
        processed_data = {}
        for pdf_file in uploaded_files:
            df = process_pdf(pdf_file, pdf_password)
            if df is not None and not df.empty:
                sheet_title = pdf_file.name.replace(".pdf", "")
                processed_data[sheet_title] = df
        
        if processed_data:
            st.success("✅ Konversi Berhasil!")
            excel_bytes = create_excel(processed_data)
            
            st.download_button(
                label="📥 Download File Excel",
                data=excel_bytes,
                file_name="Rekening_Koran_Maybank.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            for name, df in processed_data.items():
                st.subheader(f"Preview Data: {name}")
                st.dataframe(df.style.format({"DEBIT": "{:,.2f}", "KREDIT": "{:,.2f}", "SALDO": "{:,.2f}"}))
