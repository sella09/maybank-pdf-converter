import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA Table Mode)")

def to_float_indo_bca(s):
    if not s: return None
    try: 
        clean = re.sub(r"[^\d.-]", "", s)
        return float(clean)
    except: return None

def parse_bca_tables(file_obj, password=None):
    transactions = []
    pdf_args = {"password": password} if password else {}
    
    with pdfplumber.open(file_obj, **pdf_args) as pdf:
        tahun = "2026"
        page1_text = pdf.pages[0].extract_text() or ""
        match_thn = re.search(r"PERIODE\s*:\s*\w+\s+(\d{4})", page1_text, re.IGNORECASE)
        if match_thn: tahun = match_thn.group(1)

        for page in pdf.pages:
            # Gunakan mode "text" karena PDF ini tidak punya garis tabel yang jelas
            tables = page.extract_tables({
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "text_tolerance": 3,  # Toleransi spasi lebih besar
                "intersection_tolerance": 15
            })
            
            if not tables: continue
            
            for table in tables:
                for row in table:
                    row = [(str(c).replace("\n", " ").strip() if c else "") for c in row]
                    
                    if len(row) < 4: continue
                    
                    # Cari sel yang berisi format tanggal DD/MM
                    date_idx = -1
                    for i, cell in enumerate(row):
                        if re.match(r"^\d{2}/\d{2}$", cell):
                            date_idx = i
                            break
                    
                    if date_idx == -1: continue
                    
                    tgl_bulan = row[date_idx]
                    keterangan = row[date_idx+1] if len(row) > date_idx+1 else ""
                    
                    # Gabungkan sisa teks di kanan tanggal
                    for i in range(date_idx+2, len(row)):
                        if row[i]:
                            keterangan += " " + row[i]
                    
                    # Ekstrak angka dari seluruh baris
                    numbers = re.findall(r"-?\d{1,3}(?:[,.]\d{3})*\.\d{2}", " ".join(row))
                    
                    debit, kredit, saldo = None, None, None
                    
                    if len(numbers) >= 2:
                        nominal = to_float_indo_bca(numbers[-2])
                        saldo = to_float_indo_bca(numbers[-1])
                        if "DB" in " ".join(row).upper(): kredit = nominal
                        else: debit = nominal
                    elif len(numbers) == 1:
                        nominal = to_float_indo_bca(numbers[0])
                        if "DB" in " ".join(row).upper(): kredit = nominal
                        else: debit = nominal
                    
                    # Bersihkan keterangan
                    for n in numbers: keterangan = keterangan.replace(n, "")
                    keterangan = re.sub(r"\bDB\b", "", keterangan, flags=re.IGNORECASE)
                    keterangan = " ".join(keterangan.split())

                    transactions.append({
                        "Tanggal": f"{tgl_bulan}/{tahun}",
                        "Kode": "-",
                        "Keterangan": keterangan,
                        "Debit": debit,
                        "Kredit": kredit,
                        "Saldo": saldo
                    })
    return transactions

def build_df(transactions, sumber_file):
    rows = []
    for txn in transactions:
        rows.append({
            "Sumber File": sumber_file,
            "Tanggal": txn["Tanggal"],
            "Kode": txn["Kode"],
            "Keterangan": txn["Keterangan"],
            "Debit": txn["Debit"],
            "Kredit": txn["Kredit"],
            "Saldo": txn["Saldo"]
        })
    return pd.DataFrame(rows)

# UI
col1, col2, col3 = st.columns(3)
with col1:
    bank_choice = st.selectbox("Pilih Bank", ["BCA", "Maybank"])
with col2:
    password = st.text_input("🔑 Password PDF (jika ada)", type="password")
with col3:
    uploaded_files = st.file_uploader("📄 Pilih file PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        try:
            st.info(f"Memproses **{uploaded_file.name}**...")
            txns = parse_bca_tables(uploaded_file, password=password)
            df = build_df(txns, uploaded_file.name)
            
            if not df.empty:
                all_dfs.append(df)
                st.success(f"✅ {uploaded_file.name}: {len(df)} transaksi")
            else:
                st.warning(f"⚠️ {uploaded_file.name}: tidak ada transaksi terdeteksi.")
                
        except Exception as e:
            st.error(f"❌ **{uploaded_file.name}**: {type(e).__name__}")
            with st.expander("Lihat detail error"):
                st.exception(e)

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        
        try:
            final_df['Tanggal_DT'] = pd.to_datetime(final_df['Tanggal'], format='%d/%m/%Y', errors='coerce')
            final_df = final_df.sort_values(by=['Tanggal_DT'])
            final_df = final_df.drop(columns=['Tanggal_DT'])
        except Exception as e:
            st.warning(f"Gagal mengurutkan tanggal: {e}")

        st.subheader("📊 Preview Data")
        st.dataframe(final_df, width='stretch')
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Mutasi')
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 Download sebagai Excel",
            data=processed_data,
            file_name="mutasi_gabungan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
