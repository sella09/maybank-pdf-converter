import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel Converter", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA & Maybank - Fixed)")

def parse_bca_perfect(file_obj, password=None):
    transactions = []
    pdf_args = {"password": password} if password else {}
    
    with pdfplumber.open(file_obj, **pdf_args) as pdf:
        full_text = ""
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    lines = full_text.split('\n')
    
    # Deteksi Tahun & Periode
    tahun = "2026"
    match_thn = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", full_text, re.IGNORECASE)
    if match_thn:
        tahun = match_thn.group(1)

    raw_txns = []
    current_tx = None
    
    # Pattern Tanggal BCA: DD/MM di awal baris
    date_pattern = re.compile(r"^(\d{2}/\d{2})\s+(.+)")

    for line in lines:
        line_str = line.strip()
        
        # Abaikan baris header / footer bawaan PDF BCA
        if not line_str or any(header_text in line_str for header_text in [
            "TANGGAL KETERANGAN", "Bersambung ke", "SALDO AWAL", "NO. REKENING",
            "REKENING GIRO", "HALAMAN", "PERIODE", "MATA UANG", "CATATAN:"
        ]):
            continue

        m_date = date_pattern.match(line_str)
        if m_date:
            if current_tx:
                raw_txns.append(current_tx)
            current_tx = {
                "tgl": m_date.group(1),
                "text_lines": [m_date.group(2)]
            }
        elif current_tx:
            # Menggabungkan baris lanjutan deskripsi transaksi
            current_tx["text_lines"].append(line_str)

    if current_tx:
        raw_txns.append(current_tx)

    # Parsing detail tiap blok transaksi
    for tx in raw_txns:
        combined = " ".join(tx["text_lines"])
        
        # Ekstrak semua format angka nominal (contoh: 1,000,000.00 atau 1000000.00)
        amounts = re.findall(r"([\d,]+\.\d{2})\s*(DB)?", combined, flags=re.IGNORECASE)
        
        if not amounts:
            continue

        # Pasangan angka terakhir biasanya adalah (Nominal Mutasi, Saldo)
        # Ambil nilai mutasi dan penanda DB (Debet)
        mutasi_str, is_db_flag = amounts[-1][0], amounts[-1][1].upper() == 'DB'
        
        nominal = float(mutasi_str.replace(',', ''))
        
        # Ambil Saldo jika tersedia
        saldo = None
        if len(amounts) >= 2:
            saldo = float(amounts[-1][0].replace(',', ''))
            mutasi_str, is_db_flag = amounts[-2][0], amounts[-2][1].upper() == 'DB'
            nominal = float(mutasi_str.replace(',', ''))

        # Bersihkan deskripsi dari angka nominal dan bendera DB
        desc = combined
        for amt_tuple in amounts:
            desc = desc.replace(amt_tuple[0], "")
        desc = re.sub(r"\bDB\b", "", desc, flags=re.IGNORECASE)
        desc = " ".join(desc.split()).strip()

        # LOGIKA PEMBALIKAN (SWAP):
        # Asli DB (Debet/Pengeluaran BCA)  -> Pindah ke KREDIT (atau kolom Debet Swap)
        # Asli CR (Kredit/Pemasukan BCA) -> Pindah ke DEBET (atau kolom Kredit Swap)
        if is_db_flag:
            debit = 0.0
            kredit = nominal
        else:
            debit = nominal
            kredit = 0.0

        transactions.append({
            "Tanggal": f"{tx['tgl']}/{tahun}",
            "Kode": "-",
            "Keterangan": desc,
            "Debit": debit,
            "Kredit": kredit,
            "Saldo": saldo,
            "Tipe Mutasi Asli": "DB" if is_db_flag else "CR"
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
            "Saldo": txn["Saldo"],
            "Tipe Asli": txn.get("Tipe Mutasi Asli", "-")
        })
    return pd.DataFrame(rows)

# Interface Streamlit
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
            
            if bank_choice == "BCA":
                txns = parse_bca_perfect(uploaded_file, password=password)
            else:
                # Tambahkan fungsi parser Maybank jika diperlukan
                txns = parse_bca_perfect(uploaded_file, password=password)
                
            df = build_df(txns, uploaded_file.name)
            
            if not df.empty:
                all_dfs.append(df)
                st.success(f"✅ {uploaded_file.name}: {len(df)} transaksi berhasil diekstrak.")
            else:
                st.warning(f"⚠️ {uploaded_file.name}: Tidak ada transaksi terdeteksi.")
                
        except Exception as e:
            st.error(f"❌ Error saat memproses **{uploaded_file.name}**: {type(e).__name__}")
            with st.expander("Lihat detail error"):
                st.exception(e)

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        
        try:
            final_df['Tanggal_DT'] = pd.to_datetime(final_df['Tanggal'], format='%d/%m/%Y', errors='coerce')
            final_df = final_df.sort_values(by=['Tanggal_DT']).drop(columns=['Tanggal_DT'])
        except Exception as e:
            st.warning(f"Gagal mengurutkan tanggal: {e}")

        st.subheader("📊 Preview Data Hasil Swap")
        st.dataframe(final_df, use_container_width=True)
        
        # Ekspor Ke Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Mutasi_Swapped')
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 Download Excel Hasil Swap",
            data=processed_data,
            file_name="mutasi_bca_swapped.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
