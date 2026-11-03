import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA Fixed & Perfected)")

def parse_bca_perfect(file_obj, password=None):
    transactions = []
    pdf_args = {"password": password} if password else {}
    
    with pdfplumber.open(file_obj, **pdf_args) as pdf:
        full_pages_text = []
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_pages_text.append(t)

    full_text = "\n".join(full_pages_text)

    # Deteksi Tahun Periode
    tahun = "2026"
    match_thn = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", full_text, re.IGNORECASE)
    if match_thn:
        tahun = match_thn.group(1)

    lines = full_text.split('\n')
    
    raw_blocks = []
    current_block = None
    
    # Pattern Tanggal di awal baris: DD/MM
    date_pattern = re.compile(r"^(\d{2}/\d{2})\s+(.+)")

    for line in lines:
        line_str = line.strip()
        
        # Abaikan header, footer, dan ringkasan mutasi BCA di halaman terakhir
        if not line_str:
            continue
        if any(skip_kw in line_str.upper() for skip_kw in [
            "TANGGAL KETERANGAN", "BERSAMBUNG KE", "SALDO AWAL :", "SALDO AWAL",
            "NO. REKENING", "REKENING GIRO", "HALAMAN", "PERIODE", "CATATAN:",
            "MUTASI CR :", "MUTASI DB :", "SALDO AKHIR :"
        ]):
            continue

        m_date = date_pattern.match(line_str)
        if m_date:
            if current_block:
                raw_blocks.append(current_block)
            current_block = {
                "tgl": m_date.group(1),
                "lines": [m_date.group(2)]
            }
        elif current_block:
            # Baris lanjutan deskripsi/nominal
            current_block["lines"].append(line_str)

    if current_block:
        raw_blocks.append(current_block)

    # Processing tiap blok transaksi
    for block in raw_blocks:
        combined_text = " ".join(block["lines"])
        
        # Cari semua pola angka mata uang desimal (contoh: 67,000,000.00 atau 1,315,426.17)
        # Serta cek apakah angka diikuti kata "DB"
        pattern_amount = r"([\d,]+\.\d{2})\s*(DB)?"
        matches = list(re.finditer(pattern_amount, combined_text, flags=re.IGNORECASE))
        
        if not matches:
            continue

        saldo = None
        nominal = 0.0
        is_db = False

        if len(matches) >= 2:
            # Match terakhir = SALDO
            saldo_str = matches[-1].group(1).replace(',', '')
            saldo = float(saldo_str)
            
            # Match kedua dari belakang = NOMINAL MUTASI
            mutasi_match = matches[-2]
            nominal = float(mutasi_match.group(1).replace(',', ''))
            if mutasi_match.group(2) and mutasi_match.group(2).upper() == 'DB':
                is_db = True
        else:
            # Hanya ada 1 angka (Nominal saja tanpa Saldo)
            mutasi_match = matches[0]
            nominal = float(mutasi_match.group(1).replace(',', ''))
            if mutasi_match.group(2) and mutasi_match.group(2).upper() == 'DB':
                is_db = True

        # Hapus angka-angka nominal dan teks DB dari deskripsi keterangan
        desc = combined_text
        for m in matches:
            desc = desc.replace(m.group(0), "")
        
        # Bersihkan kata DB yang tersisa & rapikan spasi
        desc = re.sub(r"\bDB\b", "", desc, flags=re.IGNORECASE)
        desc = " ".join(desc.split()).strip()

        # LOGIKA SWAP (Catatan Khusus):
        # Transaksi DB (Debet Rekening Asli)  -> Dimasukkan ke Kolom KREDIT
        # Transaksi CR (Kredit Rekening Asli) -> Dimasukkan ke Kolom DEBET
        if is_db:
            debet_swap = 0.0
            kredit_swap = nominal
        else:
            debet_swap = nominal
            kredit_swap = 0.0

        transactions.append({
            "Tanggal": f"{block['tgl']}/{tahun}",
            "Kode": "-",
            "Keterangan": desc,
            "Debit": debet_swap,
            "Kredit": kredit_swap,
            "Saldo": saldo,
            "Tipe Asli": "DB" if is_db else "CR"
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
            "Tipe Asli BCA": txn["Tipe Asli"]
        })
    return pd.DataFrame(rows)

# UI Streamlit
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
            txns = parse_bca_perfect(uploaded_file, password=password)
            df = build_df(txns, uploaded_file.name)
            
            if not df.empty:
                all_dfs.append(df)
                st.success(f"✅ {uploaded_file.name}: {len(df)} transaksi terdeteksi.")
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

        st.subheader("📊 Preview Data (Hasil Pembalikan Debet/Kredit)")
        st.dataframe(final_df, use_container_width=True)
        
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
