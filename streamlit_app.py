import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel Converter", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA & Maybank - Precision Fix)")

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
    
    ignore_keywords = [
        "REKENING GIRO", "KCU", "DOMINIC DIGITAL", "KELAPA GADING", "JL BOULEVARD",
        "JAKARTA", "INDONESIA", "NO. REKENING", "HALAMAN", "PERIODE", "MATA UANG",
        "CATATAN:", "Apabila nasabah", "BCA berhak", "TANGGAL KETERANGAN", 
        "Bersambung ke halaman", "SALDO AWAL :", "MUTASI CR :", "MUTASI DB :", "SALDO AKHIR :"
    ]

    clean_lines = []
    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue
        if re.match(r"^\d+\s*/\s*\d+$", line_str) or (re.match(r"^\d+$", line_str) and len(line_str) <= 3):
            continue
        if any(kw.lower() in line_str.lower() for kw in ignore_keywords):
            continue
        clean_lines.append(line_str)

    raw_blocks = []
    current_block = None
    date_pattern = re.compile(r"^(\d{2}/\d{2})\s+(.+)")

    for line_str in clean_lines:
        m_date = date_pattern.match(line_str)
        if m_date:
            if current_block:
                raw_blocks.append(current_block)
            current_block = {
                "tgl": m_date.group(1),
                "lines": [m_date.group(2)]
            }
        elif current_block:
            current_block["lines"].append(line_str)

    if current_block:
        raw_blocks.append(current_block)

    for block in raw_blocks:
        combined_text = " ".join(block["lines"])
        
        # Ekstrak angka yang memiliki koma pemisah ribuan yang valid (format mata uang resmi)
        pattern_amount = r"(\d{1,3}(?:,\d{3})+\.\d{2})\s*(DB)?"
        matches = list(re.finditer(pattern_amount, combined_text, flags=re.IGNORECASE))
        
        if not matches:
            continue

        saldo = None
        nominal = 0.0
        is_db = False

        upper_text = combined_text.upper()
        if "PAJAK BUNGA" in upper_text or "BIAYA ADM" in upper_text:
            db_match = re.search(r"(\d{1,3}(?:,\d{3})+\.\d{2})\s+DB", combined_text, flags=re.IGNORECASE)
            if db_match:
                nominal = float(db_match.group(1).replace(',', ''))
                is_db = True
            saldo_str = matches[-1].group(1).replace(',', '')
            saldo = float(saldo_str)
        elif "BUNGA" in upper_text:
            bunga_match = re.search(r"(\d{1,3}(?:,\d{3})+\.\d{2})", combined_text)
            if bunga_match:
                nominal = float(bunga_match.group(1).replace(',', ''))
            is_db = False
            if len(matches) >= 2:
                saldo_str = matches[-1].group(1).replace(',', '')
                saldo = float(saldo_str)
        else:
            # Transaksi normal: 
            # Jika ada 2 atau lebih angka berformat koma, angka terakhir = Saldo, angka sebelumnya = Mutasi
            if len(matches) >= 2:
                saldo_str = matches[-1].group(1).replace(',', '')
                saldo = float(saldo_str)
                
                mutasi_match = matches[-2]
                nominal = float(mutasi_match.group(1).replace(',', ''))
                if mutasi_match.group(2) and mutasi_match.group(2).upper() == 'DB':
                    is_db = True
            else:
                mutasi_match = matches[0]
                nominal = float(mutasi_match.group(1).replace(',', ''))
                if mutasi_match.group(2) and mutasi_match.group(2).upper() == 'DB':
                    is_db = True

        desc = combined_text
        for m in matches:
            desc = desc.replace(m.group(0), "")
        
        desc = re.sub(r"\bDB\b", "", desc, flags=re.IGNORECASE)
        desc = " ".join(desc.split()).strip()

        # LOGIKA PEMBALIKAN (SWAP DEBET <-> KREDIT):
        # Sesuai permintaan: Nominal Debet dipindah ke Kredit, dan Kredit ke Debet.
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

def parse_maybank_statement(file_obj, password=None):
    transactions = []
    pdf_args = {"password": password} if password else {}
    
    with pdfplumber.open(file_obj, **pdf_args) as pdf:
        full_text = ""
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    lines = full_text.split('\n')
    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue
        
        match_date = re.search(r"(\d{2}[/-]\d{2}[/-]\d{4}|\d{2}[/-]\d{2})", line_str)
        if match_date:
            transactions.append({
                "Tanggal": match_date.group(1),
                "Kode": "-",
                "Keterangan": line_str,
                "Debit": 0.0,
                "Kredit": 0.0,
                "Saldo": None,
                "Tipe Asli": "CR"
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
            "Tipe Asli Bank": txn["Tipe Asli"]
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
            
            if bank_choice == "BCA":
                txns = parse_bca_perfect(uploaded_file, password=password)
            else:
                txns = parse_maybank_statement(uploaded_file, password=password)
                
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
        
        # Pengurutan Kronologis dari Bulan Terkecil (Januari) ke Terbesar
        try:
            final_df['Tanggal'] = final_df['Tanggal'].astype(str).str.strip()
            final_df['Tanggal_DT'] = pd.to_datetime(final_df['Tanggal'], format='%d/%m/%Y', errors='coerce')
            final_df = final_df.sort_values(by=['Tanggal_DT'], ascending=True).reset_index(drop=True)
            final_df = final_df.drop(columns=['Tanggal_DT'])
        except Exception as e:
            st.warning(f"Gagal mengurutkan tanggal secara otomatis: {e}")

        st.subheader("📊 Preview Data (Presisi & Urut Kronologis)")
        st.dataframe(final_df, use_container_width=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Mutasi_Swapped')
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 Download Excel Presisi & Urut",
            data=processed_data,
            file_name="mutasi_bank_swapped.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
