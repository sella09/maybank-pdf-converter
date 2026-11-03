import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel Converter", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA & Maybank - Absolute Fix)")

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
        "Bersambung ke halaman", "MUTASI CR :", "MUTASI DB :", "SALDO AKHIR :"
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
        
        if "SALDO AWAL" in combined_text.upper():
            continue

        # Ekstrak semua angka desimal resmi (mendukung angka kecil seperti pajak bunga & angka besar)
        pattern_amount = r"([\d,]+\.\d{2})\s*(DB)?\s*(CR)?"
        matches = list(re.finditer(pattern_amount, combined_text, flags=re.IGNORECASE))
        
        if not matches:
            continue

        saldo = None
        nominal = 0.0
        is_db = False
        upper_text = combined_text.upper()

        if " DB" in upper_text or "DB " in upper_text or upper_text.endswith("DB"):
            is_db = True

        # 1. BARIS PAJAK BUNGA & BIAYA ADM
        if "PAJAK BUNGA" in upper_text or "BIAYA ADM" in upper_text:
            is_db = True
            db_match = re.search(r"([\d,]+\.\d{2})\s+DB", combined_text, flags=re.IGNORECASE)
            if db_match:
                nominal = float(db_match.group(1).replace(',', ''))
            elif matches:
                nominal = float(matches[0].group(1).replace(',', ''))
            
            if len(matches) >= 2:
                saldo = float(matches[-1].group(1).replace(',', ''))
        
        # 2. BARIS BUNGA
        elif "BUNGA" in upper_text and "PAJAK" not in upper_text:
            is_db = False
            if matches:
                nominal = float(matches[0].group(1).replace(',', ''))
            if len(matches) >= 2:
                saldo = float(matches[-1].group(1).replace(',', ''))

        # 3. TRANSAKSI NORMAL BCA
        else:
            parsed_items = []
            for m in matches:
                val = float(m.group(1).replace(',', ''))
                has_db_tag = bool(m.group(2) and m.group(2).upper() == 'DB')
                parsed_items.append({"val": val, "is_db": has_db_tag, "raw": m.group(0)})

            if len(parsed_items) >= 2:
                # Algoritma Presisi BCA: 
                # Saldo biasanya adalah angka terbesar atau angka terakhir di baris. 
                # Namun jika ada 2 angka, angka mutasi adalah yang menempel dengan tag DB atau berada di posisi sebelum saldo.
                # Kita pisahkan berdasarkan ukuran / posisi:
                val1 = parsed_items[0]["val"]
                val2 = parsed_items[1]["val"]
                
                if parsed_items[0]["is_db"] or parsed_items[1]["is_db"] or is_db:
                    # Jika ada tag DB, pastikan nominal yang ber-tag DB diambil sebagai mutasi
                    if parsed_items[0]["is_db"]:
                        nominal = val1
                        saldo = val2
                        is_db = True
                    elif parsed_items[1]["is_db"] or is_db:
                        # Jika angka kedua ber-tag DB atau baris DB, angka pertama adalah saldo (atau sebaliknya)
                        # Di BCA, mutasi ada di tengah, saldo di ujung kanan.
                        nominal = val1
                        saldo = val2
                        is_db = True
                else:
                    # Transaksi CR standar: Angka pertama mutasi, angka kedua saldo
                    nominal = val1
                    saldo = val2
            elif len(parsed_items) == 1:
                nominal = parsed_items[0]["val"]
                if parsed_items[0]["is_db"]:
                    is_db = True
                saldo = None

        # Bersihkan deskripsi keterangan
        desc = combined_text
        for m in matches:
            desc = desc.replace(m.group(0), "")
        
        desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.IGNORECASE)
        desc = " ".join(desc.split()).strip()

        # ==========================================
        # LOGIKA SWAP (DEBET <-> KREDIT):
        # - Asli DB (Pengeluaran) -> Masuk KREDIT
        # - Asli CR (Pemasukan)   -> Masuk DEBET
        # ==========================================
        if is_db:
            debit_val = 0.0
            kredit_val = nominal
        else:
            debit_val = nominal
            kredit_val = 0.0

        transactions.append({
            "Tanggal": f"{block['tgl']}/{tahun}",
            "Kode": "-",
            "Keterangan": desc,
            "Debit": debit_val,
            "Kredit": kredit_val,
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
