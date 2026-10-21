import streamlit as st
import pdfplumber
import pandas as pd
import re
import pytesseract
from PIL import Image
from pdf2image import convert_from_bytes
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (Support OCR)")

# ==========================================
# FUNGSI BANTUAN
# ==========================================
def to_float_indo_mb(s):
    if not s: return None
    try: return float(s.replace(".", "").replace(",", "."))
    except: return None

def to_float_indo_bca(s):
    if not s: return None
    try: 
        clean = re.sub(r"[^\d.-]", "", s)
        return float(clean)
    except: return None

# ==========================================
# PARSER MAYBANK
# ==========================================
AMOUNT_RE_MB = re.compile(r"(?<![\w/])(\d{1,3}(?:\.\d{3})*,\d{2})(?![\w/])")
SALDO_AWAL_RE_MB = re.compile(r"Saldo\s+Awal\s*:?\s*([\d.]+,\d{2})", re.IGNORECASE)

def parse_maybank_text(text):
    lines = text.split("\n")
    transactions = []
    date_pattern = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(\d{4})\s+(.*)")
    current_txn = None
    
    for line in lines:
        line = line.strip()
        if not line: continue
        date_match = date_pattern.match(line)
        
        if date_match:
            if current_txn: transactions.append(current_txn)
            tgl, kode, sisa = date_match.groups()
            amounts = AMOUNT_RE_MB.findall(sisa)
            desc = AMOUNT_RE_MB.sub(" ", sisa).strip()
            current_txn = {"Tanggal": tgl, "Kode": kode, "Keterangan": desc, "Nominal": amounts}
        else:
            if current_txn:
                has_valid = bool(AMOUNT_RE_MB.search(line))
                amounts = AMOUNT_RE_MB.findall(line) if has_valid else []
                desc = AMOUNT_RE_MB.sub(" ", line).strip() if has_valid else line
                if desc: current_txn["Keterangan"] += " " + desc
                if len(current_txn["Nominal"]) < 2 and amounts: current_txn["Nominal"].extend(amounts)
    if current_txn: transactions.append(current_txn)
    return transactions

def build_df_maybank(transactions, sumber_file, saldo_awal=None):
    rows = []
    prev_saldo = saldo_awal
    for txn in transactions:
        amounts = txn["Nominal"]
        nominal = amounts[0] if len(amounts) >= 2 else None
        saldo = amounts[1] if len(amounts) >= 2 else (amounts[0] if len(amounts) == 1 else None)
        
        nominal_f = to_float_indo_mb(nominal)
        saldo_f = to_float_indo_mb(saldo) if saldo else prev_saldo
        
        debit, kredit = None, None
        if saldo_f is not None and prev_saldo is not None and nominal_f is not None:
            if saldo_f > prev_saldo: debit = nominal_f
            else: kredit = nominal_f
            
        rows.append({"Sumber File": sumber_file, "Tanggal": txn["Tanggal"], "Kode": txn["Kode"], "Keterangan": txn["Keterangan"], "Debit": debit, "Kredit": kredit, "Saldo": saldo_f})
        if saldo_f is not None: prev_saldo = saldo_f
    return pd.DataFrame(rows)

# ==========================================
# PARSER BCA (Khusus OCR)
# ==========================================
def parse_bca_text(text, tahun="2026"):
    transactions = []
    
    # Cari tahun
    match_thn = re.search(r"PERIODE\s*:.*?(\d{4})", text, re.IGNORECASE)
    if match_thn: tahun = match_thn.group(1)
    else:
        match_thn2 = re.search(r"PERIODE\s*:.*?(\d{2})-(\d{2})-(\d{2})", text, re.IGNORECASE)
        if match_thn2: tahun = "20" + match_thn2.group(3)

    # 1. BUANG HEADER: Hanya proses teks setelah "SALDO AWAL"
    text_lower = text.upper()
    idx_saldo_awal = text_lower.find("SALDO AWAL")
    if idx_saldo_awal != -1:
        text = text[idx_saldo_awal:]
    
    lines = text.split("\n")
    merged_lines = []
    current_line = ""
    
    for line in lines:
        line = line.strip()
        if not line: continue
        
        # Hanya anggap baris baru jika diawali DD/MM persis
        if re.match(r"^\d{2}/\d{2}", line):
            if current_line: merged_lines.append(current_line)
            current_line = line
        else:
            if current_line: current_line += " " + line

    if current_line: merged_lines.append(current_line)

    for line in merged_lines:
        match = re.match(r"^(\d{2}/\d{2})(.*)", line)
        if not match: continue
        
        tgl_bulan = match.group(1)
        content = match.group(2).strip()
        
        # Buang jika ada teks "SALDO AWAL" di baris transaksi
        if "SALDO AWAL" in content.upper(): continue
        
        tgl = f"{tgl_bulan}/{tahun}"
        
        # Cari angka desimal BCA (lebih toleran untuk OCR)
        amounts = re.findall(r"-?\d{1,3}(?:[,.]\d{3})*[,.]\d{2}", content)
        
        debit, kredit, saldo = None, None, None
        
        if len(amounts) >= 2:
            saldo_str = amounts[-1]
            nominal_str = amounts[-2]
            
            saldo = to_float_indo_bca(saldo_str)
            nominal = to_float_indo_bca(nominal_str)
            
            if "DB" in content.upper():
                kredit = nominal 
            else:
                debit = nominal
                
        elif len(amounts) == 1:
            nominal = to_float_indo_bca(amounts[0])
            if "DB" in content.upper(): kredit = nominal
            else: debit = nominal
        
        # Bersihkan Keterangan
        desc = content
        for amt in amounts: desc = desc.replace(amt, "")
        desc = re.sub(r"\bDB\b", "", desc, flags=re.IGNORECASE)
        desc = " ".join(desc.split())

        # Buang baris transaksi jika deskripsinya kosong dan tidak ada angka
        if not desc and debit is None and kredit is None and saldo is None:
            continue

        transactions.append({
            "Tanggal": tgl,
            "Kode": "-",
            "Keterangan": desc,
            "Debit": debit,
            "Kredit": kredit,
            "Saldo": saldo
        })
            
    return transactions

def build_df_bca(transactions, sumber_file):
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

# ==========================================
# UI STREAMLIT
# ==========================================
col1, col2, col3 = st.columns(3)

with col1:
    bank_choice = st.selectbox("Pilih Bank", ["Deteksi Otomatis", "BCA", "Maybank"])
with col2:
    password = st.text_input("🔑 Password PDF (jika ada)", type="password")
with col3:
    uploaded_files = st.file_uploader("📄 Pilih file (PDF / JPG / PNG)", type=["pdf", "jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        try:
            text = ""
            
            # JIKA PDF
            if uploaded_file.name.lower().endswith(".pdf"):
                pdf_args = {"password": password} if password else {}
                with pdfplumber.open(uploaded_file, **pdf_args) as pdf:
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if not page_text or len(page_text.strip()) < 50:
                            st.info(f"📸 Memproses OCR untuk {uploaded_file.name}...")
                            images = convert_from_bytes(uploaded_file.read())
                            for img in images:
                                page_text += pytesseract.image_to_string(img, lang='ind') + "\n"
                            break 
                        
                        text += page_text + "\n"
            
            # JIKA GAMBAR
            else:
                st.info(f"📸 Memproses OCR untuk {uploaded_file.name}...")
                img = Image.open(uploaded_file)
                text = pytesseract.image_to_string(img, lang='ind')

            # TAMPILKAN TEKS OCR UNTUK DEBUGGING
            with st.expander(f"🔍 Teks Asli Hasil Scan ({uploaded_file.name})"):
                st.text(text)

            # Deteksi Bank
            detected_bank = bank_choice
            if bank_choice == "Deteksi Otomatis":
                if "REKENING GIRO" in text.upper() or "BCA" in text.upper():
                    detected_bank = "BCA"
                else:
                    detected_bank = "Maybank"

            if detected_bank == "BCA":
                txns = parse_bca_text(text)
                df = build_df_bca(txns, uploaded_file.name)
            else:
                saldo_awal = None
                match = SALDO_AWAL_RE_MB.search(text)
                if match: saldo_awal = to_float_indo_mb(match.group(1))
                txns = parse_maybank_text(text)
                df = build_df_maybank(txns, uploaded_file.name, saldo_awal)
            
            if not df.empty:
                all_dfs.append(df)
                st.success(f"✅ {uploaded_file.name}: {len(df)} transaksi terdeteksi.")
            else:
                st.warning(f"⚠️ {uploaded_file.name}: Gagal membaca transaksi. Silakan cek kotak 'Teks Asli Hasil Scan' di atas untuk melihat apa yang dibaca mesin.")
                
        except Exception as e:
            st.error(f"❌ **{uploaded_file.name}**: {type(e).__name__}")
            with st.expander("Lihat detail error"):
                st.exception(e)

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
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
