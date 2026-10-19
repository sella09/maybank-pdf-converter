import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel", layout="wide")
st.title("🏦 Bank PDF → Excel Converter")
st.write("Upload PDF Rekening Koran (Maybank / BCA), lalu unduh hasilnya.")

# ==========================================
# FUNGSI BANTUAN UMUM
# ==========================================
def to_float_indo_mb(s):
    """Konversi format Maybank (1.234.567,89) ke float"""
    if s is None or s == "": return None
    try: return float(s.replace(".", "").replace(",", "."))
    except: return None

def to_float_indo_bca(s):
    """Konversi format BCA (1,234,567.89) ke float"""
    if not s: return None
    try: 
        clean = s.replace(",", "")
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
# PARSER BCA
# ==========================================
AMOUNT_RE_BCA = re.compile(r"(-?\d{1,3}(?:,\d{3})*\.\d{2})")

def parse_bca_text(text, tahun="2026"):
    # Ambil tahun dari periode
    periode_match = re.search(r"PERIODE\s*:\s*\w+\s+(\d{4})", text, re.IGNORECASE)
    if periode_match: tahun = periode_match.group(1)
    
    # Cari Saldo Awal
    saldo_awal_match = re.search(r"SALDO AWAL\s*:\s*(-?\d{1,3}(?:,\d{3})*\.\d{2})", text, re.IGNORECASE)
    saldo_awal = to_float_indo_bca(saldo_awal_match.group(1)) if saldo_awal_match else None

    transactions = []
    
    # 1. Bersihkan header/footer
    clean_text = re.sub(r"REKENING GIRO.*?SALDO AWAL", "", text, flags=re.IGNORECASE|re.DOTALL)
    clean_text = re.sub(r"CATATAN:.*?Rekening ini", "", clean_text, flags=re.IGNORECASE|re.DOTALL)
    
    # 2. Pecah teks berdasarkan kemunculan Tanggal (dd/mm)
    parts = re.split(r"(?=\b\d{2}/\d{2})", clean_text)
    
    for part in parts:
        part = part.strip()
        if not part: continue
        
        date_match = re.match(r"^(\d{2}/\d{2})(.*)", part, re.DOTALL)
        if not date_match: continue
        
        tgl_bulan = date_match.group(1)
        content = date_match.group(2)
        
        tgl = f"{tgl_bulan}/{tahun}"
        
        # Cari semua angka mutasi
        numbers = AMOUNT_RE_BCA.findall(part)
        
        nominal_str = None
        saldo_str = None
        
        if len(numbers) >= 2:
            nominal_str = numbers[-2]
            saldo_str = numbers[-1]
        elif len(numbers) == 1:
            nominal_str = numbers[0]
            
        nominal = to_float_indo_bca(nominal_str)
        saldo = to_float_indo_bca(saldo_str) if saldo_str else None
        
        # Tentukan Debit/Kredit
        debit = None
        kredit = None
        if nominal is not None:
            if " DB" in part or "DB " in part or "DB- " in part:
                debit = nominal
            else:
                kredit = nominal
                
        # Bersihkan Keterangan
        desc = part
        if nominal_str: desc = desc.replace(nominal_str, "")
        if saldo_str: desc = desc.replace(saldo_str, "")
        desc = re.sub(r"\bDB\b", "", desc)
        desc = desc.replace("\n", " ").strip()
        
        transactions.append({
            "Tanggal": tgl,
            "Kode": "-", 
            "Keterangan": desc,
            "Debit": debit,
            "Kredit": kredit,
            "Saldo": saldo
        })
    
    # Fallback saldo jika kosong
    prev_saldo = saldo_awal
    for txn in transactions:
        if txn["Saldo"] is None and prev_saldo is not None:
            if txn["Debit"]: prev_saldo -= txn["Debit"]
            elif txn["Kredit"]: prev_saldo += txn["Kredit"]
            txn["Saldo"] = prev_saldo
        elif txn["Saldo"] is not None:
            prev_saldo = txn["Saldo"]
            
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
col1, col2 = st.columns(2)
with col1:
    bank_choice = st.selectbox("Pilih Bank", ["Deteksi Otomatis", "Maybank", "BCA"])
    password = st.text_input("🔑 Password PDF (jika ada)", type="password")
with col2:
    uploaded_files = st.file_uploader("📄 Pilih file PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        try:
            pdf_args = {"password": password} if password else {}
            all_text = ""
            with pdfplumber.open(uploaded_file, **pdf_args) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text: all_text += text + "\n"
            
            # Deteksi Bank
            detected_bank = bank_choice
            if bank_choice == "Deteksi Otomatis":
                if "REKENING GIRO" in all_text and "BCA" in all_text:
                    detected_bank = "BCA"
                elif "Maybank" in all_text or "MOTOG" in all_text:
                    detected_bank = "Maybank"
                else:
                    st.warning(f"⚠️ Bank tidak terdeteksi untuk {uploaded_file.name}, mencoba Maybank sebagai default.")
                    detected_bank = "Maybank"

            st.info(f"Memproses **{uploaded_file.name}** sebagai **{detected_bank}**...")

            if detected_bank == "Maybank":
                saldo_awal = None
                match = SALDO_AWAL_RE_MB.search(all_text)
                if match: saldo_awal = to_float_indo_mb(match.group(1))
                txns = parse_maybank_text(all_text)
                df = build_df_maybank(txns, uploaded_file.name, saldo_awal)
            elif detected_bank == "BCA":
                txns = parse_bca_text(all_text)
                df = build_df_bca(txns, uploaded_file.name)
            
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
