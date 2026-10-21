import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel", layout="wide")
st.title("🏦 Bank PDF → Excel Converter")

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
# PARSER BCA (KOORDINAT & PERBAIKAN DEBIT/KREDIT)
# ==========================================
def parse_bca_coordinates(file_obj, password=None):
    transactions = []
    pdf_args = {"password": password} if password else {}
    
    with pdfplumber.open(file_obj, **pdf_args) as pdf:
        tahun = "2026"
        page1_text = pdf.pages[0].extract_text() or ""
        match_thn = re.search(r"PERIODE\s*:\s*\w+\s+(\d{4})", page1_text, re.IGNORECASE)
        if match_thn: tahun = match_thn.group(1)

        for page in pdf.pages:
            words = page.extract_words()
            
            # Cari batas Y (bawah header tabel)
            start_y = 0
            for w in words:
                if "TANGGAL" in w['text'].upper():
                    start_y = w['bottom'] + 5
                    break
            
            # Bagi halaman menjadi 3 zona X
            page_width = page.width
            x_split_saldo = page_width * 0.75 
            x_split_mutasi = page_width * 0.65
            
            # Kelompokkan kata berdasarkan baris (Y)
            lines = {}
            for w in words:
                if w['top'] < start_y: continue
                y_key = round(w['top'] / 5) * 5 
                if y_key not in lines:
                    lines[y_key] = []
                lines[y_key].append(w)

            # Proses per baris dari atas ke bawah
            for y, word_list in sorted(lines.items()):
                word_list.sort(key=lambda w: w['x0'])
                
                text_kiri = ""
                text_tengah = ""
                text_kanan = ""
                
                for w in word_list:
                    if w['x0'] < x_split_mutasi:
                        text_kiri += w['text'] + " "
                    elif w['x0'] < x_split_saldo:
                        text_tengah += w['text'] + " "
                    else:
                        text_kanan += w['text'] + " "
                
                # Apakah baris ini tanggal transaksi baru?
                match = re.match(r"^(\d{2}/\d{2})(.*)", text_kiri.strip())
                
                if match:
                    if 'current_txn' in locals() and current_txn is not None:
                        transactions.append(current_txn)
                    
                    tgl_bulan = match.group(1)
                    desc = match.group(2).strip()
                    
                    if "SALDO AWAL" in desc.upper():
                        current_txn = None
                        continue
                        
                    desc = " ".join(desc.split())
                    
                    current_txn = {
                        "Tanggal": f"{tgl_bulan}/{tahun}",
                        "Kode": "-",
                        "Keterangan": desc,
                        "Debit": None,
                        "Kredit": None,
                        "Saldo": None
                    }
                    
                    all_angka_mutasi = text_tengah + text_kiri
                    
                    # Cari Saldo di kolom kanan
                    match_saldo = re.search(r"(-?\d{1,3}(?:[,.]\d{3})*\.\d{2})", text_kanan)
                    if match_saldo:
                        current_txn["Saldo"] = to_float_indo_bca(match_saldo.group(1))
                    
                    # Cari Nominal Mutasi
                    match_mutasi = re.search(r"(\d{1,3}(?:[,.]\d{3})*\.\d{2})", all_angka_mutasi)
                    if match_mutasi:
                        nominal = to_float_indo_bca(match_mutasi.group(1))
                        
                        # --- PERBAIKAN: Posisi Debit/Kredit DITUKAR ---
                        if "DB" in text_tengah.upper() or "DB" in text_kiri.upper():
                            current_txn["Kredit"] = nominal
                        else:
                            current_txn["Debit"] = nominal
                        # ----------------------------------------------
                            
                        current_txn["Keterangan"] = current_txn["Keterangan"].replace(match_mutasi.group(1), "")
                            
                elif 'current_txn' in locals() and current_txn is not None:
                    # Baris lanjutan (teks deskripsi di bawahnya)
                    desc_lanjutan = " ".join(text_kiri.split())
                    if desc_lanjutan and not re.match(r"^-?\d{1,3}(?:[,.]\d{3})*\.\d{2}$", desc_lanjutan):
                        current_txn["Keterangan"] += " " + desc_lanjutan

        if 'current_txn' in locals() and current_txn is not None:
            transactions.append(current_txn)
            
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
    uploaded_files = st.file_uploader("📄 Pilih file PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        try:
            detected_bank = bank_choice
            if bank_choice == "Deteksi Otomatis":
                with pdfplumber.open(uploaded_file, password=password if password else "") as pdf:
                    first_page_text = pdf.pages[0].extract_text() or ""
                
                if "REKENING GIRO" in first_page_text and "BCA" in first_page_text:
                    detected_bank = "BCA"
                elif "Maybank" in first_page_text:
                    detected_bank = "Maybank"
                else:
                    detected_bank = "Maybank" 

            st.info(f"Memproses **{uploaded_file.name}** sebagai **{detected_bank}**...")

            if detected_bank == "Maybank":
                with pdfplumber.open(uploaded_file, password=password if password else "") as pdf:
                    all_text = ""
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text: all_text += text + "\n"
                
                saldo_awal = None
                match = SALDO_AWAL_RE_MB.search(all_text)
                if match: saldo_awal = to_float_indo_mb(match.group(1))
                txns = parse_maybank_text(all_text)
                df = build_df_maybank(txns, uploaded_file.name, saldo_awal)
                
            elif detected_bank == "BCA":
                txns = parse_bca_coordinates(uploaded_file, password=password)
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
