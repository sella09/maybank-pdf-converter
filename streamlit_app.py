import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA & Maybank)")

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
# PARSER BCA (Versi 19.0)
# ==========================================
def parse_bca_final(file_obj, password=None):
    transactions = []
    pdf_args = {"password": password} if password else {}
    
    with pdfplumber.open(file_obj, **pdf_args) as pdf:
        tahun = "2026"
        page1_text = pdf.pages[0].extract_text() or ""
        match_thn = re.search(r"PERIODE\s*:\s*\w+\s+(\d{4})", page1_text, re.IGNORECASE)
        if match_thn: tahun = match_thn.group(1)

        for page in pdf.pages:
            words = page.extract_words()
            
            lines = {}
            for w in words:
                y_key = round(w['top'] / 5) * 5
                if y_key not in lines:
                    lines[y_key] = []
                lines[y_key].append(w)

            for y, word_list in sorted(lines.items()):
                word_list.sort(key=lambda w: w['x0'])
                full_text = " ".join([w['text'] for w in word_list])
                
                date_match = re.search(r"(\d{2}/\d{2})", full_text)
                if not date_match: continue
                
                tgl_bulan = date_match.group(1)
                numbers = re.findall(r"-?\d{1,3}(?:[,.]\d{3})*\.\d{2}", full_text)
                
                desc = full_text
                desc = desc.replace(tgl_bulan, "") 
                for n in numbers: desc = desc.replace(n, "") 
                desc = re.sub(r"\bDB\b", "", desc, flags=re.IGNORECASE) 
                
                desc = desc.replace("TRS F", "TRSF").replace("E-B ANKING", "E-BANKING").replace("SW ITCHING", "SWITCHING")
                desc = " ".join(desc.split()).strip()

                debit, kredit, saldo = None, None, None
                
                if len(numbers) >= 2:
                    nominal = to_float_indo_bca(numbers[-2])
                    saldo = to_float_indo_bca(numbers[-1])
                    if "DB" in full_text.upper(): kredit = nominal
                    else: debit = nominal
                elif len(numbers) == 1:
                    nominal = to_float_indo_bca(numbers[0])
                    if "DB" in full_text.upper(): kredit = nominal
                    else: debit = nominal

                transactions.append({
                    "Tanggal": f"{tgl_bulan}/{tahun}",
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
                txns = parse_bca_final(uploaded_file, password=password)
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
