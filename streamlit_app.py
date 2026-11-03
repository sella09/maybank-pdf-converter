import streamlit as st
import pdfplumber
import pandas as pd
import re
import traceback
from io import BytesIO

# =========================================================
# KONFIGURASI
# =========================================================
st.set_page_config(page_title="Bank PDF to Excel Converter", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA & Maybank)")

# =========================================================
# HELPER
# =========================================================
def open_pdf(file_obj, password=None):
    file_obj.seek(0)
    try:
        if password:
            return pdfplumber.open(file_obj, password=password)
        return pdfplumber.open(file_obj)
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg or "decrypt" in msg:
            if not password:
                raise ValueError("PDF terproteksi password. Isi kolom 🔑 Password PDF.")
            raise ValueError("Password PDF salah.")
        raise

def to_float(num_str):
    if num_str is None: return None
    s = str(num_str).strip()
    if not s or s in ("-", ""): return None
    s = re.sub(r"[^\d.,]", "", s)
    if not s: return None
    if re.search(r",\d{2}$", s) and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try: return float(s)
    except ValueError: return None

def clean_desc(text):
    if not text: return ""
    t = str(text)
    t = re.sub(r"Rekening ini sampai dengan akhir bulan berikutnya.*?Rekening ini\.", "", t, flags=re.I | re.S)
    t = re.sub(r"telah menyetujui segala data.*?Rekening ini\.", "", t, flags=re.I | re.S)
    t = re.sub(r"Apabila nasabah tidak melakukan sanggahan.*?Laporan Mutasi Rekening\.", "", t, flags=re.I | re.S)
    return " ".join(t.split()).strip()

# =========================================================
# PARSER BCA (LINE-BY-LINE)
# =========================================================
@st.cache_data(show_spinner=False)
def parse_bca(file_bytes, password=None):
    transactions = []
    file_obj = BytesIO(file_bytes)
    pdf = open_pdf(file_obj, password)
    
    full_text = ""
    for page in pdf.pages:
        t = page.extract_text()
        if t: full_text += t + "\n"
    pdf.close()

    tahun = "2026"
    m = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", full_text, re.IGNORECASE)
    if m: tahun = m.group(1)

    # --- LANGKAH 1: DETEKSI PAJAK BUNGA SECARA GLOBAL ---
    # Karena di PDF BCA, PAJAK BUNGA sering menempel dengan BUNGA di satu baris.
    # Kita deteksi dulu dengan regex global, lalu hapus dari teks agar tidak dobel.
    pajak_pattern = re.compile(
        r"BUNGA\s+([\d,.]+)\s+(?:CR|DB)?\s+PAJAK\s+BUNGA\s+([\d,.]+)\s+(?:CR|DB)?",
        re.IGNORECASE
    )
    
    for match in pajak_pattern.finditer(full_text):
        bunga_val = to_float(match.group(1))
        pajak_val = to_float(match.group(2))
        
        # Cari tanggal terdekat sebelum match
        start_pos = match.start()
        context_before = full_text[max(0, start_pos - 500):start_pos]
        date_matches = re.findall(r"\b(\d{2}/\d{2})\b", context_before)
        tgl_pajak = date_matches[-1] if date_matches else None
        if not tgl_pajak:
            continue
        
        if bunga_val:
            transactions.append({
                "Tanggal": f"{tgl_pajak}/{tahun}", "Kode": "-",
                "Keterangan": "BUNGA", "Debit": bunga_val, "Kredit": 0.0,
                "Saldo": None, "Tipe Asli": "CR",
            })
        if pajak_val:
            transactions.append({
                "Tanggal": f"{tgl_pajak}/{tahun}", "Kode": "-",
                "Keterangan": "PAJAK BUNGA", "Debit": 0.0, "Kredit": pajak_val,
                "Saldo": None, "Tipe Asli": "DB",
            })

    # --- LANGKAH 2: PARSE BARIS PER BARIS (TRANSAKSI NORMAL + BIAYA ADM) ---
    lines = full_text.split('\n')
    
    ignore_keywords = [
        "REKENING GIRO", "KCU", "DOMINIC DIGITAL", "KELAPA GADING", "JL BOULEVARD",
        "JAKARTA", "INDONESIA", "NO. REKENING", "HALAMAN", "PERIODE", "MATA UANG",
        "CATATAN:", "Apabila nasabah", "BCA berhak", "TANGGAL KETERANGAN",
        "Bersambung ke halaman", "MUTASI CR :", "MUTASI DB :", "SALDO AKHIR :",
        "===== Page", "# REKENING", "Page ", "- Page "
    ]
    
    clean_lines = []
    for line in lines:
        line_str = line.strip()
        if not line_str: continue
        if re.match(r"^\d+\s*/\s*\d+$", line_str): continue
        if re.match(r"^\d+$", line_str) and len(line_str) <= 3: continue
        if any(kw.lower() in line_str.lower() for kw in ignore_keywords): continue
        # Skip baris yang sudah ditangkap pola PAJAK BUNGA global
        if re.search(r"PAJAK\s+BUNGA", line_str, re.IGNORECASE): continue
        clean_lines.append(line_str)

    raw_blocks = []
    current_block = None
    date_pattern = re.compile(r"^(\d{2}/\d{2})\s*(.+)")

    for line_str in clean_lines:
        m_date = date_pattern.match(line_str)
        if m_date:
            if current_block: raw_blocks.append(current_block)
            current_block = {"tgl": m_date.group(1), "lines": [m_date.group(2)]}
        elif current_block:
            current_block["lines"].append(line_str)

    if current_block: raw_blocks.append(current_block)

    for block in raw_blocks:
        combined = " ".join(block["lines"])
        if "SALDO AWAL" in combined.upper(): continue

        upper = combined.upper()
        nums = re.findall(r"\d[\d,]*\.\d{2}", combined)
        if not nums: continue
        
        num_vals = [to_float(n) for n in nums if to_float(n) is not None]

        # --- BUNGA (biasanya sudah ditangkap pola global, tapi ini untuk jaga-jaga) ---
        if "BUNGA" in upper and "PAJAK" not in upper:
            bunga_val = num_vals[0] if num_vals else 0.0
            transactions.append({
                "Tanggal": f"{block['tgl']}/{tahun}", "Kode": "-",
                "Keterangan": "BUNGA", "Debit": bunga_val, "Kredit": 0.0,
                "Saldo": None, "Tipe Asli": "CR",
            })
            continue

        # --- BIAYA ADM / TRANSFER / REFERENSI ---
        if "BIAYA ADM" in upper or "BIAYA TRANSFER" in upper or "BIAYA REFERENSI" in upper:
            biaya_val = num_vals[0] if num_vals else 0.0
            saldo_val = None
            if len(num_vals) >= 2:
                saldo_val = num_vals[-1]
                if saldo_val == biaya_val:
                    saldo_val = None
            transactions.append({
                "Tanggal": f"{block['tgl']}/{tahun}", "Kode": "-",
                "Keterangan": "BIAYA ADM", "Debit": 0.0, "Kredit": biaya_val,
                "Saldo": saldo_val, "Tipe Asli": "DB",
            })
            continue

        # --- TRANSAKSI NORMAL ---
        is_db = False
        if re.search(r"\bDB\b", upper): is_db = True
        elif any(k in upper for k in ["EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET", "PENERIMAAN NEGARA"]): is_db = True

        nominal = num_vals[0] if num_vals else 0.0
        saldo = None
        if len(num_vals) >= 2:
            saldo_candidate = num_vals[-1]
            if saldo_candidate > 100 and saldo_candidate != nominal:
                saldo = saldo_candidate

        desc = combined
        for n in nums: desc = desc.replace(n, "")
        desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
        desc = clean_desc(desc)

        debit_val, kredit_val = (0.0, nominal) if is_db else (nominal, 0.0)
        transactions.append({
            "Tanggal": f"{block['tgl']}/{tahun}", "Kode": "-", "Keterangan": desc,
            "Debit": debit_val, "Kredit": kredit_val,
            "Saldo": saldo, "Tipe Asli": "DB" if is_db else "CR",
        })

    return transactions

# =========================================================
# PARSER MAYBANK
# =========================================================
@st.cache_data(show_spinner=False)
def parse_maybank(file_bytes, password=None):
    transactions = []
    file_obj = BytesIO(file_bytes)
    pdf = open_pdf(file_obj, password)
    full_text = "".join([(p.extract_text() or "") + "\n" for p in pdf.pages])
    pdf.close()
    
    lines = full_text.split("\n")
    date_re = re.compile(r"^(\d{2}[/-]\d{2}[/-]\d{4}|\d{2}[/-]\d{2})\b")

    for line in lines:
        line_str = line.strip()
        if not line_str: continue
        m = date_re.match(line_str)
        if not m: continue
        tgl = m.group(1)
        
        nums = re.findall(r"\d[\d,]*\.\d{2}", line_str)
        parsed = [to_float(n) for n in nums if to_float(n) is not None]
        if not parsed: continue
        
        nominal = parsed[0]
        saldo = parsed[-1] if len(parsed) >= 2 and parsed[-1] != nominal and parsed[-1] > 100 else None
        is_db = bool(re.search(r"\bDB\b", line_str.upper()))
        
        desc = re.sub(r"^\d{2}[/-]\d{2}[/-]?\d{0,4}\s*", "", line_str)
        for n in nums: desc = desc.replace(n, "")
        desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
        desc = clean_desc(desc)

        debit_val, kredit_val = (0.0, nominal) if is_db else (nominal, 0.0)
        transactions.append({
            "Tanggal": tgl, "Kode": "-", "Keterangan": desc,
            "Debit": debit_val, "Kredit": kredit_val,
            "Saldo": saldo, "Tipe Asli": "DB" if is_db else "CR",
        })
    return transactions

# =========================================================
# UI
# =========================================================
col1, col2, col3 = st.columns(3)
with col1: 
    bank_choice = st.selectbox("Pilih Bank", ["BCA", "Maybank"])
with col2: 
    password = st.text_input("🔑 Password PDF (jika ada)", type="password")
with col3: 
    uploaded_files = st.file_uploader("📄 Pilih file PDF/TXT", type=["pdf", "txt"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    
    for uploaded_file in uploaded_files:
        try:
            st.info(f"Memproses **{uploaded_file.name}**...")
            file_bytes = uploaded_file.read()
            
            if bank_choice == "BCA":
                txns = parse_bca(file_bytes, password=password)
            else:
                txns = parse_maybank(file_bytes, password=password)

            df = pd.DataFrame([{
                "Sumber File": uploaded_file.name, "Tanggal": t["Tanggal"], "Kode": t["Kode"],
                "Keterangan": t["Keterangan"], "Debit": t["Debit"], "Kredit": t["Kredit"],
                "Saldo": t["Saldo"], "Tipe Asli Bank": t["Tipe Asli"]
            } for t in txns])
            
            if not df.empty:
                all_dfs.append(df)
                st.success(f"✅ {uploaded_file.name}: {len(df)} transaksi.")
            else:
                st.warning(f"⚠️ {uploaded_file.name}: Tidak ada transaksi.")
                
        except ValueError as ve:
            st.error(f"🔒 {uploaded_file.name}: {ve}")
        except Exception as e:
            st.error(f"❌ Gagal memproses {uploaded_file.name}: {str(e)}")
            with st.expander("Detail Error"):
                st.code(traceback.format_exc())

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        final_df["Tanggal"] = final_df["Tanggal"].astype(str).str.strip()
        final_df["Tanggal_DT"] = pd.to_datetime(final_df["Tanggal"], format="%d/%m/%Y", errors="coerce")
        final_df = final_df.sort_values("Tanggal_DT").reset_index(drop=True)
        final_df = final_df.drop(columns=["Tanggal_DT"])
        final_df = final_df[(final_df["Debit"] > 0) | (final_df["Kredit"] > 0)]

        st.subheader("📊 Preview Data")
        st.dataframe(final_df, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Mutasi")
        
        st.download_button(
            "📥 Download Excel",
            data=output.getvalue(),
            file_name="mutasi_bank_correct.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
