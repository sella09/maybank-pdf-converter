import streamlit as st
import pdfplumber
import pandas as pd
import re
import traceback
from io import BytesIO

# =========================================================
# KONFIGURASI HALAMAN
# =========================================================
st.set_page_config(page_title="Bank PDF to Excel Converter", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA & Maybank)")

# =========================================================
# HELPER FUNCTIONS
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
    return " ".join(t.split()).strip()

# =========================================================
# PARSER BCA (SALDO PRESISI & PAJAK BUNGA AMAN)
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

    # 1. Tahun & Saldo Awal
    tahun = "2026"
    m = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", full_text, re.IGNORECASE)
    if m: tahun = m.group(1)

    saldo_awal = 0.0
    m_saldo = re.search(r"SALDO\s+AWAL\s*[\n:]?\s*([\d,.]+)", full_text, re.IGNORECASE)
    if m_saldo: 
        saldo_awal = to_float(m_saldo.group(1)) or 0.0

    # 2. Pembersihan Teks Header (Tetap pertahankan BUNGA & PAJAK BUNGA)
    full_text = re.sub(r"KELAPA GADING HALAMAN : \d+ /\d+ KELAPA GADING BARAT JL BOULEVARD RAYA QF 1 NO 7 INDONESIA TANGGAL KETERANGAN CBG MUTASI SALDO", "", full_text, flags=re.IGNORECASE)
    full_text = re.sub(r"KELAPA GADING HALAMAN : \d+ / \d+ KELAPA GADING BARAT JL BOULEVARD RAYA QF 1 NO 7 INDONESIA TANGGAL KETERANGAN CBG MUTASI SALDO", "", full_text, flags=re.IGNORECASE)
    full_text = re.sub(r"Bersambung ke halaman berikut", "", full_text, flags=re.IGNORECASE)

    # 3. Gabungkan Baris
    lines = full_text.split('\n')
    clean_lines = []
    for line in lines:
        s = line.strip()
        if not s: continue
        if any(k in s for k in ["REKENING GIRO", "KCU", "DOMINIC DIGITAL", "NO. REKENING", "PERIODE", "MATA UANG", "CATATAN:", "BCA berhak"]):
            continue
        # Filter BUNGA dihapus agar PAJAK BUNGA muncul
        clean_lines.append(s)
        
    merged_lines = []
    for s in clean_lines:
        if re.match(r"^\d{2}/\d{2}", s):
            merged_lines.append(s)
        elif merged_lines:
            merged_lines[-1] += " " + s

    # 4. Parsing Transaksi
    for line in merged_lines:
        m_date = re.match(r"^(\d{2}/\d{2})(.*)", line)
        if not m_date: continue
        
        tgl = m_date.group(1)
        rest = m_date.group(2).strip()

        if "SALDO AWAL" in rest.upper() or "SALDO AKHIR" in rest.upper(): continue

        nums = re.findall(r"\d[\d,]*\.\d{2}", rest)
        if not nums: continue
        
        nominal = to_float(nums[0])
        if nominal is None: continue

        # Prioritas ambil saldo dari PDF (angka terakhir di baris)
        saldo = to_float(nums[-1]) if len(nums) > 1 else None
        if saldo is not None and (saldo < 0 or saldo > 9999999999):
            saldo = None

        upper_rest = rest.upper()
        
        # DETEKSI DB / CR
        is_db = False
        
        # Uang Masuk (CR)
        cr_markers = ["SETORAN TUNAI", "BI-FAST CR", "TRSF E-BANKING CR", "SWITCHING CR", "KR OTOMATIS", "BUNGA", "REFUND", "PENERIMAAN NEGARA"]
        # Uang Keluar (DB)
        db_markers = [
            "EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET", "TIGA MUSTIKA KONSU", "PT.FINNET", "PT. FINNET", 
            "BA JASA E-BANKING", "BIAYA ADM", "BIAYA TRANSFER", "BIAYA REFERENSI", "PAJAK BUNGA",
            "PEMBELIAN", "BY PASANG", "BAYAR", "PENGELUARAN", "SWITCHING DB", "TRSF E-BANKING DB"
        ]

        if any(k in upper_rest for k in db_markers):
            is_db = True
        elif any(k in upper_rest for k in cr_markers):
            is_db = False
        elif "DB" in upper_rest:
            is_db = True
        elif "CR" in upper_rest:
            is_db = False

        desc = rest
        for n in nums:
            desc = desc.replace(n, "")
        desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
        desc = clean_desc(desc)

        debit_val, kredit_val = (0.0, nominal) if is_db else (nominal, 0.0)

        transactions.append({
            "Tanggal": f"{tgl}/{tahun}", 
            "Kode": "-", 
            "Keterangan": desc,
            "Debit": debit_val, 
            "Kredit": kredit_val,
            "Saldo": saldo,
            "Tipe Asli": "DB" if is_db else "CR",
        })

    # 5. HITUNG RUNNING BALANCE (Hanya untuk baris yang saldonya kosong)
    transactions.sort(key=lambda x: x["Tanggal"])
    
    saldo_berjalan = saldo_awal
    for txn in transactions:
        if txn["Saldo"] is not None:
            saldo_berjalan = txn["Saldo"]
        else:
            saldo_berjalan = saldo_berjalan + txn["Debit"] - txn["Kredit"]
            txn["Saldo"] = round(saldo_berjalan, 2)

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
# UI STREAMLIT
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

        st.subheader("📊 Preview Data (Saldo Presisi)")
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
