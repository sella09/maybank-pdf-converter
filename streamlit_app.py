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
# PARSER BCA (HYBRID: REGEX + AMBIL SALDO ASLI)
# =========================================================
@st.cache_data(show_spinner=False)
def parse_bca(file_bytes, password=None):
    transactions = []
    file_obj = BytesIO(file_bytes)
    pdf = open_pdf(file_obj, password)

    tahun = "2026"

    for page in pdf.pages:
        # Dapatkan tahun dari teks halaman pertama
        if page.page_number == 1:
            text = page.extract_text() or ""
            m = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", text, re.IGNORECASE)
            if m: tahun = m.group(1)

        # Ekstrak teks halaman (karena tabel tidak terbaca, kita olah teksnya)
        text = page.extract_text()
        if not text: continue
        
        lines = text.split("\n")
        
        # Gabungkan baris yang terpotong (tidak diawali tanggal)
        merged_lines = []
        for line in lines:
            s = line.strip()
            if not s: continue
            # Jika diawali tanggal DD/MM, mulai baris baru. Jika tidak, gabung ke baris sebelumnya.
            if re.match(r"^\d{2}/\d{2}", s):
                merged_lines.append(s)
            elif merged_lines and not any(k in s for k in ["REKENING GIRO", "KCU", "PERIODE", "CATATAN", "HALAMAN"]):
                merged_lines[-1] += " " + s

        for line in merged_lines:
            m_date = re.match(r"^(\d{2}/\d{2})(.*)", line)
            if not m_date: continue
            
            tgl = m_date.group(1)
            rest = m_date.group(2).strip()

            # Lewati saldo awal / akhir yang bukan transaksi
            if "SALDO AWAL" in rest.upper() or "SALDO AKHIR" in rest.upper(): continue

            # Ekstrak semua format angka desimal (titik 2 digit di belakang)
            nums = re.findall(r"\d[\d,]*\.\d{2}", rest)
            if not nums: continue

            # Ambil nominal mutasi (angka pertama) dan saldo (angka kedua)
            nominal = to_float(nums[0])
            saldo = to_float(nums[1]) if len(nums) > 1 else None
            
            # Deteksi DB / CR
            upper_rest = rest.upper()
            is_db = False
            if "DB" in upper_rest:
                is_db = True
            else:
                db_keywords = ["EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET", "PENERIMAAN NEGARA",
                               "TIGA MUSTIKA KONSU", "PT.FINNET", "PT. FINNET", "BA JASA E-BANKING",
                               "BIAYA ADM", "BIAYA TRANSFER", "BIAYA REFERENSI", "SWITCHING KE",
                               "SWITCHING TRANSFER KE", "BI-FAST BIF TRANSFER KE",
                               "BI-FAST BIF BIAYA TXN KE", "OTOMATIS", "FEE TELEX", "FEE FULL AMT",
                               "FEE VALUE TDY", "PAJAK BUNGA"]
                if any(k in upper_rest for k in db_keywords):
                    is_db = True

            # Bersihkan deskripsi dari angka dan kata DB/CR
            desc = rest
            for n in nums:
                desc = desc.replace(n, "")
            desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
            desc = clean_desc(desc)

            # Pemetaan Anda: DB -> Kredit, CR -> Debit
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

    pdf.close()
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

        st.subheader("📊 Preview Data (Saldo Asli dari Bank)")
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
