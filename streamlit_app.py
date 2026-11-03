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
    # Hapus karakter selain angka, koma, dan titik
    s = re.sub(r"[^\d.,]", "", s)
    if not s: return None
    # Format Indonesia: titik ribuan, koma desimal
    if re.search(r",\d{2}$", s) and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try: return float(s)
    except ValueError: return None

def clean_desc(text):
    if not text: return ""
    t = str(text).replace("\n", " ")
    # Bersihkan teks sampah
    t = re.sub(r"Rekening ini sampai dengan akhir bulan berikutnya.*?Rekening ini\.", "", t, flags=re.I | re.S)
    return " ".join(t.split()).strip()

# =========================================================
# PARSER BCA (TABEL EKSTRAKSI)
# =========================================================
@st.cache_data(show_spinner=False)
def parse_bca(file_bytes, password=None):
    transactions = []
    file_obj = BytesIO(file_bytes)
    pdf = open_pdf(file_obj, password)

    tahun = "2026"
    saldo_awal = 0.0

    for page in pdf.pages:
        # Dapatkan tahun dan saldo awal dari teks halaman pertama
        if page.page_number == 1:
            text = page.extract_text() or ""
            m = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", text, re.IGNORECASE)
            if m: tahun = m.group(1)
            m_saldo = re.search(r"SALDO\s+AWAL\s*[\n:]?\s*([\d,.]+)", text, re.IGNORECASE)
            if m_saldo: saldo_awal = to_float(m_saldo.group(1)) or 0.0

        # Ekstrak tabel secara struktural
        tables = page.extract_tables()
        if not tables: continue

        for table in tables:
            for row in table:
                # Lewati baris yang kosong atau bukan transaksi
                if not row or len(row) < 4: continue
                
                # Cek apakah kolom pertama adalah tanggal (format DD/MM)
                tgl_raw = str(row[0]).strip() if row[0] else ""
                if not re.match(r"^\d{2}/\d{2}$", tgl_raw): continue
                
                # Bersihkan setiap sel dari nilai None
                row = [str(c).replace("\n", " ").strip() if c else "" for c in row]
                tgl = row[0]
                desc = row[1]
                
                # Di PDF BCA, mutasi (nominal) biasanya di kolom 2, saldo di kolom 3
                mutasi_str = row[2] if len(row) > 2 else ""
                saldo_str = row[3] if len(row) > 3 else ""

                # Parsing nilai
                saldo_val = to_float(saldo_str)
                mutasi_val = to_float(mutasi_str)

                # Jika saldo kosong, skip
                if saldo_val is None: continue
                if mutasi_val is None: mutasi_val = 0.0

                # Tentukan arah (DB / CR) berdasarkan teks deskripsi atau mutasi
                upper_desc = desc.upper()
                is_db = "DB" in upper_desc or "DB" in mutasi_str.upper()
                # Fallback deteksi DB berdasarkan keyword
                if not is_db:
                    db_keywords = ["EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET", "PENERIMAAN NEGARA",
                                   "TIGA MUSTIKA KONSU", "PT.FINNET", "PT. FINNET", "BA JASA E-BANKING",
                                   "BIAYA ADM", "BIAYA TRANSFER", "BIAYA REFERENSI", "SWITCHING KE",
                                   "SWITCHING TRANSFER KE", "BI-FAST BIF TRANSFER KE",
                                   "BI-FAST BIF BIAYA TXN KE", "OTOMATIS", "FEE TELEX", "FEE FULL AMT",
                                   "FEE VALUE TDY"]
                    if any(k in upper_desc for k in db_keywords):
                        is_db = True
                    elif re.search(r"\bCR\b", upper_desc):
                        is_db = False

                # Bersihkan keterangan dari kata DB / CR
                clean_keterangan = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
                clean_keterangan = clean_desc(clean_keterangan)

                # Pemetaan: DB = uang keluar (Kredit), CR = uang masuk (Debit)
                debit_val, kredit_val = (0.0, mutasi_val) if is_db else (mutasi_val, 0.0)

                transactions.append({
                    "Tanggal": f"{tgl}/{tahun}",
                    "Kode": "-",
                    "Keterangan": clean_keterangan,
                    "Debit": debit_val,
                    "Kredit": kredit_val,
                    "Saldo": saldo_val,
                    "Tipe Asli": "DB" if is_db else "CR"
                })

    pdf.close()
    return transactions

# =========================================================
# PARSER MAYBANK (TETAP)
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

        st.subheader("📊 Preview Data (Saldo Mengikuti PDF)")
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
