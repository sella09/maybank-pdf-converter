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

# ==== BERSIHKAN TEKS FOOTER/HEADER BCA ====
FOOTER_PATTERNS = [
    r"Bersambung ke halaman berikut\.?",
    r"SALDO AWAL\s*:\s*[\d,.]+\s*MUTASI\s*(CR|DB)?\s*:\s*[\d,.]+\s*MUTASI\s*(CR|DB)?\s*:\s*[\d,.]+\s*SALDO AKHIR\s*:\s*[\d,.]+",
    r"MUTASI\s*(CR|DB)?\s*:\s*[\d,.]+\s*MUTASI\s*(CR|DB)?\s*:\s*[\d,.]+\s*SALDO AKHIR\s*:\s*[\d,.]+",
    r"CATATAN\s*:.*?Rekening\.?",
    r"Apabila nasabah.*?Rekening\.?",
    r"BCA berhak.*?Rekening\.?",
    r"HALAMAN\s*:\s*\d+\s*/\s*\d+",
    r"PERIODE\s*:\s*[A-Za-z]+\s+\d{4}",
    r"MATA UANG\s*:\s*IDR",
    r"REKENING GIRO",
    r"KCIU RAWAMANGUN",
]

def clean_desc(text):
    if not text: return ""
    t = str(text)
    for pat in FOOTER_PATTERNS:
        t = re.sub(pat, "", t, flags=re.I | re.S)
    return " ".join(t.split()).strip()

# =========================================================
# PARSER BCA — pdfplumber, keterangan UTUH, DB/CR DIPERBAIKI
# =========================================================
DB_KEYWORDS = [
    "EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET",
    "PENERIMAAN NEGARA", "TIGA MUSTIKA KONSU", "PT.FINNET", "PT. FINNET",
    "BA JASA E-BANKING", "BIAYA ADM", "BIAYA TRANSFER", "BIAYA REFERENSI",
    "SWITCHING KE", "SWITCHING TRANSFER KE", "SWITCHING DB",
    "BI-FAST BIF TRANSFER KE", "BI-FAST BIF BIAYA TXN KE",
    "OTOMATIS", "FEE TELEX", "FEE FULL AMT", "FEE VALUE TDY",
    "PAJAK BUNGA", "PEMBELIAN", "BY PASANG", "BAYAR",
    "TRSF E-BANKING DB", "KR OTOMATIS", "ND-BIAYA ADM",
    "ND-LAINNYA", "BYR VIA E-BANKING", "SETORAN TUNAI BAYAR TAGIHAN",
]
CR_KEYWORDS = [
    "TRSF E-BANKING CR", "BI-FAST CR", "BI-FAST CRBIF", "KR OTOMATIS LLG",
    "SETORAN TUNAI", "BUNGA",
]

def detect_dbcr(upper_text):
    if re.search(r"\bDB\b", upper_text):
        return True
    if re.search(r"\bCR\b", upper_text):
        return False
    if any(k in upper_text for k in DB_KEYWORDS):
        return True
    if any(k in upper_text for k in CR_KEYWORDS):
        return False
    return False

@st.cache_data(show_spinner=False)
def parse_bca(file_bytes, password=None):
    transactions = []
    file_obj = BytesIO(file_bytes)
    pdf = open_pdf(file_obj, password)

    tahun = "2026"
    saldo_awal = 0.0
    order = 0

    for page in pdf.pages:
        if page.page_number == 1:
            text = page.extract_text() or ""
            m = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", text, re.IGNORECASE)
            if m: tahun = m.group(1)
            m_saldo = re.search(r"SALDO\s+AWAL\s*[\n:]?\s*([\d,.]+)", text, re.IGNORECASE)
            if m_saldo: saldo_awal = to_float(m_saldo.group(1)) or 0.0

        words = page.extract_words(keep_blank_chars=True)
        lines = {}
        for w in words:
            y = round(w['top'], 1)
            if y not in lines: lines[y] = []
            lines[y].append(w)

        prev_txn = None

        for y in sorted(lines.keys()):
            row = sorted(lines[y], key=lambda w: w['x0'])
            row_text = " ".join([w['text'] for w in row])

            # Cek dulu apakah ini baris footer (tanpa tanggal)
            m_date = re.match(r"^(\d{2}/\d{2})\s+(.+)", row_text)
            if not m_date:
                if prev_txn is not None:
                    extra = re.sub(r"^\d[\d,]*\.\d{2}\b", "", row_text).strip()
                    extra = re.sub(r"[\d,]+\.\d{2}", "", extra).strip()
                    extra = re.sub(r"\bDB\b|\bCR\b", "", extra, flags=re.I).strip()
                    extra = clean_desc(extra)
                    if extra:
                        prev_txn["Keterangan"] = (prev_txn["Keterangan"] + " " + extra).strip()
                continue

            tgl = m_date.group(1)
            rest_text = m_date.group(2)

            nums = []
            for w in row:
                if re.match(r"^\d[\d,]*\.\d{2}$", w['text']):
                    nums.append({"val": to_float(w['text']), "x": w['x0']})

            if len(nums) < 1: continue

            nums = sorted(nums, key=lambda n: n['x'])
            nominal = nums[0]['val']
            saldo = nums[-1]['val'] if len(nums) >= 2 else None
            if saldo == nominal:
                saldo = None

            upper_rest = rest_text.upper()

            # === KHUSUS BARIS SALDO AWAL ===
            if upper_rest.startswith("SALDO AWAL") or "SALDO AWAL" in upper_rest[:20]:
                desc = "SALDO AWAL"
                # Baris SALDO AWAL bukan transaksi: Debit 0, Kredit 0, Saldo = nominal
                saldo_awal = nominal
                transactions.append({
                    "Tanggal": f"{tgl}/{tahun}",
                    "Kode": "-",
                    "Keterangan": desc,
                    "Debit": 0.0,
                    "Kredit": 0.0,
                    "Saldo": nominal,
                    "Tipe Asli": "AWAL",
                    "_order": order,
                })
                order += 1
                prev_txn = transactions[-1]
                continue

            is_db = detect_dbcr(upper_rest)

            desc = rest_text
            desc = re.sub(r"\b\d[\d,]*\.\d{2}\b", "", desc)
            desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
            desc = clean_desc(desc)

            debit_val, kredit_val = (nominal, 0.0) if is_db else (0.0, nominal)

            if "BUNGA" in upper_rest and "PAJAK" not in upper_rest:
                debit_val, kredit_val = 0.0, nominal
                is_db = False

            txn = {
                "Tanggal": f"{tgl}/{tahun}",
                "Kode": "-",
                "Keterangan": desc,
                "Debit": debit_val,
                "Kredit": kredit_val,
                "Saldo": saldo,
                "Tipe Asli": "DB" if is_db else "CR",
                "_order": order,
            }
            transactions.append(txn)
            order += 1
            prev_txn = txn

    pdf.close()

    # Sort stabil: Tanggal lalu urutan asli PDF
    transactions.sort(key=lambda x: (x["Tanggal"], x["_order"]))

    # Running balance
    saldo_berjalan = saldo_awal
    for txn in transactions:
        if txn["Tipe Asli"] == "AWAL":
            saldo_berjalan = txn["Saldo"]
            continue
        if txn["Saldo"] is not None:
            saldo_berjalan = txn["Saldo"]
        else:
            saldo_berjalan = saldo_berjalan - txn["Debit"] + txn["Kredit"]
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

    order = 0
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

        debit_val, kredit_val = (nominal, 0.0) if is_db else (0.0, nominal)
        transactions.append({
            "Tanggal": tgl, "Kode": "-", "Keterangan": desc,
            "Debit": debit_val, "Kredit": kredit_val,
            "Saldo": saldo, "Tipe Asli": "DB" if is_db else "CR",
            "_order": order,
        })
        order += 1
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
                "Saldo": t["Saldo"], "Tipe Asli Bank": t["Tipe Asli"],
                "_order": t.get("_order", 0),
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
        # Sort stabil: Tanggal → urutan asli PDF
        final_df = final_df.sort_values(["Tanggal_DT", "_order"]).reset_index(drop=True)
        final_df = final_df.drop(columns=["Tanggal_DT", "_order"])

        # Tampilkan semua, termasuk SALDO AWAL
        st.subheader("📊 Preview Data (Keterangan Bersih)")
        st.dataframe(final_df, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Mutasi")

        st.download_button(
            "📥 Download Excel",
            data=output.getvalue(),
            file_name="mutasi_bank_clean.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
