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
    t = re.sub(r"telah menyetujui segala data.*?Rekening ini\.", "", t, flags=re.I | re.S)
    t = re.sub(r"Apabila nasabah tidak melakukan sanggahan.*?Laporan Mutasi Rekening\.", "", t, flags=re.I | re.S)
    return " ".join(t.split()).strip()

# =========================================================
# PARSER BCA (SALDO MENGIKUTI PDF)
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

    # 1. SALDO AWAL
    saldo_awal = 0.0
    m_saldo = re.search(r"SALDO\s+AWAL\s*[\n:]?\s*([\d,.]+)", full_text, re.IGNORECASE)
    if m_saldo:
        saldo_awal = to_float(m_saldo.group(1)) or 0.0

    # 2. PAJAK BUNGA & BUNGA
    for match in re.finditer(r"PAJAK\s+BUNGA\s+([\d,.]+)", full_text, re.IGNORECASE):
        pajak_val = to_float(match.group(1))
        if pajak_val:
            ctx = full_text[max(0, match.start() - 500):match.start()]
            dates = re.findall(r"\b(\d{2}/\d{2})\b", ctx)
            tgl = dates[-1] if dates else None
            if tgl:
                transactions.append({
                    "Tanggal": f"{tgl}/{tahun}", "Kode": "-",
                    "Keterangan": "PAJAK BUNGA", "Debit": 0.0, "Kredit": pajak_val,
                    "Saldo": None, "Tipe Asli": "DB",
                })
    for match in re.finditer(r"(?<!PAJAK\s)BUNGA\s+([\d,.]+)", full_text, re.IGNORECASE):
        if "PAJAK" in full_text[max(0, match.start()-20):match.start()].upper():
            continue
        bunga_val = to_float(match.group(1))
        if bunga_val:
            ctx = full_text[max(0, match.start() - 500):match.start()]
            dates = re.findall(r"\b(\d{2}/\d{2})\b", ctx)
            tgl = dates[-1] if dates else None
            if tgl:
                transactions.append({
                    "Tanggal": f"{tgl}/{tahun}", "Kode": "-",
                    "Keterangan": "BUNGA", "Debit": bunga_val, "Kredit": 0.0,
                    "Saldo": None, "Tipe Asli": "CR",
                })

    # 3. PARSE BARIS PER BARIS
    lines = full_text.split('\n')
    ignore_keywords = [
        "REKENING GIRO", "KCU", "DOMINIC DIGITAL", "KELAPA GADING", "JL BOULEVARD",
        "JAKARTA", "INDONESIA", "NO. REKENING", "HALAMAN", "PERIODE", "MATA UANG",
        "CATATAN:", "Apabila nasabah", "BCA berhak", "TANGGAL KETERANGAN",
        "Bersambung ke halaman", "MUTASI CR :", "MUTASI DB :", "SALDO AKHIR :",
        "SALDO AWAL", "===== Page", "# REKENING", "Page ", "- Page ",
    ]

    clean_lines = []
    for line in lines:
        s = line.strip()
        if not s: continue
        if re.match(r"^\d+\s*/\s*\d+$", s): continue
        if re.match(r"^\d+$", s) and len(s) <= 3: continue
        if any(kw.lower() in s.lower() for kw in ignore_keywords): continue
        if re.search(r"BUNGA", s, re.IGNORECASE): continue
        clean_lines.append(s)

    raw_blocks = []
    current_block = None
    date_pattern = re.compile(r"^(\d{2}/\d{2})\s*(.+)")

    for s in clean_lines:
        m_date = date_pattern.match(s)
        if m_date:
            if current_block: raw_blocks.append(current_block)
            current_block = {"tgl": m_date.group(1), "lines": [m_date.group(2)]}
        elif current_block:
            current_block["lines"].append(s)
    if current_block: raw_blocks.append(current_block)

    db_keywords = [
        "EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET", "PENERIMAAN NEGARA",
        "TIGA MUSTIKA KONSU", "PT.FINNET", "PT. FINNET", "BA JASA E-BANKING",
        "BIAYA TRANSFER", "BIAYA ADM", "BIAYA REFERENSI", "SWITCHING KE",
        "SWITCHING TRANSFER KE", "BI-FAST BIF TRANSFER KE",
        "BI-FAST BIF BIAYA TXN KE", "OTOMATIS", "FEE TELEX", "FEE FULL AMT",
        "FEE VALUE TDY",
    ]

    for block in raw_blocks:
        combined = " ".join(block["lines"])
        upper = combined.upper()
        nums = re.findall(r"\d[\d,]*\.\d{2}", combined)
        if not nums: continue
        num_vals = [to_float(n) for n in nums if to_float(n) is not None]
        if not num_vals: continue

        # BIAYA ADM / TRANSFER / REFERENSI
        if any(k in upper for k in ["BIAYA ADM", "BIAYA TRANSFER", "BIAYA REFERENSI"]):
            biaya_val = num_vals[0]
            saldo_val = None
            if len(num_vals) >= 2 and num_vals[-1] != biaya_val and num_vals[-1] > 1000:
                saldo_val = num_vals[-1]
            transactions.append({
                "Tanggal": f"{block['tgl']}/{tahun}", "Kode": "-",
                "Keterangan": "BIAYA ADM", "Debit": 0.0, "Kredit": biaya_val,
                "Saldo": saldo_val, "Tipe Asli": "DB",
            })
            continue

        # TENTUKAN DB / CR
        is_db = False
        if re.search(r"\bDB\b", upper): is_db = True
        elif any(k in upper for k in db_keywords): is_db = True
        elif re.search(r"\bCR\b", upper): is_db = False

        nominal = num_vals[0]

        # Ambil SALDO dari PDF
        saldo_pdf = None
        if len(num_vals) >= 2:
            cand = num_vals[-1]
            if cand != nominal and cand > 1000:
                saldo_pdf = cand

        desc = combined
        for n in nums: desc = desc.replace(n, "")
        desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
        desc = clean_desc(desc)

        # Pemetaan: DB -> Kredit, CR -> Debit
        debit_val, kredit_val = (0.0, nominal) if is_db else (nominal, 0.0)

        transactions.append({
            "Tanggal": f"{block['tgl']}/{tahun}", "Kode": "-", "Keterangan": desc,
            "Debit": debit_val, "Kredit": kredit_val,
            "Saldo": saldo_pdf, "Tipe Asli": "DB" if is_db else "CR",
        })

    # 4. RUNNING BALANCE (Fallback)
    for i, txn in enumerate(transactions):
        txn["_order"] = i
    transactions.sort(key=lambda x: (x["Tanggal"], x["_order"]))

    saldo_berjalan = saldo_awal
    for txn in transactions:
        if txn["Saldo"] is None:
            saldo_berjalan = saldo_berjalan + txn["Debit"] - txn["Kredit"]
            txn["Saldo"] = saldo_berjalan
        else:
            saldo_berjalan = txn["Saldo"]

    for txn in transactions:
        txn.pop("_order", None)

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

        # Pemetaan konsisten: DB -> Kredit, CR -> Debit
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

        st.subheader("📊 Preview Data (Saldo Running Balance)")
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
