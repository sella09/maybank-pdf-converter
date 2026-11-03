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
def sanitize(text):
    if not text: return ""
    t = str(text).replace('\xa0', ' ').replace('"', '').replace("'", "")
    return " ".join(t.split()).strip()

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
    if num_str is None: return 0.0
    s = str(num_str).strip()
    if not s or s in ("-", ""): return 0.0
    s = re.sub(r"[^\d.,]", "", s)
    if not s: return 0.0
    if re.search(r",\d{2}$", s) and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try: return float(s)
    except ValueError: return 0.0

def clean_desc(text):
    t = sanitize(text)
    t = re.sub(r"Bersambung ke halaman berikut\.?", "", t, flags=re.I)
    t = re.sub(r"SALDO\s+AWAL\s*:.*?(SALDO\s+AKHIR\s*:.*)?", "", t, flags=re.I)
    t = re.sub(r"CATATAN\s*:.*?Rekening\.?", "", t, flags=re.I | re.S)
    t = re.sub(r"HALAMAN\s*:.*?IDR", "", t, flags=re.I)
    return t

# =========================================================
# PARSER BCA (BLOCK MULTI-LINE)
# =========================================================
@st.cache_data(show_spinner=False)
def parse_bca(file_bytes, password=None):
    file_obj = BytesIO(file_bytes)
    pdf = open_pdf(file_obj, password)
    transactions = []
    tahun = "2026"
    order = 0

    for page in pdf.pages:
        if page.page_number == 1:
            text = page.extract_text() or ""
            m = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", text, re.IGNORECASE)
            if m: tahun = m.group(1)

        # Ambil semua kata dengan posisi
        words = page.extract_words(keep_blank_chars=True)
        lines = {}
        for w in words:
            y = round(w['top'], 1)
            if y not in lines: lines[y] = []
            lines[y].append(w)

        # Gabungkan baris menjadi satu list teks utuh
        raw_lines = []
        for y in sorted(lines.keys()):
            row = sorted(lines[y], key=lambda w: w['x0'])
            raw_lines.append(sanitize(" ".join([w['text'] for w in row])))

        current_txn = None

        for row_text in raw_lines:
            if not row_text: continue
            if any(k in row_text.upper() for k in ["REKENING GIRO", "KCIU RAWAMANGUN", "NO. REKENING", "PERIODE :", "MATA UANG", "CATATAN:", "BCA BERHAK"]):
                continue

            m_date = re.match(r"^(\d{2}/\d{2})\s+(.+)", row_text)
            
            if m_date:
                # Simpan transaksi sebelumnya jika ada
                if current_txn:
                    current_txn["Keterangan"] = clean_desc(current_txn["Keterangan"])
                    transactions.append(current_txn)

                tgl = m_date.group(1)
                rest_text = m_date.group(2)

                current_txn = {
                    "Tanggal": f"{tgl}/{tahun}",
                    "Kode": "-",
                    "Keterangan": rest_text,
                    "Debit": 0.0,
                    "Kredit": 0.0,
                    "Saldo": None,
                    "_order": order,
                }
                order += 1

            else:
                # Jika bukan tanggal, ini baris lanjutan. Gabungkan ke transaksi aktif.
                if current_txn:
                    current_txn["Keterangan"] += " " + row_text

        if current_txn:
            current_txn["Keterangan"] = clean_desc(current_txn["Keterangan"])
            transactions.append(current_txn)
            current_txn = None

    pdf.close()

    # EKSTRAKSI ANGKA DARI BLOK KETERANGAN YANG SUDAH DIGABUNG
    for txn in transactions:
        full_text = txn["Keterangan"]
        upper_text = full_text.upper()

        # Cari semua angka (Rupiah)
        nums = re.findall(r"\b\d{1,3}(?:,\d{3})*\.\d{2}\b|\b\d+\.\d{2}\b", full_text)
        parsed_nums = [to_float(n) for n in nums]

        # Hapus angka dan kode referensi dari keterangan
        clean_text = full_text
        for n in nums:
            clean_text = clean_text.replace(n, "")
        clean_text = re.sub(r"\b\d{4}/[A-Z0-9/]+\b", "", clean_text) # hapus FTSCY dll
        clean_text = re.sub(r"\bDB\b|\bCR\b", "", clean_text, flags=re.I)
        txn["Keterangan"] = sanitize(clean_text)

        # Tentukan Nominal dan Saldo
        if len(parsed_nums) >= 2:
            nominal = parsed_nums[0]
            saldo = parsed_nums[-1]
        elif len(parsed_nums) == 1:
            nominal = parsed_nums[0]
            saldo = None
        else:
            nominal = 0.0
            saldo = None

        # ATURAN DEBIT / KREDIT BCA
        is_db = False
        # 1. Prioritas Utama (Uang Keluar)
        if any(k in upper_text for k in ["EDDY SUFENDI", "BIAYA ADM", "PAJAK BUNGA", "PENERIMAAN NEGARA", "BA JASA", "AUTODEBET", "ND-LAINNYA", "SWITCHING DB", "TRSF E-BANKING DB"]):
            is_db = True
        # 2. Prioritas Kedua (Uang Masuk)
        elif any(k in upper_text for k in ["BI-FAST", "TRSF E-BANKING CR", "SETORAN TUNAI", "BUNGA"]):
            is_db = False

        debit_val = nominal if is_db else 0.0
        kredit_val = nominal if not is_db else 0.0

        # Bunga masuk Kredit
        if "BUNGA" in upper_text and "PAJAK" not in upper_text:
            debit_val, kredit_val = 0.0, nominal

        txn["Debit"] = debit_val
        txn["Kredit"] = kredit_val
        txn["Saldo"] = saldo

    # Validasi Silang Saldo (Paksa pembalik)
    saldo_berjalan = 0.0
    for txn in transactions:
        # Jika saldo terdeteksi, jadikan patokan perhitungan selanjutnya
        if txn["Saldo"] is not None:
            if saldo_berjalan != 0.0:
                expected = round(saldo_berjalan - txn["Debit"] + txn["Kredit"], 2)
                # Jika selisih > 1.0, coba tukar Debit/Kredit
                if abs(expected - txn["Saldo"]) > 1.0:
                    expected_swap = round(saldo_berjalan + txn["Debit"] - txn["Kredit"], 2)
                    if abs(expected_swap - txn["Saldo"]) <= 1.0:
                        txn["Debit"], txn["Kredit"] = txn["Kredit"], txn["Debit"]
            saldo_berjalan = txn["Saldo"]
        else:
            # Jika saldo tidak ada, hitung manual
            saldo_berjalan = round(saldo_berjalan - txn["Debit"] + txn["Kredit"], 2)
            txn["Saldo"] = saldo_berjalan

        if txn["Debit"] > 0: txn["Tipe Asli"] = "DB"
        elif txn["Kredit"] > 0: txn["Tipe Asli"] = "CR"
        else: txn["Tipe Asli"] = "-"

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
        parsed = [to_float(n) for n in nums]
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
                "Saldo": t["Saldo"], "Tipe Asli Bank": t.get("Tipe Asli", "-"),
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
        final_df = final_df.sort_values(["Tanggal_DT", "_order"]).reset_index(drop=True)
        final_df = final_df.drop(columns=["Tanggal_DT", "_order"])

        st.subheader("📊 Preview Data (Bersih Total)")
        st.dataframe(final_df, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Mutasi")

        st.download_button(
            "📥 Download Excel",
            data=output.getvalue(),
            file_name="mutasi_bank_final_multiline.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
