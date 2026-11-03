import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

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
                raise ValueError("PDF ini terproteksi password. Isi kolom 🔑 Password PDF.")
            raise ValueError("Password PDF salah.")
        raise


def to_float(num_str):
    """Konversi '1,234.56' / '1.234,56' / '1241.00' ke float."""
    if num_str is None:
        return None
    s = str(num_str).strip()
    if not s or s in ("-", ""):
        return None
    # buang simbol mata uang & spasi
    s = re.sub(r"[^\d.,]", "", s)
    if not s:
        return None
    # format Indonesia: 1.234,56
    if re.search(r",\d{2}$", s) and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def clean_desc(text):
    """Bersihkan keterangan dari footer/noise BCA."""
    if not text:
        return ""
    t = str(text)
    # buang kalimat footer BCA
    t = re.sub(
        r"Rekening ini sampai dengan akhir bulan berikutnya.*?Rekening ini\.",
        "", t, flags=re.IGNORECASE | re.DOTALL,
    )
    t = re.sub(r"telah menyetujui segala data.*?Rekening ini\.", "", t, flags=re.I | re.S)
    # rapikan spasi
    t = " ".join(t.split())
    return t.strip()


# =========================================================
# PARSER BCA (berbasis TABEL)
# =========================================================
def parse_bca_perfect(file_obj, password=None):
    transactions = []
    pdf = open_pdf(file_obj, password)
    try:
        # deteksi tahun dari teks
        full_text = ""
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

        tahun = "2026"
        m = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", full_text, re.IGNORECASE)
        if m:
            tahun = m.group(1)

        # parse tabel per halaman
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row:
                        continue
                    # normalisasi cell
                    cells = [clean_desc(c) if c else "" for c in row]
                    # gabung semua cell untuk cek tanggal
                    joined = " | ".join(cells)

                    # cari tanggal DD/MM di cell pertama (atau mana saja)
                    tgl = None
                    for c in cells:
                        mt = re.match(r"^(\d{2}/\d{2})\b", c)
                        if mt:
                            tgl = mt.group(1)
                            break
                    if not tgl:
                        continue

                    # gabungkan seluruh teks baris (tanpa tanggal) untuk analisis
                    row_text = " ".join(cells)
                    # buang tanggal di awal
                    row_text = re.sub(r"^\d{2}/\d{2}\s*", "", row_text)

                    upper = row_text.upper()

                    # skip baris saldo awal / header
                    if "SALDO AWAL" in upper:
                        continue
                    if "TANGGAL" in upper and "KETERANGAN" in upper:
                        continue

                    # --- deteksi arah mutasi (DB/CR) ---
                    # Prioritas: cari token DB/CR yang berdiri sendiri
                    is_db = bool(re.search(r"\bDB\b", upper))
                    is_cr = bool(re.search(r"\bCR\b", upper))

                    # kasus khusus
                    if "PAJAK BUNGA" in upper or "BIAYA ADM" in upper:
                        is_db = True
                        is_cr = False
                    elif "BUNGA" in upper and "PAJAK" not in upper:
                        is_db = False
                        is_cr = True

                    # jika tidak ada penanda, default CR (pemasukan) — tapi
                    # cek kata kunci DB yang umum
                    if not is_db and not is_cr:
                        # cek keyword yang biasanya DB
                        db_keywords = ["EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET",
                                       "PENERIMAAN NEGARA", "BIAYA TRANSFER SME",
                                       "BIAYA REFERENSI", "FEE TELEX", "FEE FULL",
                                       "FEE VALUE", "TRANSFER OR VIA IB",
                                       "SWITCHING DB", "BIF TRANSFER KE",
                                       "PAJAK BUNGA", "BIAYA ADM"]
                        if any(k in upper for k in db_keywords):
                            is_db = True

                    # --- ekstrak semua angka desimal ---
                    # pola: 1,234,567.89 atau 1234567.89
                    decimals = re.findall(r"\d[\d,]*\.\d{2}", row_text)
                    if not decimals:
                        continue

                    # nominal = angka pertama, saldo = angka terakhir (jika berbeda)
                    nominal = to_float(decimals[0]) or 0.0
                    saldo = None
                    if len(decimals) >= 2:
                        saldo_val = to_float(decimals[-1])
                        # saldo valid jika berbeda dari nominal dan > 0
                        if saldo_val is not None and saldo_val != nominal and saldo_val > 0:
                            saldo = saldo_val

                    # --- bersihkan deskripsi ---
                    desc = row_text
                    for d in decimals:
                        desc = desc.replace(d, "")
                    # buang token DB/CR
                    desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
                    # buang kode referensi panjang (CBG, dll) — opsional
                    desc = " ".join(desc.split()).strip()
                    if not desc:
                        desc = "(tanpa keterangan)"

                    # --- SWAP: DB asli -> Kredit, CR asli -> Debit ---
                    if is_db:
                        debit_val = 0.0
                        kredit_val = nominal
                        tipe = "DB"
                    else:
                        debit_val = nominal
                        kredit_val = 0.0
                        tipe = "CR"

                    transactions.append({
                        "Tanggal": f"{tgl}/{tahun}",
                        "Kode": "-",
                        "Keterangan": desc,
                        "Debit": debit_val,
                        "Kredit": kredit_val,
                        "Saldo": saldo,
                        "Tipe Asli": tipe,
                    })
    finally:
        pdf.close()

    return transactions


# =========================================================
# PARSER MAYBANK (fallback sederhana)
# =========================================================
def parse_maybank_statement(file_obj, password=None):
    transactions = []
    pdf = open_pdf(file_obj, password)
    try:
        full_text = ""
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"
    finally:
        pdf.close()

    lines = full_text.split("\n")
    date_re = re.compile(r"^(\d{2}[/-]\d{2}[/-]\d{4}|\d{2}[/-]\d{2})\b")

    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue
        m = date_re.match(line_str)
        if not m:
            continue
        tgl = m.group(1)
        nums = re.findall(r"\d[\d,]*\.\d{2}", line_str)
        parsed = [to_float(n) for n in nums if to_float(n) is not None]
        if not parsed:
            continue
        nominal = parsed[0]
        saldo = parsed[-1] if len(parsed) >= 2 and parsed[-1] != nominal else None
        is_db = bool(re.search(r"\bDB\b", line_str.upper()))
        desc = re.sub(r"^\d{2}[/-]\d{2}[/-]?\d{0,4}\s*", "", line_str)
        for n in nums:
            desc = desc.replace(n, "")
        desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
        desc = " ".join(desc.split()).strip()

        if is_db:
            debit_val, kredit_val = 0.0, nominal
        else:
            debit_val, kredit_val = nominal, 0.0

        transactions.append({
            "Tanggal": tgl, "Kode": "-", "Keterangan": desc,
            "Debit": debit_val, "Kredit": kredit_val,
            "Saldo": saldo, "Tipe Asli": "DB" if is_db else "CR",
        })
    return transactions


# =========================================================
# BUILD DF
# =========================================================
def build_df(transactions, sumber_file):
    rows = []
    for txn in transactions:
        rows.append({
            "Sumber File": sumber_file,
            "Tanggal": txn["Tanggal"],
            "Kode": txn["Kode"],
            "Keterangan": txn["Keterangan"],
            "Debit": txn["Debit"],
            "Kredit": txn["Kredit"],
            "Saldo": txn["Saldo"],
            "Tipe Asli Bank": txn["Tipe Asli"],
        })
    return pd.DataFrame(rows)


# =========================================================
# UI
# =========================================================
col1, col2, col3 = st.columns(3)
with col1:
    bank_choice = st.selectbox("Pilih Bank", ["BCA", "Maybank"])
with col2:
    password = st.text_input("🔑 Password PDF (jika ada)", type="password")
with col3:
    uploaded_files = st.file_uploader("📄 Pilih file PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        try:
            st.info(f"Memproses **{uploaded_file.name}**...")
            if bank_choice == "BCA":
                txns = parse_bca_perfect(uploaded_file, password=password)
            else:
                txns = parse_maybank_statement(uploaded_file, password=password)

            df = build_df(txns, uploaded_file.name)
            if not df.empty:
                all_dfs.append(df)
                st.success(f"✅ {uploaded_file.name}: {len(df)} transaksi.")
            else:
                st.warning(f"⚠️ {uploaded_file.name}: Tidak ada transaksi.")
        except ValueError as ve:
            st.error(f"🔒 {uploaded_file.name}: {ve}")
        except Exception as e:
            st.error(f"❌ {uploaded_file.name}: {type(e).__name__}")
            with st.expander("Detail error"):
                st.exception(e)

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        final_df["Tanggal"] = final_df["Tanggal"].astype(str).str.strip()
        final_df["Tanggal_DT"] = pd.to_datetime(final_df["Tanggal"], format="%d/%m/%Y", errors="coerce")
        final_df = final_df.sort_values("Tanggal_DT").reset_index(drop=True)
        final_df = final_df.drop(columns=["Tanggal_DT"])

        st.subheader("📊 Preview Data (Debet ↔ Kredit swapped)")
        st.dataframe(final_df, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Mutasi_Swapped")
        st.download_button(
            "📥 Download Excel",
            data=output.getvalue(),
            file_name="mutasi_bank_correct_saldo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
