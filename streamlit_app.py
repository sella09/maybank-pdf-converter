import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel Converter", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA & Maybank)")

# =========================================================
# HELPER: buka PDF dengan fallback password
# =========================================================
def open_pdf(file_obj, password=None):
    """
    Coba buka PDF. Jika butuh password dan password kosong/salah,
    lempar exception dengan pesan yang jelas.
    """
    file_obj.seek(0)
    try:
        if password:
            return pdfplumber.open(file_obj, password=password)
        return pdfplumber.open(file_obj)
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg or "decrypt" in msg:
            if not password:
                raise ValueError(
                    "PDF ini terproteksi password. Silakan isi kolom "
                    "🔑 Password PDF lalu proses ulang."
                )
            raise ValueError("Password PDF salah. Periksa kembali.")
        raise


def to_float(num_str):
    """Konversi '1,234.56' atau '1.234,56' ke float."""
    if num_str is None:
        return None
    s = str(num_str).strip()
    if not s:
        return None
    # Jika ada koma sebagai desimal (format ID) dan titik sebagai ribuan
    if re.search(r",\d{2}$", s) and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


# =========================================================
# PARSER BCA
# =========================================================
def parse_bca_perfect(file_obj, password=None):
    transactions = []
    pdf = open_pdf(file_obj, password)
    try:
        full_pages_text = []
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_pages_text.append(t)
    finally:
        pdf.close()

    full_text = "\n".join(full_pages_text)

    # Deteksi Tahun Periode
    tahun = "2026"
    match_thn = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", full_text, re.IGNORECASE)
    if match_thn:
        tahun = match_thn.group(1)

    lines = full_text.split("\n")

    ignore_keywords = [
        "REKENING GIRO", "KCU", "DOMINIC DIGITAL", "KELAPA GADING", "JL BOULEVARD",
        "JAKARTA", "INDONESIA", "NO. REKENING", "HALAMAN", "PERIODE", "MATA UANG",
        "CATATAN:", "Apabila nasabah", "BCA berhak", "TANGGAL KETERANGAN",
        "Bersambung ke halaman", "MUTASI CR :", "MUTASI DB :", "SALDO AKHIR :",
    ]

    clean_lines = []
    for line in lines:
        line_str = line.strip()
        if not line_str:
            continue
        # buang nomor halaman
        if re.match(r"^\d+\s*/\s*\d+$", line_str) or (
            re.match(r"^\d+$", line_str) and len(line_str) <= 3
        ):
            continue
        if any(kw.lower() in line_str.lower() for kw in ignore_keywords):
            continue
        clean_lines.append(line_str)

    raw_blocks = []
    current_block = None
    date_pattern = re.compile(r"^(\d{2}/\d{2})\s+(.+)")

    for line_str in clean_lines:
        m_date = date_pattern.match(line_str)
        if m_date:
            if current_block:
                raw_blocks.append(current_block)
            current_block = {"tgl": m_date.group(1), "lines": [m_date.group(2)]}
        elif current_block:
            upper_line = line_str.upper()
            # Bunga/Pajak bunga dianggap transaksi baru walau tanpa tanggal
            if "BUNGA" in upper_line or "PAJAK BUNGA" in upper_line:
                raw_blocks.append(current_block)
                current_block = {"tgl": current_block["tgl"], "lines": [line_str]}
            else:
                current_block["lines"].append(line_str)

    if current_block:
        raw_blocks.append(current_block)

    for block in raw_blocks:
        combined_text = " ".join(block["lines"])

        if "SALDO AWAL" in combined_text.upper():
            continue

        upper_text = combined_text.upper()

        # Deteksi arah mutasi asli
        # DB = debet (uang keluar), CR = kredit (uang masuk)
        has_db = bool(re.search(r"\bDB\b", upper_text))
        has_cr = bool(re.search(r"\bCR\b", upper_text))

        # Ambil semua angka berformat desimal
        all_decimals = re.findall(r"[\d,]+\.\d{2}", combined_text)
        if not all_decimals:
            continue

        nominal = 0.0
        saldo = None

        # --- Kasus khusus: PAJAK BUNGA / BIAYA ADM ---
        if "PAJAK BUNGA" in upper_text or "BIAYA ADM" in upper_text:
            is_db = True  # selalu pengeluaran
            db_match = re.search(r"([\d,]+\.\d{2})\s+DB", combined_text, re.IGNORECASE)
            if db_match:
                nominal = to_float(db_match.group(1)) or 0.0
            else:
                nominal = to_float(all_decimals[0]) or 0.0
            if len(all_decimals) >= 2:
                saldo_val = to_float(all_decimals[-1])
                if saldo_val is not None and saldo_val != nominal:
                    saldo = saldo_val

        # --- Kasus khusus: BUNGA ---
        elif "BUNGA" in upper_text and "PAJAK" not in upper_text:
            is_db = False  # selalu pemasukan
            nominal = to_float(all_decimals[0]) or 0.0
            if len(all_decimals) >= 2:
                saldo_val = to_float(all_decimals[-1])
                if saldo_val is not None and saldo_val != nominal:
                    saldo = saldo_val

        # --- Transaksi normal ---
        else:
            # nominal biasanya angka pertama, saldo angka terakhir
            if len(all_decimals) >= 2:
                nominal = to_float(all_decimals[0]) or 0.0
                saldo_val = to_float(all_decimals[-1])
                if saldo_val is not None and saldo_val != nominal:
                    saldo = saldo_val
            elif len(all_decimals) == 1:
                nominal = to_float(all_decimals[0]) or 0.0
                saldo = None

            # DB menang jika ada DB; jika tidak, CR -> kredit
            if has_db:
                is_db = True
            elif has_cr:
                is_db = False
            else:
                # fallback: tanpa penanda, anggap kredit
                is_db = False

        # Bersihkan deskripsi
        desc = combined_text
        for num_str in all_decimals:
            desc = desc.replace(num_str, "")
        desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.IGNORECASE)
        desc = " ".join(desc.split()).strip()

        # ==================================================
        # SWAP: DB asli -> masuk kolom KREDIT
        #       CR asli -> masuk kolom DEBIT
        # ==================================================
        if is_db:
            debit_val = 0.0
            kredit_val = nominal
        else:
            debit_val = nominal
            kredit_val = 0.0

        transactions.append({
            "Tanggal": f"{block['tgl']}/{tahun}",
            "Kode": "-",
            "Keterangan": desc,
            "Debit": debit_val,
            "Kredit": kredit_val,
            "Saldo": saldo,
            "Tipe Asli": "DB" if is_db else "CR",
        })

    return transactions


# =========================================================
# PARSER MAYBANK
# =========================================================
def parse_maybank_statement(file_obj, password=None):
    """
    Parser Maybank yang mem-parse baris transaksi lengkap:
    Tanggal | Keterangan | Debet | Kredit | Saldo
    lalu melakukan swap Debet <-> Kredit.
    """
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

    ignore_keywords = [
        "MUTASI", "SALDO AWAL", "SALDO AKHIR", "TOTAL", "HALAMAN", "REKENING",
        "STATEMENT", "PERIODE", "CABANG", "HORMAT KAMI", "CATATAN",
    ]

    # regex: DD/MM/YYYY atau DD-MM-YYYY atau DD/MM
    date_re = re.compile(r"^(\d{2}[/-]\d{2}[/-]\d{4}|\d{2}[/-]\d{2})\b")

    for raw_line in lines:
        line_str = raw_line.strip()
        if not line_str:
            continue
        if any(kw.lower() in line_str.lower() for kw in ignore_keywords):
            continue

        m_date = date_re.match(line_str)
        if not m_date:
            continue

        tgl = m_date.group(1)

        # cari semua angka berformat ribuan/desimal
        nums = re.findall(r"[\d.]+,\d{2}|[\d,]+\." + r"\d{2}", line_str)
        # fallback format umum
        if not nums:
            nums = re.findall(r"\d[\d.,]*\d", line_str)

        # heuristik: jika minimal 2 angka -> (nominal, saldo) atau (debet, kredit) atau (debet, saldo)
        nominal = 0.0
        saldo = None
        is_db = False

        # deteksi kata DB/CR untuk Maybank
        if re.search(r"\bDB\b", line_str.upper()):
            is_db = True

        # coba deteksi kolom debet/kredit berdasarkan posisi
        # Pola paling umum: ... <nominal> <saldo>  atau  <debet> <kredit> <saldo>
        parsed_nums = [to_float(n) for n in nums if to_float(n) is not None]
        parsed_nums = [n for n in parsed_nums if n is not None]

        if len(parsed_nums) >= 2:
            # anggap angka terakhir = saldo, angka pertama = nominal
            nominal = parsed_nums[0]
            saldo_candidate = parsed_nums[-1]
            if saldo_candidate != nominal:
                saldo = saldo_candidate
        elif len(parsed_nums) == 1:
            nominal = parsed_nums[0]
        else:
            # tidak ada nominal -> masih catat sebagai baris? skip
            continue

        # bersihkan keterangan
        desc = line_str
        for n in nums:
            desc = desc.replace(n, "")
        desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.IGNORECASE)
        desc = " ".join(desc.split()).strip()

        # SWAP: DB -> kredit, CR -> debit
        if is_db:
            debit_val = 0.0
            kredit_val = nominal
        else:
            debit_val = nominal
            kredit_val = 0.0

        transactions.append({
            "Tanggal": tgl,
            "Kode": "-",
            "Keterangan": desc,
            "Debit": debit_val,
            "Kredit": kredit_val,
            "Saldo": saldo,
            "Tipe Asli": "DB" if is_db else "CR",
        })

    return transactions


# =========================================================
# BUILD DATAFRAME
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
# UI STREAMLIT
# =========================================================
col1, col2, col3 = st.columns(3)
with col1:
    bank_choice = st.selectbox("Pilih Bank", ["BCA", "Maybank"])
with col2:
    password = st.text_input("🔑 Password PDF (jika ada)", type="password")
with col3:
    uploaded_files = st.file_uploader(
        "📄 Pilih file PDF", type=["pdf"], accept_multiple_files=True
    )

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
                st.success(f"✅ {uploaded_file.name}: {len(df)} transaksi terdeteksi.")
            else:
                st.warning(f"⚠️ {uploaded_file.name}: Tidak ada transaksi terdeteksi.")

        except ValueError as ve:
            st.error(f"🔒 {uploaded_file.name}: {ve}")
        except Exception as e:
            st.error(f"❌ Error saat memproses **{uploaded_file.name}**: {type(e).__name__}")
            with st.expander("Lihat detail error"):
                st.exception(e)

    if all_dfs:
        try:
            final_df = pd.concat(all_dfs, ignore_index=True)

            # Pengurutan kronologis
            final_df["Tanggal"] = final_df["Tanggal"].astype(str).str.strip()
            final_df["Tanggal_DT"] = pd.to_datetime(
                final_df["Tanggal"], format="%d/%m/%Y", errors="coerce"
            )
            final_df = final_df.sort_values(by=["Tanggal_DT"], ascending=True).reset_index(drop=True)
            final_df = final_df.drop(columns=["Tanggal_DT"])

            st.subheader("📊 Preview Data (Debet ↔ Kredit swapped)")
            st.dataframe(final_df, use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False, sheet_name="Mutasi_Swapped")
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Excel (Debet ↔ Kredit swapped)",
                data=processed_data,
                file_name="mutasi_bank_correct_saldo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Terjadi kesalahan saat menggabungkan data: {e}")
