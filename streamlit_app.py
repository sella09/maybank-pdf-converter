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
    if num_str is None or pd.isna(num_str): return 0.0
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
    if not text: return ""
    t = str(text)
    t = re.sub(r"Bersambung ke halaman berikut\.?", "", t, flags=re.I)
    t = re.sub(r"SALDO\s+AWAL\s*:.*?(SALDO\s+AKHIR\s*:.*)?", "", t, flags=re.I)
    t = re.sub(r"CATATAN\s*:.*?Rekening\.?", "", t, flags=re.I | re.S)
    t = re.sub(r"HALAMAN\s*:.*?IDR", "", t, flags=re.I)
    return " ".join(t.split()).strip()

# =========================================================
# PARSER BCA (LAYOUT TEXT)
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

        layout_text = page.extract_text(layout=True)
        if not layout_text: continue

        lines = layout_text.split("\n")
        current_txn = None

        for line in lines:
            line_str = line.strip()
            if not line_str: continue
            if any(k in line_str.upper() for k in ["REKENING GIRO", "KCIU RAWAMANGUN", "NO. REKENING", "PERIODE :", "MATA UANG", "CATATAN:", "BCA BERHAK"]):
                continue

            m_date = re.match(r"^(\d{2}/\d{2})\s+(.+)", line_str)
            
            if m_date:
                if current_txn:
                    current_txn["Keterangan"] = clean_desc(current_txn["Keterangan"])
                    transactions.append(current_txn)

                tgl = m_date.group(1)
                rest_text = m_date.group(2)

                parts = rest_text.split()
                desc_words = []
                nums = []
                for p in parts:
                    if re.match(r"^\d[\d,]*\.\d{2}$", p):
                        nums.append(to_float(p))
                    else:
                        desc_words.append(p)
                
                desc_text = " ".join(desc_words)
                upper_desc = desc_text.upper()

                is_db = False
                if any(k in upper_desc for k in ["EDDY SUFENDI", "BIAYA ADM", "PAJAK BUNGA", "PENERIMAAN NEGARA", "BA JASA", "AUTODEBET", "ND-LAINNYA"]):
                    is_db = True
                elif "TRSF E-BANKING DB" in upper_desc or "SWITCHING DB" in upper_desc:
                    is_db = True

                if len(nums) >= 2:
                    nominal = nums[0]
                    saldo = nums[-1]
                elif len(nums) == 1:
                    nominal = nums[0]
                    saldo = None
                else:
                    nominal = 0.0
                    saldo = None

                debit_val = nominal if is_db else 0.0
                kredit_val = nominal if not is_db else 0.0

                if "BUNGA" in upper_desc and "PAJAK" not in upper_desc:
                    debit_val, kredit_val = 0.0, nominal

                current_txn = {
                    "Tanggal": f"{tgl}/{tahun}",
                    "Kode": "-",
                    "Keterangan": desc_text,
                    "Debit": debit_val,
                    "Kredit": kredit_val,
                    "Saldo": saldo,
                    "_order": order,
                }
                order += 1

            else:
                if current_txn:
                    extra_parts = line_str.split()
                    for p in extra_parts:
                        if re.match(r"^\d[\d,]*\.\d{2}$", p):
                            val = to_float(p)
                            if current_txn["Saldo"] is None and current_txn["Debit"] == 0 and current_txn["Kredit"] == 0:
                                current_txn["Saldo"] = val
                            elif current_txn["Debit"] == 0 and current_txn["Kredit"] == 0:
                                current_txn["Debit"] = val
                        else:
                            current_txn["Keterangan"] += " " + p

        if current_txn:
            current_txn["Keterangan"] = clean_desc(current_txn["Keterangan"])
            transactions.append(current_txn)
            current_txn = None

    pdf.close()

    # Validasi Saldo BCA
    saldo_berjalan = 0.0
    for txn in transactions:
        if txn["Saldo"] is None:
            txn["Saldo"] = round(saldo_berjalan - txn["Debit"] + txn["Kredit"], 2)
        else:
            if saldo_berjalan != 0.0:
                expected = round(saldo_berjalan - txn["Debit"] + txn["Kredit"], 2)
                if abs(expected - txn["Saldo"]) > 1.0:
                    expected_swap = round(saldo_berjalan + txn["Debit"] - txn["Kredit"], 2)
                    if abs(expected_swap - txn["Saldo"]) <= 1.0:
                        txn["Debit"], txn["Kredit"] = txn["Kredit"], txn["Debit"]
        
        saldo_berjalan = txn["Saldo"]

        txn["Keterangan"] = re.sub(r"\bDB\b|\bCR\b", "", txn["Keterangan"]).strip()
        txn["Keterangan"] = re.sub(r"\b\d{4}/[A-Z0-9/]+\b", "", txn["Keterangan"]).strip()
        txn["Keterangan"] = " ".join(txn["Keterangan"].split())

        if txn["Debit"] > 0: txn["Tipe Asli"] = "DB"
        elif txn["Kredit"] > 0: txn["Tipe Asli"] = "CR"
        else: txn["Tipe Asli"] = "-"

    return transactions

# =========================================================
# PARSER MAYBANK (TABEL GRID)
# =========================================================
@st.cache_data(show_spinner=False)
def parse_maybank(file_bytes, password=None):
    file_obj = BytesIO(file_bytes)
    pdf = open_pdf(file_obj, password)
    transactions = []
    order = 0

    for page in pdf.pages:
        # Baca tabel langsung
        tables = page.extract_tables({
            "vertical_strategy": "lines", 
            "horizontal_strategy": "lines"
        })
        
        # Jika tidak ada garis tabel, coba mode teks
        if not tables:
            tables = page.extract_tables({
                "vertical_strategy": "text", 
                "horizontal_strategy": "text"
            })

        for table in tables:
            for row in table:
                if not row or len(row) < 4: continue
                
                # Cari tanggal (format 01/09/2026)
                tgl_match = None
                for cell in row:
                    if cell and re.match(r"^\d{2}/\d{2}/\d{4}$", str(cell).strip()):
                        tgl_match = str(cell).strip()
                        break
                
                if not tgl_match: continue

                # Ambil semua angka desimal di baris
                full_text = " ".join([str(c).replace("\n", " ") for c in row if c])
                nums = re.findall(r"\d[\d,]*\.\d{2}", full_text)
                parsed_nums = [to_float(n) for n in nums]

                if len(parsed_nums) >= 2:
                    nominal = parsed_nums[0]
                    saldo = parsed_nums[-1]
                elif len(parsed_nums) == 1:
                    nominal = parsed_nums[0]
                    saldo = None
                else:
                    nominal = 0.0
                    saldo = None

                # Bersihkan Teks Keterangan (Gabungkan semua sel kecuali tanggal)
                desc_text = full_text.replace(tgl_match, "")
                for n in nums:
                    desc_text = desc_text.replace(n, "")
                
                # Cek Debit/Kredit
                is_db = False
                if "DEBIT" in full_text.upper() and "KREDIT" in full_text.upper():
                    # Jika ada judul DEBIT dan KREDIT, cek posisi selnya
                    pass

                # Keterangan default
                current_txn = {
                    "Tanggal": tgl_match,
                    "Kode": "-",
                    "Keterangan": clean_desc(desc_text),
                    "Debit": nominal if is_db else 0.0,
                    "Kredit": nominal if not is_db else 0.0,
                    "Saldo": saldo,
                    "_order": order,
                }
                
                # Deteksi Debit/Kredit dengan validasi Saldo
                if saldo is not None and len(transactions) > 0:
                    prev_saldo = transactions[-1]["Saldo"]
                    if prev_saldo is not None:
                        # Cek apakah masuk Debit atau Kredit
                        if round(prev_saldo - nominal, 2) == round(saldo, 2):
                            current_txn["Debit"] = nominal
                            current_txn["Kredit"] = 0.0
                        elif round(prev_saldo + nominal, 2) == round(saldo, 2):
                            current_txn["Debit"] = 0.0
                            current_txn["Kredit"] = nominal

                transactions.append(current_txn)
                order += 1

    pdf.close()

    # Pembersihan Akhir
    for txn in transactions:
        txn["Keterangan"] = clean_desc(txn["Keterangan"])
        txn["Keterangan"] = re.sub(r"\b\d{2}/\d{2}/\d{4}\b", "", txn["Keterangan"]).strip()
        txn["Keterangan"] = " ".join(txn["Keterangan"].split())

        if txn["Debit"] > 0: txn["Tipe Asli"] = "DB"
        elif txn["Kredit"] > 0: txn["Tipe Asli"] = "CR"
        else: txn["Tipe Asli"] = "-"

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
            file_name="mutasi_bank_fixed_indo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
