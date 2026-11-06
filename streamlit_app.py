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
    return " ".join(t.split()).strip()

# =========================================================
# PARSER MAYBANK (POSISI KOLOM X)
# =========================================================
@st.cache_data(show_spinner=False)
def parse_maybank(file_bytes, password=None):
    file_obj = BytesIO(file_bytes)
    pdf = open_pdf(file_obj, password)
    transactions = []
    order = 0

    for page in pdf.pages:
        words = page.extract_words()
        
        # Kelompokkan per baris (Y)
        lines = {}
        for w in words:
            y = round(w['top'], 1)
            if y not in lines: lines[y] = []
            lines[y].append(w)

        current_txn = None

        for y in sorted(lines.keys()):
            row = sorted(lines[y], key=lambda w: w['x0'])
            row_text = " ".join([w['text'] for w in row])

            # Abaikan header
            if any(k in row_text.upper() for k in ["TGL KODE KETERANGAN", "TRANSAKSI REKENING GIRO"]):
                continue

            # Cek tanggal (format 01/09/2026)
            m_date = re.match(r"^(\d{2}/\d{2}/\d{4})\s+(.+)", row_text)
            
            if m_date:
                if current_txn:
                    transactions.append(current_txn)

                tgl = m_date.group(1)
                rest_text = m_date.group(2)

                # Pisahkan berdasarkan X (Kolom)
                # Asumsi X: Tgl < 50, Kode < 100, Keterangan < 400, Debit < 500, Kredit < 600, Saldo > 600
                desc_parts = []
                debit_val = 0.0
                kredit_val = 0.0
                saldo = None
                
                for w in row:
                    txt = w['text']
                    x = w['x0']
                    
                    if re.match(r"^\d[\d,.]*$", txt) and x > 350:
                        val = to_float(txt)
                        if x > 600:   # Kolom Saldo
                            saldo = val
                        elif x > 480: # Kolom Kredit
                            kredit_val = val
                        elif x > 350: # Kolom Debit
                            debit_val = val
                    else:
                        if x < 350:   # Kolom Keterangan
                            desc_parts.append(txt)
                
                # Bersihkan teks keterangan
                desc_text = " ".join(desc_parts)
                desc_text = re.sub(r"\b\d{2}/\d{2}/\d{4}\b", "", desc_text) # hapus sisa tanggal
                
                # VALIDASI MATEMATIKA UNTUK MEMASTIKAN DEBIT/KREDIT BENAR
                if saldo is not None and len(transactions) > 0:
                    prev_saldo = transactions[-1]["Saldo"]
                    if prev_saldo is not None:
                        # Jika Debit > 0 tapi secara hitungan harusnya Kredit
                        if debit_val > 0 and round(prev_saldo + debit_val, 2) == round(saldo, 2):
                            kredit_val = debit_val
                            debit_val = 0.0
                        # Jika Kredit > 0 tapi secara hitungan harusnya Debit
                        elif kredit_val > 0 and round(prev_saldo - kredit_val, 2) == round(saldo, 2):
                            debit_val = kredit_val
                            kredit_val = 0.0

                current_txn = {
                    "Tanggal": tgl,
                    "Kode": "-",
                    "Keterangan": clean_desc(desc_text),
                    "Debit": debit_val,
                    "Kredit": kredit_val,
                    "Saldo": saldo,
                    "_order": order,
                }
                order += 1
                
            else:
                # Baris lanjutan keterangan
                if current_txn:
                    # Cek jika ada angka nyasar di baris ini
                    for w in row:
                        txt = w['text']
                        x = w['x0']
                        if re.match(r"^\d[\d,.]*$", txt) and x > 350:
                            val = to_float(txt)
                            if x > 600 and current_txn["Saldo"] is None: current_txn["Saldo"] = val
                            elif x > 480 and current_txn["Kredit"] == 0: current_txn["Kredit"] = val
                            elif x > 350 and current_txn["Debit"] == 0: current_txn["Debit"] = val
                        else:
                            if x < 350:
                                current_txn["Keterangan"] += " " + txt

        if current_txn:
            transactions.append(current_txn)
            current_txn = None

    pdf.close()

    # Pembersihan Akhir
    for txn in transactions:
        txn["Keterangan"] = clean_desc(txn["Keterangan"])
        if txn["Debit"] > 0: txn["Tipe Asli"] = "DB"
        elif txn["Kredit"] > 0: txn["Tipe Asli"] = "CR"
        else: txn["Tipe Asli"] = "-"

    return transactions

# =========================================================
# PARSER BCA (TETAP SEPERTI SEBELUMNYA)
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
            file_name="mutasi_bank_final_indo_fixed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
