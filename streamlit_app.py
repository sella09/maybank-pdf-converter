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
    if not text: return ""
    t = str(text)
    t = re.sub(r"Bersambung ke halaman berikut\.?", "", t, flags=re.I)
    t = re.sub(r"SALDO\s+AWAL\s*:.*?(SALDO\s+AKHIR\s*:.*)?", "", t, flags=re.I)
    t = re.sub(r"CATATAN\s*:.*?Rekening\.?", "", t, flags=re.I | re.S)
    t = re.sub(r"HALAMAN\s*:.*?IDR", "", t, flags=re.I)
    return " ".join(t.split()).strip()

# =========================================================
# PARSER BCA (TEKS UTUH LAYOUT)
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

        # Mengambil teks dengan layout asli (spasi terjaga)
        layout_text = page.extract_text(layout=True)
        if not layout_text: continue

        lines = layout_text.split("\n")
        current_txn = None

        for line in lines:
            line_str = line.strip()
            if not line_str: continue
            if any(k in line_str.upper() for k in ["REKENING GIRO", "KCIU RAWAMANGUN", "NO. REKENING", "PERIODE :", "MATA UANG", "CATATAN:", "BCA BERHAK"]):
                continue

            # Cek apakah baris ini memulai transaksi baru (diawali DD/MM)
            m_date = re.match(r"^(\d{2}/\d{2})\s+(.+)", line_str)
            
            if m_date:
                if current_txn:
                    current_txn["Keterangan"] = clean_desc(current_txn["Keterangan"])
                    transactions.append(current_txn)

                tgl = m_date.group(1)
                rest_text = m_date.group(2)

                # Ekstrak angka murni (format 1.000.000.00 atau 1000.00)
                nums_raw = re.findall(r"\b\d{1,3}(?:,\d{3})*\.\d{2}\b|\b\d+\.\d{2}\b", rest_text)
                nums = [to_float(n) for n in nums_raw]

                # Hapus angka dari teks untuk dijadikan keterangan
                desc_text = rest_text
                for n in nums_raw:
                    desc_text = desc_text.replace(n, "")
                
                upper_desc = desc_text.upper()

                # Deteksi DB/CR
                is_db = False
                if any(k in upper_desc for k in ["EDDY SUFENDI", "BIAYA ADM", "PAJAK BUNGA", "PENERIMAAN NEGARA", "BA JASA", "AUTODEBET", "ND-LAINNYA"]):
                    is_db = True
                elif "TRSF E-BANKING DB" in upper_desc:
                    is_db = True

                current_txn = {
                    "Tanggal": f"{tgl}/{tahun}",
                    "Kode": "-",
                    "Keterangan": desc_text,
                    "Debit": 0.0,
                    "Kredit": 0.0,
                    "Saldo": None,
                    "Nums": nums, # Simpan dulu angkanya
                    "_order": order,
                }
                order += 1

            else:
                # Ini baris lanjutan
                if current_txn:
                    current_txn["Keterangan"] += " " + line_str
                    
                    # Cek jika di baris lanjutan ini ternyata ada angka (misal 7194000.00)
                    nums_raw = re.findall(r"\b\d{1,3}(?:,\d{3})*\.\d{2}\b|\b\d+\.\d{2}\b", line_str)
                    if nums_raw:
                        for n in nums_raw:
                            current_txn["Nums"].append(to_float(n))
                            current_txn["Keterangan"] = current_txn["Keterangan"].replace(n, "")

        if current_txn:
            current_txn["Keterangan"] = clean_desc(current_txn["Keterangan"])
            transactions.append(current_txn)
            current_txn = None

    pdf.close()

    # PROSES AKHIR: TENTUKAN NOMINAL, SALDO, DAN DEBIT/KREDIT
    for txn in transactions:
        nums = txn["Nums"]
        del txn["Nums"] # Hapus dari output akhir

        upper_text = txn["Keterangan"].upper()

        # Aturan Utama BCA
        is_db = False
        if any(k in upper_text for k in ["EDDY SUFENDI", "BIAYA ADM", "PAJAK BUNGA", "PENERIMAAN NEGARA", "BA JASA", "AUTODEBET", "ND-LAINNYA", "SWITCHING KE", "SWITCHING TRANSFER KE", "TRSF E-BANKING DB"]):
            is_db = True
            
        # Tentukan Nominal dan Saldo
        # Asumsi: Angka pertama = Nominal, Angka terakhir = Saldo
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

        if "BUNGA" in upper_text and "PAJAK" not in upper_text:
            debit_val, kredit_val = 0.0, nominal

        txn["Debit"] = debit_val
        txn["Kredit"] = kredit_val
        txn["Saldo"] = saldo

        # Bersihkan sisa teks
        txn["Keterangan"] = re.sub(r"\bDB\b|\bCR\b", "", txn["Keterangan"]).strip()
        txn["Keterangan"] = re.sub(r"\b\d{4}/[A-Z0-9/]+\b", "", txn["Keterangan"]).strip()
        txn["Keterangan"] = " ".join(txn["Keterangan"].split())

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
            file_name="mutasi_bank_fixed_grid_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
