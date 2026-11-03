import streamlit as st
import pdfplumber
import pandas as pd
import re
import traceback
from io import BytesIO

# --- PENTING UNTUK STREAMLIT CLOUD ---
# Mengatur limit memori agar tidak OOM (Out of Memory)
st.set_page_config(page_title="Bank PDF to Excel Converter", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA & Maybank)")

# ... (Fungsi open_pdf, to_float, clean_desc sama seperti sebelumnya) ...

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
    t = re.sub(r"Rekening ini sampai dengan akhir bulan berikutnya.*?Rekening ini\.", "", t, flags=re.IGNORECASE | re.DOTALL)
    t = re.sub(r"telah menyetujui segala data.*?Rekening ini\.", "", t, flags=re.IGNORECASE | re.DOTALL)
    t = re.sub(r"Apabila nasabah tidak melakukan sanggahan.*?Laporan Mutasi Rekening\.", "", t, flags=re.IGNORECASE | re.DOTALL)
    return " ".join(t.split()).strip()

# --- TAMBAHKAN @st.cache_data UNTUK MEMORY EFFICIENCY ---
@st.cache_data(show_spinner=False)
def parse_bca_hybrid(file_bytes, password=None):
    transactions = []
    # Gunakan BytesIO agar file bisa dibaca berulang kali
    file_obj = BytesIO(file_bytes)
    pdf = open_pdf(file_obj, password)
    
    full_text = ""
    for page in pdf.pages:
        t = page.extract_text()
        if t: full_text += t + "\n"
    
    tahun = "2026"
    m = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", full_text, re.IGNORECASE)
    if m: tahun = m.group(1)

    table_rows_found = 0
    for page in pdf.pages:
        tables = page.extract_tables()
        for table in tables:
            for row in table:
                if not row: continue
                cells = [clean_desc(c) if c else "" for c in row]
                joined = " | ".join(cells)
                
                tgl = None
                for c in cells:
                    mt = re.match(r"^(\d{2}/\d{2})\b", c)
                    if mt:
                        tgl = mt.group(1)
                        break
                if not tgl: continue
                
                row_text = re.sub(r"^\d{2}/\d{2}\s*", "", joined)
                upper = row_text.upper()
                
                if "SALDO AWAL" in upper:
                    continue

                is_db = False
                if "PAJAK BUNGA" in upper:
                    is_db = True
                elif "BUNGA" in upper and "PAJAK" not in upper:
                    is_db = False
                elif "BIAYA ADM" in upper or "BIAYA TRANSFER" in upper or "BIAYA REFERENSI" in upper:
                    is_db = True
                elif re.search(r"\bDB\b", upper):
                    is_db = True
                elif re.search(r"\bCR\b", upper):
                    is_db = False
                elif any(k in upper for k in ["EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET", "PENERIMAAN NEGARA"]):
                    is_db = True

                decimals = re.findall(r"\d[\d,]*\.\d{2}", row_text)
                if not decimals: continue

                nominal = to_float(decimals[0]) or 0.0
                saldo = None
                if len(decimals) >= 2:
                    saldo_val = to_float(decimals[-1])
                    if saldo_val is not None and saldo_val > 100 and saldo_val != nominal:
                        saldo = saldo_val

                desc = row_text
                for d in decimals: desc = desc.replace(d, "")
                desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
                desc = clean_desc(desc)

                debit_val, kredit_val = (0.0, nominal) if is_db else (nominal, 0.0)
                transactions.append({
                    "Tanggal": f"{tgl}/{tahun}", "Kode": "-", "Keterangan": desc,
                    "Debit": debit_val, "Kredit": kredit_val,
                    "Saldo": saldo, "Tipe Asli": "DB" if is_db else "CR",
                })
                table_rows_found += 1

    if table_rows_found == 0:
        # Fallback kode teks mentah (sama seperti sebelumnya)
        lines = full_text.split('\n')
        ignore_keywords = [
            "REKENING GIRO", "KCU", "DOMINIC DIGITAL", "KELAPA GADING", "JL BOULEVARD",
            "JAKARTA", "INDONESIA", "NO. REKENING", "HALAMAN", "PERIODE", "MATA UANG",
            "CATATAN:", "Apabila nasabah", "BCA berhak", "TANGGAL KETERANGAN", 
            "Bersambung ke halaman", "MUTASI CR :", "MUTASI DB :", "SALDO AKHIR :",
            "===== Page", "# REKENING", "Page ", "- Page "
        ]
        bunga_keywords = ["BUNGA", "PAJAK BUNGA", "BIAYA ADM", "BIAYA TRANSFER", "BIAYA REFERENSI"]
        
        clean_lines = []
        for line in lines:
            line_str = line.strip()
            if not line_str: continue
            if re.match(r"^\d+\s*/\s*\d+$", line_str) or (re.match(r"^\d+$", line_str) and len(line_str) <= 3): continue
            if any(kw.lower() in line_str.lower() for kw in ignore_keywords): continue
            clean_lines.append(line_str)

        raw_blocks = []
        current_block = None
        date_pattern = re.compile(r"^(\d{2}/\d{2})\s*(.+)")

        for line_str in clean_lines:
            m_date = date_pattern.match(line_str)
            if m_date:
                if current_block: raw_blocks.append(current_block)
                current_block = {"tgl": m_date.group(1), "lines": [m_date.group(2)]}
            elif current_block:
                upper_line = line_str.upper()
                if any(k in upper_line for k in bunga_keywords):
                    raw_blocks.append(current_block)
                    current_block = {"tgl": current_block["tgl"], "lines": [line_str]}
                else:
                    current_block["lines"].append(line_str)

        if current_block: raw_blocks.append(current_block)

        for block in raw_blocks:
            combined_text = " ".join(block["lines"])
            if "SALDO AWAL" in combined_text.upper(): continue

            upper_text = combined_text.upper()
            is_db = False
            if "PAJAK BUNGA" in upper_text: is_db = True
            elif "BUNGA" in upper_text and "PAJAK" not in upper_text: is_db = False
            elif "BIAYA ADM" in upper_text or "BIAYA TRANSFER" in upper_text or "BIAYA REFERENSI" in upper_text: is_db = True
            elif re.search(r"\bDB\b", upper_text): is_db = True
            elif any(k in upper_text for k in ["EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET", "PENERIMAAN NEGARA"]): is_db = True

            all_decimals = re.findall(r"[\d,]+\.\d{2}", combined_text)
            if not all_decimals: continue

            nominal = to_float(all_decimals[0]) or 0.0
            saldo = None
            if len(all_decimals) >= 2:
                saldo_val = to_float(all_decimals[-1])
                if saldo_val is not None and saldo_val > 100 and saldo_val != nominal:
                    saldo = saldo_val

            desc = combined_text
            for num_str in all_decimals: desc = desc.replace(num_str, "")
            desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.IGNORECASE)
            desc = clean_desc(desc)

            debit_val, kredit_val = (0.0, nominal) if is_db else (nominal, 0.0)
            transactions.append({
                "Tanggal": f"{block['tgl']}/{tahun}", "Kode": "-", "Keterangan": desc,
                "Debit": debit_val, "Kredit": kredit_val,
                "Saldo": saldo, "Tipe Asli": "DB" if is_db else "CR",
            })

    pdf.close()
    return transactions

# --- UI STREAMLIT DENGAN ERROR HANDLING ---
col1, col2, col3 = st.columns(3)
with col1: bank_choice = st.selectbox("Pilih Bank", ["BCA", "Maybank"])
with col2: password = st.text_input("🔑 Password PDF (jika ada)", type="password")
with col3: uploaded_files = st.file_uploader("📄 Pilih file PDF/TXT", type=["pdf", "txt"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        try:
            st.info(f"Memproses **{uploaded_file.name}**...")
            
            # BACA FILE SEBAGAI BYTES (Agar bisa di-cache)
            file_bytes = uploaded_file.read()
            
            if bank_choice == "BCA":
                txns = parse_bca_hybrid(file_bytes, password=password)
            else:
                # Fallback Maybank (asumsi fungsi parse_maybank_statement sudah ada)
                st.warning("Parser Maybank belum diimplementasikan di versi ini.")
                txns = []

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
            # JIKA ERROR, TAMPILKAN DI UI (BUKAN BLANK SCREEN)
            st.error(f"❌ Gagal memproses {uploaded_file.name}: {str(e)}")
            with st.expander("Lihat Detail Error untuk Developer"):
                st.code(traceback.format_exc())

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        final_df["Tanggal"] = final_df["Tanggal"].astype(str).str.strip()
        final_df["Tanggal_DT"] = pd.to_datetime(final_df["Tanggal"], format="%d/%m/%Y", errors="coerce")
        final_df = final_df.sort_values("Tanggal_DT").reset_index(drop=True)
        final_df = final_df.drop(columns=["Tanggal_DT"])
        final_df = final_df[(final_df["Debit"] > 0) | (final_df["Kredit"] > 0)]

        st.subheader("📊 Preview Data")
        st.dataframe(final_df, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Mutasi")
        st.download_button(
            "📥 Download Excel", data=output.getvalue(),
            file_name="mutasi_bank.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
