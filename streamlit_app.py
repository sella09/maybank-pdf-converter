import streamlit as st
import pdfplumber
import pandas as pd
import re
import pytesseract
import cv2
import numpy as np
from PIL import Image
from pdf2image import convert_from_bytes
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (OCR Enhanced)")

# ==========================================
# FUNGSI BANTUAN
# ==========================================
def to_float_indo_mb(s):
    if not s: return None
    try: return float(s.replace(".", "").replace(",", "."))
    except: return None

def to_float_indo_bca(s):
    if not s: return None
    try: 
        clean = re.sub(r"[^\d.-]", "", s)
        return float(clean)
    except: return None

# ==========================================
# FUNGSI PRE-PROCESSING GAMBAR UNTUK OCR
# ==========================================
def preprocess_image(image_pil):
    # Konversi PIL ke format OpenCV
    img = np.array(image_pil)
    
    # 1. Grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    
    # 2. Thresholding (Hitam Putih tegas)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
    
    # 3. Naikkan Resolusi (Zoom 2x) agar huruf kecil tajam
    resized = cv2.resize(thresh, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    
    return Image.fromarray(resized)

# ==========================================
# PARSER BCA (Khusus OCR)
# ==========================================
def parse_bca_ocr(text, tahun="2026"):
    transactions = []
    
    # Ambil tahun dari "PERIODE : 31-08-26"
    match_thn = re.search(r"PERIODE\s*:.*?(\d{2})-\d{2}-\d{2}", text, re.IGNORECASE)
    if match_thn: 
        tahun = "20" + match_thn.group(1)

    # Potong teks hanya dari "SALDO AWAL" ke bawah
    text_lower = text.upper()
    idx_saldo_awal = text_lower.find("SALDO AWAL")
    if idx_saldo_awal != -1:
        text = text[idx_saldo_awal:]
    
    lines = text.split("\n")
    merged_lines = []
    current_line = ""
    
    for line in lines:
        line = line.strip()
        if not line: continue
        
        # Baris baru jika diawali DD/MM
        if re.match(r"^\d{2}/\d{2}", line):
            if current_line: merged_lines.append(current_line)
            current_line = line
        else:
            if current_line: current_line += " " + line

    if current_line: merged_lines.append(current_line)

    for line in merged_lines:
        match = re.match(r"^(\d{2}/\d{2})(.*)", line)
        if not match: continue
        
        tgl_bulan = match.group(1)
        content = match.group(2).strip()
        
        if "SALDO AWAL" in content.upper(): continue
        
        tgl = f"{tgl_bulan}/{tahun}"
        
        # Cari angka desimal .xx
        amounts = re.findall(r"-?\d{1,3}(?:[,.]\d{3})*[,.]\d{2}", content)
        
        debit, kredit, saldo = None, None, None
        
        if len(amounts) >= 2:
            # Angka paling kanan adalah Saldo, angka sebelumnya Mutasi
            saldo_str = amounts[-1]
            nominal_str = amounts[-2]
            
            saldo = to_float_indo_bca(saldo_str)
            nominal = to_float_indo_bca(nominal_str)
            
            if "DB" in content.upper():
                kredit = nominal 
            else:
                debit = nominal
                
        elif len(amounts) == 1:
            nominal = to_float_indo_bca(amounts[0])
            if "DB" in content.upper(): kredit = nominal
            else: debit = nominal
        
        # Bersihkan Keterangan
        desc = content
        for amt in amounts: desc = desc.replace(amt, "")
        desc = re.sub(r"\bDB\b", "", desc, flags=re.IGNORECASE)
        desc = " ".join(desc.split())

        transactions.append({
            "Tanggal": tgl,
            "Kode": "-",
            "Keterangan": desc,
            "Debit": debit,
            "Kredit": kredit,
            "Saldo": saldo
        })
            
    return transactions

def build_df_bca(transactions, sumber_file):
    rows = []
    for txn in transactions:
        rows.append({
            "Sumber File": sumber_file,
            "Tanggal": txn["Tanggal"],
            "Kode": txn["Kode"],
            "Keterangan": txn["Keterangan"],
            "Debit": txn["Debit"],
            "Kredit": txn["Kredit"],
            "Saldo": txn["Saldo"]
        })
    return pd.DataFrame(rows)

# ==========================================
# UI STREAMLIT
# ==========================================
col1, col2 = st.columns(2)
with col1:
    bank_choice = st.selectbox("Pilih Bank", ["BCA", "Maybank"])
with col2:
    uploaded_files = st.file_uploader("📄 Pilih file gambar (JPG/PNG) atau PDF", type=["pdf", "jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        try:
            text = ""
            
            # JIKA GAMBAR
            if not uploaded_file.name.lower().endswith(".pdf"):
                st.info(f"📸 Memproses OCR untuk {uploaded_file.name}...")
                img = Image.open(uploaded_file).convert('RGB')
                
                # Proses Gambar (Tajamkan & Perbesar)
                processed_img = preprocess_image(img)
                
                # OCR
                text = pytesseract.image_to_string(processed_img, lang='ind')
            
            # JIKA PDF
            else:
                with pdfplumber.open(uploaded_file) as pdf:
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if not page_text or len(page_text.strip()) < 50:
                            images = convert_from_bytes(uploaded_file.read())
                            for img in images:
                                processed_img = preprocess_image(img)
                                page_text += pytesseract.image_to_string(processed_img, lang='ind') + "\n"
                            break
                        text += page_text + "\n"

            # Tampilkan teks mentah hasil scan
            with st.expander(f"🔍 Teks Asli Hasil Scan ({uploaded_file.name})"):
                st.text(text)

            if bank_choice == "BCA":
                txns = parse_bca_ocr(text)
                df = build_df_bca(txns, uploaded_file.name)
            
            if not df.empty:
                all_dfs.append(df)
                st.success(f"✅ {uploaded_file.name}: {len(df)} transaksi terbaca.")
            else:
                st.warning(f"⚠️ {uploaded_file.name}: Gagal membaca transaksi. Cek teks asli di atas.")
                
        except Exception as e:
            st.error(f"❌ **{uploaded_file.name}**: {e}")

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        st.subheader("📊 Preview Data")
        st.dataframe(final_df, width='stretch')
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Mutasi')
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 Download sebagai Excel",
            data=processed_data,
            file_name="mutasi_gabungan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
