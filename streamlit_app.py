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
st.title("🏦 Bank PDF → Excel Converter (OCR Table)")

# ==========================================
# FUNGSI BANTUAN
# ==========================================
def to_float_indo_bca(s):
    if not s: return None
    try: 
        clean = re.sub(r"[^\d.-]", "", s)
        return float(clean)
    except: return None

def to_float_indo_mb(s):
    if not s: return None
    try: return float(s.replace(".", "").replace(",", "."))
    except: return None

# ==========================================
# FUNGSI PEMBACA TABEL GAMBAR (OPENCV)
# ==========================================
def extract_table_from_image(image_pil):
    # 1. Konversi ke OpenCV format
    img = np.array(image_pil.convert('RGB'))
    
    # 2. Grayscale dan Thresholding
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)[1]
    
    # 3. Deteksi Garis Horizontal dan Vertikal
    # (Trik untuk membuang teks dan menyisakan garis tabel)
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40,1))
    detect_horizontal = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
    
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1,40))
    detect_vertical = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=2)
    
    # Gabungkan garis
    table_mask = cv2.addWeighted(detect_horizontal, 0.5, detect_vertical, 0.5, 0.0)
    
    # 4. Cari kontur (kotak-kotak) dari tabel
    contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # Urutkan kontur dari atas ke bawah
    bounding_boxes = [cv2.boundingRect(c) for c in contours]
    bounding_boxes.sort(key=lambda b: b[1]) # Sort by Y
    
    extracted_text = ""
    
    # 5. Potong gambar per baris tabel dan baca dengan OCR
    for (x, y, w, h) in bounding_boxes:
        if w > 100 and h > 20: # Abaikan garis kecil
            roi = gray[y:y+h, x:x+w]
            # Perbesar kotak agar OCR lancar
            roi = cv2.resize(roi, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
            # Baca teks di dalam kotak
            text = pytesseract.image_to_string(roi, lang='ind').strip()
            extracted_text += text + "\n"
            
    return extracted_text

# ==========================================
# PARSER BCA (Sama seperti sebelumnya)
# ==========================================
def parse_bca_ocr(text, tahun="2026"):
    transactions = []
    match_thn = re.search(r"PERIODE\s*:.*?(\d{2})-\d{2}-\d{2}", text, re.IGNORECASE)
    if match_thn: tahun = "20" + match_thn.group(1)
    
    lines = text.split("\n")
    merged_lines = []
    current_line = ""
    
    for line in lines:
        line = line.strip()
        if not line: continue
        if re.match(r"^\d{2}/\d{2}", line):
            if current_line: merged_lines.append(current_line)
            current_line = line
        else:
            if current_line: current_line += " " + line
    if current_line: merged_lines.append(current_line)

    for line in merged_lines:
        match = re.match(r"^(\d{2}/\d{2})(.*)", line)
        if not match: continue
        tgl_bulan, content = match.groups()
        content = content.strip()
        
        if "SALDO AWAL" in content.upper(): continue
        tgl = f"{tgl_bulan}/{tahun}"
        
        # Cari angka desimal .xx (toleran terhadap OCR)
        amounts = re.findall(r"-?\d{1,3}(?:[,.]\d{3})*[,.]\d{2}", content)
        
        debit, kredit, saldo = None, None, None
        if len(amounts) >= 2:
            saldo = to_float_indo_bca(amounts[-1])
            nominal = to_float_indo_bca(amounts[-2])
            if "DB" in content.upper(): kredit = nominal 
            else: debit = nominal
        elif len(amounts) == 1:
            nominal = to_float_indo_bca(amounts[0])
            if "DB" in content.upper(): kredit = nominal
            else: debit = nominal
        
        desc = content
        for amt in amounts: desc = desc.replace(amt, "")
        desc = re.sub(r"\bDB\b", "", desc, flags=re.IGNORECASE)
        desc = " ".join(desc.split())

        transactions.append({
            "Tanggal": tgl, "Kode": "-", "Keterangan": desc,
            "Debit": debit, "Kredit": kredit, "Saldo": saldo
        })
    return transactions

# ==========================================
# UI STREAMLIT
# ==========================================
col1, col2, col3 = st.columns(3)
with col1:
    bank_choice = st.selectbox("Pilih Bank", ["BCA", "Maybank"])
with col2:
    password = st.text_input("🔑 Password PDF (jika ada)", type="password")
with col3:
    uploaded_files = st.file_uploader("📄 Pilih file (PDF / JPG / PNG)", type=["pdf", "jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        try:
            text = ""
            
            if not uploaded_file.name.lower().endswith(".pdf"):
                st.info(f"📸 Memproses Tabel OCR untuk {uploaded_file.name}...")
                img = Image.open(uploaded_file)
                # Gunakan fungsi ekstraksi tabel gambar
                text = extract_table_from_image(img)
            else:
                # (Logika PDF tetap sama)
                pass

            with st.expander(f"🔍 Teks Asli Hasil Scan ({uploaded_file.name})"):
                st.text(text)

            if bank_choice == "BCA":
                txns = parse_bca_ocr(text)
                df = pd.DataFrame(txns)
            # ...
            
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
