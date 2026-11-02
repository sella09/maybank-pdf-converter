import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA Final)")

# ==========================================
# FUNGSI BANTUAN
# ==========================================
def to_float_indo_bca(s):
    if not s: return None
    try: 
        # Hapus semua karakter selain angka, minus, dan titik
        clean = re.sub(r"[^\d.-]", "", s)
        return float(clean)
    except: return None

# ==========================================
# PARSER BCA (Versi 16.0 - Teks Mentah Agresif)
# ==========================================
def parse_bca_pdf(file_obj, password=None):
    transactions = []
    pdf_args = {"password": password} if password else {}
    
    with pdfplumber.open(file_obj, **pdf_args) as pdf:
        tahun = "2026"
        # 1. Ambil Teks Mentah dari semua halaman
        all_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"

        # 2. Cari Tahun dari "PERIODE : JANUARI 2026"
        match_thn = re.search(r"PERIODE\s*:\s*\w+\s+(\d{4})", all_text, re.IGNORECASE)
        if match_thn: 
            tahun = match_thn.group(1)

        # 3. Potong teks hanya dari "SALDO AWAL" ke bawah
        idx_awal = all_text.upper().find("SALDO AWAL")
        if idx_awal != -1:
            all_text = all_text[idx_awal:]

        # 4. Bersihkan Teks Sampah (Header/Footer BCA)
        all_text = re.sub(r"Bersambung.*?(TANGGAL|$)", "", all_text, flags=re.DOTALL|re.IGNORECASE)
        all_text = re.sub(r"REKENING GIRO.*?(MATA UANG : IDR|INDONESIA)", "", all_text, flags=re.DOTALL|re.IGNORECASE)
        all_text = re.sub(r"TANGGAL\s+KETERANGAN\s+CBG\s+MUTASI\s+SALDO", "", all_text, flags=re.IGNORECASE)
        
        # 5. Gabungkan baris yang terpotong
        lines = all_text.split("\n")
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

        # 6. Ekstrak Data
        urutan = 0
        for line in merged_lines:
            match = re.match(r"^(\d{2}/\d{2})(.*)", line)
            if not match: continue
            
            tgl_bulan = match.group(1)
            content = match.group(2).strip()
            
            if "SALDO AWAL" in content.upper(): continue
            
            # Cari semua angka desimal (mendukung titik dan koma)
            amounts = re.findall(r"-?\d{1,3}(?:[,.]\d{3})*[,.]\d{2}", content)
            
            saldo = None
            debit = None
            kredit = None
            desc = content

            if len(amounts) >= 1:
                # Angka terakhir adalah Saldo
                saldo_str = amounts[-1]
                saldo = to_float_indo_bca(saldo_str)
                desc = desc.replace(saldo_str, "")
                
                if len(amounts) >= 2:
                    # Angka sebelumnya adalah Nominal Mutasi
                    nominal_str = amounts[-2]
                    nominal = to_float_indo_bca(nominal_str)
                    desc = desc.replace(nominal_str, "")
                    
                    # Sesuai permintaanmu: Debit/Kredit ditukar
                    if "DB" in content.upper():
                        kredit = nominal
                    else:
                        debit = nominal
            
            # Bersihkan Deskripsi
            desc = re.sub(r"\bDB\b", "", desc, flags=re.IGNORECASE)
            desc = " ".join(desc.split()).strip(" ,-")
            
            if not desc and debit is None and kredit is None:
                continue

            tgl_full = f"{tgl_bulan}/{tahun}"
            urutan += 1

            transactions.append({
                "Urutan": urutan,
                "Tanggal": tgl_full,
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
            "Urutan": txn["Urutan"],
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
            txns = parse_bca_pdf(uploaded_file, password=password)
            df = build_df_bca(txns, uploaded_file.name)
            
            if not df.empty:
                all_dfs.append(df)
                st.success(f"✅ {uploaded_file.name}: {len(df)} transaksi")
            else:
                st.warning(f"⚠️ {uploaded_file.name}: tidak ada transaksi terdeteksi.")
                
        except Exception as e:
            st.error(f"❌ **{uploaded_file.name}**: {type(e).__name__}")
            with st.expander("Lihat detail error"):
                st.exception(e)

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        
        # FITUR URUTKAN
        try:
            final_df['Tanggal_DT'] = pd.to_datetime(final_df['Tanggal'], format='%d/%m/%Y', errors='coerce')
            final_df = final_df.sort_values(by=['Tanggal_DT', 'Urutan'])
            final_df = final_df.drop(columns=['Tanggal_DT', 'Urutan'])
        except Exception as e:
            st.warning(f"Gagal mengurutkan tanggal: {e}")

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
