import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Bank PDF to Excel", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (BCA Fixed)")

def to_float_indo_bca(s):
    if not s: return None
    try: 
        clean = re.sub(r"[^\d.-]", "", s)
        return float(clean)
    except: return None

def parse_bca_final(file_obj, password=None):
    transactions = []
    pdf_args = {"password": password} if password else {}
    
    with pdfplumber.open(file_obj, **pdf_args) as pdf:
        tahun = "2026"
        page1_text = pdf.pages[0].extract_text() or ""
        match_thn = re.search(r"PERIODE\s*:\s*\w+\s+(\d{4})", page1_text, re.IGNORECASE)
        if match_thn: tahun = match_thn.group(1)

        for page in pdf.pages:
            words = page.extract_words()
            
            # Kelompokkan kata berdasarkan baris (Y)
            lines = {}
            for w in words:
                y_key = round(w['top'] / 5) * 5
                if y_key not in lines:
                    lines[y_key] = []
                lines[y_key].append(w)

            for y, word_list in sorted(lines.items()):
                # Urutkan kata dari kiri ke kanan
                word_list.sort(key=lambda w: w['x0'])
                
                # Gabungkan semua teks di baris ini
                full_text = " ".join([w['text'] for w in word_list])
                
                # Cek apakah baris ini diawali tanggal DD/MM
                # Kadang ada karakter aneh sebelum tanggal, jadi kita cari tanggal di mana saja
                date_match = re.search(r"(\d{2}/\d{2})", full_text)
                if not date_match: continue
                
                tgl_bulan = date_match.group(1)
                
                # Ambil semua angka desimal di baris ini
                numbers = re.findall(r"-?\d{1,3}(?:[,.]\d{3})*\.\d{2}", full_text)
                
                # Hapus tanggal dan angka dari teks untuk mendapatkan keterangan
                desc = full_text
                desc = desc.replace(tgl_bulan, "") # Hapus tanggal
                for n in numbers: desc = desc.replace(n, "") # Hapus angka
                desc = re.sub(r"\bDB\b", "", desc, flags=re.IGNORECASE) # Hapus kata DB
                
                # Perbaiki teks yang terpotong
                desc = desc.replace("TRS F", "TRSF").replace("E-B ANKING", "E-BANKING").replace("SW ITCHING", "SWITCHING")
                desc = " ".join(desc.split()).strip() # Rapikan spasi

                debit, kredit, saldo = None, None, None
                
                if len(numbers) >= 2:
                    # Angka terakhir adalah Saldo, sebelumnya Nominal
                    nominal = to_float_indo_bca(numbers[-2])
                    saldo = to_float_indo_bca(numbers[-1])
                    if "DB" in full_text.upper(): kredit = nominal
                    else: debit = nominal
                elif len(numbers) == 1:
                    nominal = to_float_indo_bca(numbers[0])
                    if "DB" in full_text.upper(): kredit = nominal
                    else: debit = nominal

                transactions.append({
                    "Tanggal": f"{tgl_bulan}/{tahun}",
                    "Kode": "-",
                    "Keterangan": desc,
                    "Debit": debit,
                    "Kredit": kredit,
                    "Saldo": saldo
                })
    return transactions

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
            "Saldo": txn["Saldo"]
        })
    return pd.DataFrame(rows)

# UI
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
            txns = parse_bca_final(uploaded_file, password=password)
            df = build_df(txns, uploaded_file.name)
            
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
        
        try:
            final_df['Tanggal_DT'] = pd.to_datetime(final_df['Tanggal'], format='%d/%m/%Y', errors='coerce')
            final_df = final_df.sort_values(by=['Tanggal_DT'])
            final_df = final_df.drop(columns=['Tanggal_DT'])
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
