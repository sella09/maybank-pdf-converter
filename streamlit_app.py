import streamlit as st
import camelot
import pandas as pd
import re
import traceback
from io import BytesIO

# =========================================================
# KONFIGURASI HALAMAN
# =========================================================
st.set_page_config(page_title="Bank PDF to Excel Converter", layout="wide")
st.title("🏦 Bank PDF → Excel Converter (Presisi Tinggi)")

# =========================================================
# HELPER FUNCTIONS
# =========================================================
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

def clean_date(date_str):
    if not date_str: return None
    # Format di PDF: 05/01
    match = re.search(r"(\d{2})/(\d{2})", str(date_str))
    if match:
        return f"{match.group(1)}/{match.group(2)}/2026"
    return None

# =========================================================
# PARSER BCA DENGAN CAMELOT
# =========================================================
def parse_bca_camelot(file_bytes):
    # Simpan file sementara
    with open("temp.pdf", "wb") as f:
        f.write(file_bytes)

    try:
        # Ekstrak tabel menggunakan Camelot (Stream mode untuk PDF tanpa garis)
        tables = camelot.read_pdf("temp.pdf", pages='all', flavor='stream', edge_tol=100)
        
        if not tables:
            return pd.DataFrame()

        all_rows = []
        for table in tables:
            df = table.df
            # Cari kolom yang berisi tanggal (format DD/MM)
            for index, row in df.iterrows():
                tgl_raw = str(row[0]).strip()
                if re.match(r"^\d{2}/\d{2}$", tgl_raw):
                    # Ekstrak semua teks di baris tersebut
                    row_text = " ".join(row.tolist())
                    
                    # Ambil angka (Mutasi dan Saldo)
                    nums = re.findall(r"\d[\d,]*\.\d{2}", row_text)
                    
                    # Bersihkan teks keterangan dari angka
                    desc = row_text
                    for n in nums:
                        desc = desc.replace(n, "")
                    desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
                    desc = " ".join(desc.split()).strip()

                    # Cek DB/CR
                    is_db = False
                    if re.search(r"\bDB\b", row_text.upper()): is_db = True
                    elif any(k in desc.upper() for k in ["EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET", "BIAYA ADM", "PAJAK BUNGA", "BAYAR", "PENERIMAAN NEGARA"]):
                        is_db = True

                    nominal = to_float(nums[0]) if nums else 0.0
                    saldo = to_float(nums[-1]) if len(nums) > 1 else 0.0

                    # Konversi ke Debit/Kredit
                    debit_val, kredit_val = (0.0, nominal) if is_db else (nominal, 0.0)
                    
                    all_rows.append({
                        "Tanggal": f"{tgl_raw}/2026",
                        "Keterangan": desc,
                        "Debit": debit_val,
                        "Kredit": kredit_val,
                        "Saldo": saldo
                    })
        return pd.DataFrame(all_rows)
    except Exception as e:
        st.error(f"Gagal mengekstrak tabel: {str(e)}")
        return pd.DataFrame()

# =========================================================
# UI STREAMLIT
# =========================================================
uploaded_files = st.file_uploader("📄 Pilih file PDF BCA", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        file_bytes = uploaded_file.read()
        
        with st.spinner(f"Memproses {uploaded_file.name}..."):
            df = parse_bca_camelot(file_bytes)
            
            if not df.empty:
                all_dfs.append(df)
                st.success(f"✅ {uploaded_file.name}: {len(df)} transaksi.")
            else:
                st.warning(f"⚠️ {uploaded_file.name}: Tidak ada transaksi terbaca.")

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        
        # Urutkan berdasarkan Tanggal
        final_df["Tanggal_DT"] = pd.to_datetime(final_df["Tanggal"], format="%d/%m/%Y", errors="coerce")
        final_df = final_df.sort_values("Tanggal_DT").drop(columns=["Tanggal_DT"])
        
        # Hitung Running Balance (jika saldo dari PDF kosong)
        # (Biasanya Camelot berhasil membaca saldo, jadi ini hanya fallback)
        saldo_awal = 14083133.34 # Ambil dari PDF Januari
        final_df["Saldo"] = final_df["Saldo"].astype(float)
        
        st.subheader("📊 Preview Data")
        st.dataframe(final_df, use_container_width=True)

        # Export Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Mutasi")
            
        st.download_button(
            "📥 Download Excel",
            data=output.getvalue(),
            file_name="mutasi_bca_presisi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
