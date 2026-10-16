import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Maybank PDF to Excel", layout="wide")
st.title("🏦 Maybank PDF → Excel Converter")
st.write("Upload **satu atau lebih** file PDF RK Maybank, masukkan password, lalu unduh hasilnya.")

AMOUNT_RE = re.compile(r"(?<![\w/])(\d{1,3}(?:\.\d{3})*,\d{2})(?![\w/])")

# Regex untuk ambil saldo awal dari header PDF
SALDO_AWAL_RE = re.compile(r"Saldo\s+Awal\s*:?\s*([\d.]+,\d{2})", re.IGNORECASE)


def extract_saldo_awal(text):
    """Cari saldo awal di header PDF."""
    match = SALDO_AWAL_RE.search(text)
    if match:
        try:
            return float(match.group(1).replace(".", "").replace(",", "."))
        except:
            return None
    return None


def parse_maybank_text(text):
    lines = text.split("\n")
    transactions = []
    date_pattern = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(\d{4})\s+(.*)")
    current_txn = None
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        date_match = date_pattern.match(line)
        
        if date_match:
            if current_txn:
                transactions.append(current_txn)
            
            tgl = date_match.group(1)
            kode = date_match.group(2)
            sisa = date_match.group(3)
            
            amounts = AMOUNT_RE.findall(sisa)
            desc = AMOUNT_RE.sub(" ", sisa).strip()
            
            current_txn = {
                "Tanggal": tgl,
                "Kode": kode,
                "Keterangan": desc,
                "Nominal": amounts,
            }
        else:
            if current_txn:
                has_valid_amount = bool(AMOUNT_RE.search(line))
                amounts = AMOUNT_RE.findall(line) if has_valid_amount else []
                desc = AMOUNT_RE.sub(" ", line).strip() if has_valid_amount else line
                
                if desc:
                    current_txn["Keterangan"] += " " + desc
                
                if len(current_txn["Nominal"]) < 2 and amounts:
                    current_txn["Nominal"].extend(amounts)
    
    if current_txn:
        transactions.append(current_txn)
    
    return transactions


def to_float(s):
    if s is None or s == "":
        return None
    try:
        return float(s.replace(".", "").replace(",", "."))
    except:
        return None


def build_dataframe(transactions, sumber_file="", saldo_awal=None):
    rows = []
    prev_saldo = saldo_awal  # Mulai dari saldo awal PDF
    
    for txn in transactions:
        amounts = txn["Nominal"]
        saldo = None
        nominal = None
        
        # Ambil nominal & saldo dari 2 angka pertama
        if len(amounts) >= 2:
            nominal = amounts[0]
            saldo = amounts[1]
        elif len(amounts) == 1:
            # Cuma ada 1 angka — mungkin cuma saldo (transaksi tanpa nominal jelas)
            saldo = amounts[0]
        
        saldo_f = to_float(saldo)
        nominal_f = to_float(nominal)
        
        if saldo_f is None:
            saldo_f = prev_saldo
        
        # Tentukan Debit/Kredit dari selisih saldo
        debit = None
        kredit = None
        if saldo_f is not None and prev_saldo is not None and nominal_f is not None:
            if saldo_f > prev_saldo:
                debit = nominal_f     # saldo naik = uang masuk = Debit
            elif saldo_f < prev_saldo:
                kredit = nominal_f    # saldo turun = uang keluar = Kredit
        
        rows.append({
            "Sumber File": sumber_file,
            "Tanggal": txn["Tanggal"],
            "Kode": txn["Kode"],
            "Keterangan": txn["Keterangan"].strip(),
            "Debit": debit,
            "Kredit": kredit,
            "Saldo": saldo_f
        })
        
        if saldo_f is not None:
            prev_saldo = saldo_f
    
    return pd.DataFrame(rows)


# ---------- UI ----------
col1, col2 = st.columns(2)

with col1:
    password = st.text_input("🔑 Password PDF (jika ada)", type="password")

with col2:
    uploaded_files = st.file_uploader(
        "📄 Pilih file PDF Maybank (bisa lebih dari 1)",
        type="pdf",
        accept_multiple_files=True
    )

if uploaded_files:
    all_dfs = []
    progress = st.progress(0)
    status = st.empty()
    
    for idx, uploaded_file in enumerate(uploaded_files):
        status.info(f"Memproses: {uploaded_file.name} ({idx+1}/{len(uploaded_files)})")
        try:
            pdf_args = {"password": password} if password else {}
            
            all_text = ""
            with pdfplumber.open(uploaded_file, **pdf_args) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        all_text += text + "\n"
            
            # Ambil saldo awal dari header
            saldo_awal = extract_saldo_awal(all_text)
            
            with st.expander(f"🔍 Debug: {uploaded_file.name}"):
                st.write(f"**Saldo Awal terdeteksi:** {saldo_awal}")
                st.text(all_text[:3000])
            
            transactions = parse_maybank_text(all_text)
            df = build_dataframe(transactions, sumber_file=uploaded_file.name, saldo_awal=saldo_awal)
            
            if not df.empty:
                all_dfs.append(df)
                st.success(f"✅ {uploaded_file.name}: {len(df)} transaksi (saldo awal: {saldo_awal})")
            else:
                st.warning(f"⚠️ {uploaded_file.name}: tidak ada transaksi terdeteksi.")
                
        except Exception as e:
            st.error(f"❌ **{uploaded_file.name}**: {type(e).__name__}")
            st.exception(e)
        
        progress.progress((idx + 1) / len(uploaded_files))
    
    status.empty()
    progress.empty()
    
    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        st.success(f"✅ Total: {len(final_df)} transaksi dari {len(all_dfs)} file!")
        
        st.subheader("📊 Preview Data")
        st.dataframe(final_df, width='stretch')
        
        st.subheader("📁 Ringkasan per File")
        ringkasan = final_df.groupby("Sumber File").agg(
            Jumlah_Transaksi=("Tanggal", "count"),
            Total_Debit=("Debit", "sum"),
            Total_Kredit=("Kredit", "sum")
        ).reset_index()
        st.dataframe(ringkasan, width='stretch')
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='RK Maybank')
            ringkasan.to_excel(writer, index=False, sheet_name='Ringkasan')
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 Download sebagai Excel",
            data=processed_data,
            file_name="rk_maybank_gabungan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Tidak ada data yang berhasil diekstrak.")
else:
    st.info("Silakan upload minimal 1 file PDF Maybank untuk mulai.")