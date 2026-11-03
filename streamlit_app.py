@st.cache_data(show_spinner=False)
def parse_bca(file_bytes, password=None):
    transactions = []
    file_obj = BytesIO(file_bytes)
    pdf = open_pdf(file_obj, password)
    
    full_text = ""
    for page in pdf.pages:
        t = page.extract_text()
        if t: full_text += t + "\n"
    pdf.close()

    tahun = "2026"
    m = re.search(r"PERIODE\s*:\s*[A-Za-z]+\s+(\d{4})", full_text, re.IGNORECASE)
    if m: tahun = m.group(1)

    # --- DETEKSI BUNGA & PAJAK BUNGA GLOBAL ---
    # Cari pola "BUNGA <num> ... PAJAK BUNGA <num>" atau hanya "BUNGA <num>"
    # Kita proses semua kemungkinan.
    
    # 1. PAJAK BUNGA
    for match in re.finditer(
        r"PAJAK\s+BUNGA\s+([\d,.]+)", full_text, re.IGNORECASE
    ):
        pajak_val = to_float(match.group(1))
        if pajak_val:
            # Cari tanggal terdekat sebelum match
            start_pos = match.start()
            context_before = full_text[max(0, start_pos - 500):start_pos]
            date_matches = re.findall(r"\b(\d{2}/\d{2})\b", context_before)
            tgl_pajak = date_matches[-1] if date_matches else None
            if tgl_pajak:
                transactions.append({
                    "Tanggal": f"{tgl_pajak}/{tahun}", "Kode": "-",
                    "Keterangan": "PAJAK BUNGA", "Debit": 0.0, "Kredit": pajak_val,
                    "Saldo": None, "Tipe Asli": "DB",
                })
    
    # 2. BUNGA (hanya jika BUKAN bagian dari PAJAK BUNGA)
    for match in re.finditer(
        r"(?<!PAJAK\s)BUNGA\s+([\d,.]+)", full_text, re.IGNORECASE
    ):
        # Cek apakah sebelum kata BUNGA ada kata PAJAK
        start_check = max(0, match.start() - 20)
        prefix = full_text[start_check:match.start()].upper()
        if "PAJAK" in prefix:
            continue  # skip karena ini bagian dari PAJAK BUNGA
        bunga_val = to_float(match.group(1))
        if bunga_val:
            start_pos = match.start()
            context_before = full_text[max(0, start_pos - 500):start_pos]
            date_matches = re.findall(r"\b(\d{2}/\d{2})\b", context_before)
            tgl_bunga = date_matches[-1] if date_matches else None
            if tgl_bunga:
                transactions.append({
                    "Tanggal": f"{tgl_bunga}/{tahun}", "Kode": "-",
                    "Keterangan": "BUNGA", "Debit": bunga_val, "Kredit": 0.0,
                    "Saldo": None, "Tipe Asli": "CR",
                })

    # --- PARSE BARIS PER BARIS (TRANSAKSI NORMAL + BIAYA ADM) ---
    lines = full_text.split('\n')
    
    ignore_keywords = [
        "REKENING GIRO", "KCU", "DOMINIC DIGITAL", "KELAPA GADING", "JL BOULEVARD",
        "JAKARTA", "INDONESIA", "NO. REKENING", "HALAMAN", "PERIODE", "MATA UANG",
        "CATATAN:", "Apabila nasabah", "BCA berhak", "TANGGAL KETERANGAN",
        "Bersambung ke halaman", "MUTASI CR :", "MUTASI DB :", "SALDO AKHIR :",
        "===== Page", "# REKENING", "Page ", "- Page "
    ]
    
    clean_lines = []
    for line in lines:
        line_str = line.strip()
        if not line_str: continue
        if re.match(r"^\d+\s*/\s*\d+$", line_str): continue
        if re.match(r"^\d+$", line_str) and len(line_str) <= 3: continue
        if any(kw.lower() in line_str.lower() for kw in ignore_keywords): continue
        # Skip baris yang mengandung BUNGA/PAJAK BUNGA, karena sudah ditangani regex global
        if re.search(r"BUNGA", line_str, re.IGNORECASE): continue
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
            current_block["lines"].append(line_str)

    if current_block: raw_blocks.append(current_block)

    for block in raw_blocks:
        combined = " ".join(block["lines"])
        if "SALDO AWAL" in combined.upper(): continue

        upper = combined.upper()
        nums = re.findall(r"\d[\d,]*\.\d{2}", combined)
        if not nums: continue
        
        num_vals = [to_float(n) for n in nums if to_float(n) is not None]

        # BIAYA ADM / TRANSFER / REFERENSI
        if "BIAYA ADM" in upper or "BIAYA TRANSFER" in upper or "BIAYA REFERENSI" in upper:
            biaya_val = num_vals[0] if num_vals else 0.0
            saldo_val = None
            if len(num_vals) >= 2:
                saldo_val = num_vals[-1]
                if saldo_val == biaya_val:
                    saldo_val = None
            transactions.append({
                "Tanggal": f"{block['tgl']}/{tahun}", "Kode": "-",
                "Keterangan": "BIAYA ADM", "Debit": 0.0, "Kredit": biaya_val,
                "Saldo": saldo_val, "Tipe Asli": "DB",
            })
            continue

        # TRANSAKSI NORMAL
        is_db = False
        if re.search(r"\bDB\b", upper): is_db = True
        elif any(k in upper for k in ["EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET", "PENERIMAAN NEGARA"]): is_db = True

        nominal = num_vals[0] if num_vals else 0.0
        saldo = None
        if len(num_vals) >= 2:
            saldo_candidate = num_vals[-1]
            if saldo_candidate > 100 and saldo_candidate != nominal:
                saldo = saldo_candidate

        desc = combined
        for n in nums: desc = desc.replace(n, "")
        desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
        desc = clean_desc(desc)

        debit_val, kredit_val = (0.0, nominal) if is_db else (nominal, 0.0)
        transactions.append({
            "Tanggal": f"{block['tgl']}/{tahun}", "Kode": "-", "Keterangan": desc,
            "Debit": debit_val, "Kredit": kredit_val,
            "Saldo": saldo, "Tipe Asli": "DB" if is_db else "CR",
        })

    return transactions
