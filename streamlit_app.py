# =========================================================
# PARSER BCA - SALDO MENGIKUTI PDF
# =========================================================
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

    # 1. SALDO AWAL
    saldo_awal = 0.0
    m_saldo = re.search(r"SALDO\s+AWAL\s*[\n:]?\s*([\d,.]+)", full_text, re.IGNORECASE)
    if m_saldo:
        saldo_awal = to_float(m_saldo.group(1)) or 0.0

    # 2. PAJAK BUNGA & BUNGA (pemetaan Anda: DB->Kredit, CR->Debit)
    for match in re.finditer(r"PAJAK\s+BUNGA\s+([\d,.]+)", full_text, re.IGNORECASE):
        pajak_val = to_float(match.group(1))
        if pajak_val:
            ctx = full_text[max(0, match.start() - 500):match.start()]
            dates = re.findall(r"\b(\d{2}/\d{2})\b", ctx)
            tgl = dates[-1] if dates else None
            if tgl:
                transactions.append({
                    "Tanggal": f"{tgl}/{tahun}", "Kode": "-",
                    "Keterangan": "PAJAK BUNGA", "Debit": 0.0, "Kredit": pajak_val,
                    "Saldo": None, "Tipe Asli": "DB",
                })
    for match in re.finditer(r"(?<!PAJAK\s)BUNGA\s+([\d,.]+)", full_text, re.IGNORECASE):
        if "PAJAK" in full_text[max(0, match.start()-20):match.start()].upper():
            continue
        bunga_val = to_float(match.group(1))
        if bunga_val:
            ctx = full_text[max(0, match.start() - 500):match.start()]
            dates = re.findall(r"\b(\d{2}/\d{2})\b", ctx)
            tgl = dates[-1] if dates else None
            if tgl:
                transactions.append({
                    "Tanggal": f"{tgl}/{tahun}", "Kode": "-",
                    "Keterangan": "BUNGA", "Debit": bunga_val, "Kredit": 0.0,
                    "Saldo": None, "Tipe Asli": "CR",
                })

    # 3. PARSE BARIS PER BARIS
    lines = full_text.split('\n')
    ignore_keywords = [
        "REKENING GIRO", "KCU", "DOMINIC DIGITAL", "KELAPA GADING", "JL BOULEVARD",
        "JAKARTA", "INDONESIA", "NO. REKENING", "HALAMAN", "PERIODE", "MATA UANG",
        "CATATAN:", "Apabila nasabah", "BCA berhak", "TANGGAL KETERANGAN",
        "Bersambung ke halaman", "MUTASI CR :", "MUTASI DB :", "SALDO AKHIR :",
        "SALDO AWAL", "===== Page", "# REKENING", "Page ", "- Page ",
    ]

    clean_lines = []
    for line in lines:
        s = line.strip()
        if not s: continue
        if re.match(r"^\d+\s*/\s*\d+$", s): continue
        if re.match(r"^\d+$", s) and len(s) <= 3: continue
        if any(kw.lower() in s.lower() for kw in ignore_keywords): continue
        if re.search(r"BUNGA", s, re.IGNORECASE): continue
        clean_lines.append(s)

    raw_blocks = []
    current_block = None
    date_pattern = re.compile(r"^(\d{2}/\d{2})\s*(.+)")

    for s in clean_lines:
        m_date = date_pattern.match(s)
        if m_date:
            if current_block: raw_blocks.append(current_block)
            current_block = {"tgl": m_date.group(1), "lines": [m_date.group(2)]}
        elif current_block:
            current_block["lines"].append(s)
    if current_block: raw_blocks.append(current_block)

    db_keywords = [
        "EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET", "PENERIMAAN NEGARA",
        "TIGA MUSTIKA KONSU", "PT.FINNET", "PT. FINNET", "BA JASA E-BANKING",
        "BIAYA TRANSFER", "BIAYA ADM", "BIAYA REFERENSI", "SWITCHING KE",
        "SWITCHING TRANSFER KE", "BI-FAST BIF TRANSFER KE",
        "BI-FAST BIF BIAYA TXN KE", "OTOMATIS", "FEE TELEX", "FEE FULL AMT",
        "FEE VALUE TDY",
    ]

    for block in raw_blocks:
        combined = " ".join(block["lines"])
        upper = combined.upper()
        nums = re.findall(r"\d[\d,]*\.\d{2}", combined)
        if not nums: continue
        num_vals = [to_float(n) for n in nums if to_float(n) is not None]
        if not num_vals: continue

        # BIAYA ADM / TRANSFER / REFERENSI
        if any(k in upper for k in ["BIAYA ADM", "BIAYA TRANSFER", "BIAYA REFERENSI"]):
            biaya_val = num_vals[0]
            # Saldo dari PDF: angka terakhir jika beda & wajar
            saldo_val = None
            if len(num_vals) >= 2 and num_vals[-1] != biaya_val and num_vals[-1] > 1000:
                saldo_val = num_vals[-1]
            transactions.append({
                "Tanggal": f"{block['tgl']}/{tahun}", "Kode": "-",
                "Keterangan": "BIAYA ADM", "Debit": 0.0, "Kredit": biaya_val,
                "Saldo": saldo_val, "Tipe Asli": "DB",
            })
            continue

        # TENTUKAN DB / CR
        is_db = False
        if re.search(r"\bDB\b", upper): is_db = True
        elif any(k in upper for k in db_keywords): is_db = True
        elif re.search(r"\bCR\b", upper): is_db = False

        nominal = num_vals[0]

        # --- KUNCI: ambil SALDO dari PDF ---
        # PDF BCA menaruh saldo berjalan sebagai angka terakhir pada baris transaksi.
        # Ambil kandidat saldo jika berbeda dari nominal & masuk akal.
        saldo_pdf = None
        if len(num_vals) >= 2:
            cand = num_vals[-1]
            if cand != nominal and cand > 1000:
                saldo_pdf = cand

        desc = combined
        for n in nums: desc = desc.replace(n, "")
        desc = re.sub(r"\bDB\b|\bCR\b", "", desc, flags=re.I)
        desc = clean_desc(desc)

        # Pemetaan Anda: DB -> Kredit, CR -> Debit
        debit_val, kredit_val = (0.0, nominal) if is_db else (nominal, 0.0)

        transactions.append({
            "Tanggal": f"{block['tgl']}/{tahun}", "Kode": "-", "Keterangan": desc,
            "Debit": debit_val, "Kredit": kredit_val,
            "Saldo": saldo_pdf, "Tipe Asli": "DB" if is_db else "CR",
        })

    # 4. RUNNING BALANCE (fallback)
    # Urutkan berdasarkan tanggal + urutan asli (stabil) agar kronologis.
    for i, txn in enumerate(transactions):
        txn["_order"] = i
    transactions.sort(key=lambda x: (x["Tanggal"], x["_order"]))

    saldo_berjalan = saldo_awal
    for txn in transactions:
        if txn["Saldo"] is None:
            # Pemetaan Anda: Debit = uang masuk, Kredit = uang keluar
            saldo_berjalan = saldo_berjalan + txn["Debit"] - txn["Kredit"]
            txn["Saldo"] = saldo_berjalan
        else:
            # Ikuti saldo PDF apa adanya
            saldo_berjalan = txn["Saldo"]

    for txn in transactions:
        txn.pop("_order", None)

    return transactions
