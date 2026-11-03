# =========================================================
# PARSER BCA — FOKUS: PAJAK BUNGA & SALDO AKHIR
# =========================================================
def parse_bca_hybrid(file_obj, password=None):
    transactions = []
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
                # --- LOGIKA KHUSUS ---
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

    # --- FALLBACK: TEKS MENTAH ---
    if table_rows_found == 0:
        lines = full_text.split('\n')
        ignore_keywords = [
            "REKENING GIRO", "KCU", "DOMINIC DIGITAL", "KELAPA GADING", "JL BOULEVARD",
            "JAKARTA", "INDONESIA", "NO. REKENING", "HALAMAN", "PERIODE", "MATA UANG",
            "CATATAN:", "Apabila nasabah", "BCA berhak", "TANGGAL KETERANGAN", 
            "Bersambung ke halaman", "MUTASI CR :", "MUTASI DB :", "SALDO AKHIR :",
            "===== Page", "# REKENING", "Page ", "- Page "
        ]
        # TAMBAHAN PENTING: keyword yang menandakan BUNGA / PAJAK BUNGA / BIAYA ADM
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
                # Pisahkan BUNGA / PAJAK BUNGA / BIAYA ADM sebagai blok baru
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
            
            # --- LOGIKA KHUSUS ---
            is_db = False
            if "PAJAK BUNGA" in upper_text:
                is_db = True
            elif "BUNGA" in upper_text and "PAJAK" not in upper_text:
                is_db = False
            elif "BIAYA ADM" in upper_text or "BIAYA TRANSFER" in upper_text or "BIAYA REFERENSI" in upper_text:
                is_db = True
            elif re.search(r"\bDB\b", upper_text):
                is_db = True
            elif any(k in upper_text for k in ["EDDY SUFENDI", "TRANSFER VIA LLG", "AUTODEBET", "PENERIMAAN NEGARA"]):
                is_db = True

            all_decimals = re.findall(r"[\d,]+\.\d{2}", combined_text)
            if not all_decimals: continue

            # Untuk BUNGA/PAJAK BUNGA/BIAYA ADM: ambil angka pertama
            # Untuk transaksi lain: juga ambil angka pertama
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
