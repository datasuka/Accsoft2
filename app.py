    # --- TABEL UTAMA ---
    total_width = pdf.w - pdf.l_margin - pdf.r_margin
    base_col_widths = [25, 55, 55, 30, 30]
    scale = total_width / sum(base_col_widths)
    col_widths = [w*scale for w in base_col_widths]
    headers = ["Akun Perkiraan","Nama Akun","Memo","Debit","Kredit"]

    pdf.set_font("Arial","B",9)
    for h,w in zip(headers,col_widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln()

    total_debit, total_kredit = 0,0
    pdf.set_font("Arial","",9)
    first_desc = ""

    for _, row in data.iterrows():
        debit_val = row["Debet"]
        kredit_val = row["Kredit"]

        memo_text = ""
        if row.get("Departemen"):
            memo_text += f"- Departemen : {row['Departemen']}\n"
        if row.get("Proyek"):
            memo_text += f"- Proyek     : {row['Proyek']}"

        values = [
            str(row.get("No Akun","")),     # Akun Perkiraan
            str(row.get("Akun","")),        # Nama Akun
            memo_text.strip(),              # Memo
            fmt_num(debit_val),             # Debit
            fmt_num(kredit_val)             # Kredit
        ]

        # Hitung tinggi baris (wrap)
        line_counts = []
        for i2, (val, w) in enumerate(zip(values, col_widths)):
            if i2 in [3,4]:
                line_counts.append(1)
            else:
                lines = pdf.multi_cell(w, 6, val, split_only=True)
                line_counts.append(len(lines))
        max_lines = max(line_counts)
        row_height = max_lines * 6

        # gambar cell per kolom
        x_start = pdf.get_x()
        y_start = pdf.get_y()
        for i2, (val, w) in enumerate(zip(values, col_widths)):
            pdf.rect(x_start, y_start, w, row_height)
            pdf.set_xy(x_start, y_start)
            if i2 in [3,4]:
                pdf.cell(w, row_height, val, align="R")
            else:
                pdf.multi_cell(w, 6, val, align="L")
            x_start += w
        pdf.set_y(y_start + row_height)

        total_debit += debit_val
        total_kredit += kredit_val
        if first_desc == "" and str(row.get("Deskripsi","")).strip():
            first_desc = str(row["Deskripsi"]).strip()

    # total row
    pdf.set_font("Arial","B",9)
    pdf.cell(sum(col_widths[:-2]),8,"Total",border=1,align="R")
    pdf.cell(col_widths[3],8,fmt_num(total_debit),border=1,align="R")
    pdf.cell(col_widths[4],8,fmt_num(total_kredit),border=1,align="R")
    pdf.ln(12)

    # --- TERBILANG ---
    angka = int(total_debit)   # pastikan integer, biar ga ada koma nol
    terbilang = num2words(angka, lang='id')
    terbilang = " ".join([w.capitalize() for w in terbilang.split()])
    pdf.set_font("Arial", "I", 9)
    pdf.cell(total_width, 8, f"Terbilang : {terbilang} Rupiah", border=1, align="L")
    pdf.ln(12)
