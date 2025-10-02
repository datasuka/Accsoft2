def buat_voucher(df, no_voucher, settings, pejabat, tinggi_ttd=30):
    pdf = FPDF("P", "mm", "A4")
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.add_page()

    # Header kiri (logo + perusahaan + alamat)
    if settings.get("logo"):
        pdf.image(settings["logo"], 15, 10, settings.get("logo_size", 20))

    pdf.set_font("Arial", "B", 12)
    pdf.set_xy(40, 10)
    pdf.cell(100, 6, settings.get("perusahaan",""), ln=1)
    pdf.set_font("Arial", "", 9)
    pdf.set_x(40)
    pdf.multi_cell(60, 5, settings.get("alamat",""), align="L")

    # Header kanan (judul dalam garis)
    judul = settings.get("judul_dokumen", "Jurnal Voucher")
    header_x = 120
    pdf.set_xy(header_x, 10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(80, 6, judul, align="R")
    pdf.set_draw_color(0,0,0)
    pdf.set_line_width(0.6)
    pdf.line(header_x, 15, 200, 15)
    pdf.line(header_x, 22, 200, 22)

    # Info voucher
    data = df[df["Nomor Voucher Jurnal"] == no_voucher]
    try:
        tgl = pd.to_datetime(data.iloc[0]["Tanggal"]).strftime("%d %b %Y")
    except:
        tgl = str(data.iloc[0]["Tanggal"])
    subjek_val = str(data.iloc[0].get("Subjek",""))

    pdf.set_xy(header_x, 25)
    pdf.set_font("Arial", "", 10)
    pdf.cell(30, 6, "Nomor Voucher", align="L")
    pdf.cell(50, 6, f": {no_voucher}", ln=1)

    pdf.set_x(header_x)
    pdf.cell(30, 6, "Tanggal", align="L")
    pdf.cell(50, 6, f": {tgl}", ln=1)

    pdf.set_x(header_x)
    pdf.cell(30, 6, settings.get("label_subjek","Subjek"), align="L")
    pdf.cell(50, 6, f": {subjek_val}", ln=1)
    pdf.ln(5)

    # Tabel utama
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
            str(row.get("No Akun","")),
            str(row.get("Akun","")),
            memo_text.strip(),
            fmt_num(debit_val),
            fmt_num(kredit_val)
        ]

        # cek tinggi baris
        line_counts = []
        for i2, (val, w) in enumerate(zip(values, col_widths)):
            if i2 in [3,4]:
                line_counts.append(1)
            else:
                lines = pdf.multi_cell(w, 6, val, split_only=True)
                line_counts.append(len(lines))
        max_lines = max(line_counts)
        row_height = max_lines * 6

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
    pdf.ln()

    # --- TERBILANG tanpa koma nol ---
    angka_terbilang = int(round(total_debit))
    terbilang = num2words(angka_terbilang, lang='id')
    terbilang = " ".join([w.capitalize() for w in terbilang.split()])
    terbilang = terbilang.replace("Koma Nol", "").replace("koma nol", "").strip()

    pdf.set_font("Arial", "I", 9)
    pdf.cell(total_width, 8, f"Terbilang : {terbilang} Rupiah", border=1, align="L")
    pdf.ln(10)

    # --- KETERANGAN & TTD ---
    ket_width = total_width * 0.4
    ttd_width = total_width * 0.6

    y_start = pdf.get_y()

    # Keterangan
    pdf.set_font("Arial", "", 9)
    pdf.rect(pdf.l_margin, y_start, ket_width, tinggi_ttd)
    pdf.set_xy(pdf.l_margin + 2, y_start + 2)
    pdf.cell(0, 6, "Keterangan", ln=1)
    if first_desc:
        pdf.set_x(pdf.l_margin + 2)
        pdf.multi_cell(ket_width - 4, 6, str(first_desc))
    # garis putus-putus bawah
    y_dashed = y_start + tinggi_ttd - 5
    pdf.dashed_line(pdf.l_margin + 2, y_dashed, pdf.l_margin + ket_width - 2, y_dashed, 1, 2)

    # TTD box
    pdf.set_xy(pdf.l_margin + ket_width + 5, y_start)
    col_width = (ttd_width - 5) / len(pejabat)

    pdf.set_font("Arial", "B", 9)
    for jabatan, _ in pejabat:
        pdf.cell(col_width, 8, jabatan if jabatan else "", border=1, align="C")
    pdf.ln()

    pdf.set_x(pdf.l_margin + ket_width + 5)
    for _ in pejabat:
        pdf.cell(col_width, tinggi_ttd-16, "", border=1, align="C")
    pdf.ln()

    pdf.set_x(pdf.l_margin + ket_width + 5)
    pdf.set_font("Arial", "", 9)
    for _, nama in pejabat:
        pdf.cell(col_width, 8, nama if nama else "", border=1, align="C")
    pdf.ln(15)

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer
