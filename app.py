# --- KETERANGAN & TTD ---
block_height = 35
ket_width = total_width * 0.4
ttd_width = total_width * 0.6

y_start = pdf.get_y()

# Box Keterangan
pdf.set_font("Arial", "", 9)
pdf.rect(pdf.l_margin, y_start, ket_width, block_height)
pdf.set_xy(pdf.l_margin + 2, y_start + 2)
pdf.cell(0, 6, "Keterangan", ln=1)
if first_desc:
    pdf.set_x(pdf.l_margin + 2)
    pdf.multi_cell(ket_width - 4, 6, str(first_desc))
# garis putus-putus
y_dashed = y_start + block_height - 5
pdf.dashed_line(pdf.l_margin + 2, y_dashed, pdf.l_margin + ket_width - 2, y_dashed, 1, 2)

# Filter pejabat valid (yang ada jabatan atau nama)
pejabat_valid = [p for p in pejabat_names if p.get("jabatan") or p.get("nama")]

if pejabat_valid:
    pdf.set_xy(pdf.l_margin + ket_width + 5, y_start)
    col_width = (ttd_width - 5) / len(pejabat_valid)

    # Baris jabatan
    pdf.set_font("Arial", "B", 9)
    for p in pejabat_valid:
        pdf.cell(col_width, 8, p.get("jabatan",""), border=1, align="C")
    pdf.ln()

    # Baris tanda tangan kosong (1 kotak tinggi)
    pdf.set_x(pdf.l_margin + ket_width + 5)
    for _ in pejabat_valid:
        pdf.cell(col_width, 25, "", border=1, align="C")
    pdf.ln()

    # Baris nama pejabat
    pdf.set_x(pdf.l_margin + ket_width + 5)
    pdf.set_font("Arial", "", 9)
    for p in pejabat_valid:
        pdf.cell(col_width, 8, p.get("nama",""), border=1, align="C")
    pdf.ln(15)
