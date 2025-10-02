    # --- keterangan + tanda tangan ---
    pdf.ln(5)
    total_width = pdf.w - pdf.l_margin - pdf.r_margin
    half_width = total_width * 0.45  # untuk keterangan
    ttd_width = total_width - half_width

    # keterangan
    pdf.set_font("Arial","",9)
    pdf.multi_cell(half_width, 6, f"Keterangan:\n{first_desc}", border=1, align="L")
    y_after_ket = pdf.get_y()

    # tanda tangan
    pdf.set_xy(pdf.l_margin + half_width, y_after_ket - 12)
    ttd_labels = [lbl for lbl in [settings.get("ttd1",""), settings.get("ttd2",""), settings.get("ttd3","")] if lbl]
    ttd_names = [nm for nm in [settings.get("nama1",""), settings.get("nama2",""), settings.get("nama3","")]]

    if ttd_labels:
        col_width = ttd_width / len(ttd_labels)

        # header
        pdf.set_font("Arial","B",9)
        for lbl in ttd_labels:
            pdf.cell(col_width, 8, lbl, border=1, align="C")
        pdf.ln()

        # kotak tanda tangan
        pdf.set_x(pdf.l_margin + half_width)
        for _ in ttd_labels:
            pdf.cell(col_width, 20, "", border=1, align="C")
        pdf.ln()

        # nama pejabat
        pdf.set_x(pdf.l_margin + half_width)
        pdf.set_font("Arial","",9)
        for nm in ttd_names[:len(ttd_labels)]:
            pdf.cell(col_width, 8, nm, border=1, align="C")
        pdf.ln(10)

    # garis putus-putus di bawah keterangan
    pdf.set_dash_pattern(2, 2)
    pdf.line(pdf.l_margin, y_after_ket+15, pdf.l_margin+half_width, y_after_ket+15)
    pdf.set_dash_pattern()  # reset solid
