import streamlit as st
from fpdf import FPDF
from datetime import datetime

class PDF(FPDF):
    def header(self):
        pass

    def footer(self):
        pass

def buat_voucher(data, settings):
    pdf = PDF()
    pdf.add_page()
    pdf.set_font("Arial", "", 11)

    # --- LOGO + HEADER ---
    if settings.get("logo_path"):
        pdf.image(settings["logo_path"], 10, 10, 25)

    pdf.set_xy(40, 10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 7, settings.get("company", "NAMA PERUSAHAAN"), ln=1)

    pdf.set_font("Arial", "", 9)
    pdf.multi_cell(70, 5, settings.get("alamat", ""))

    # --- JUDUL ---
    pdf.set_xy(140, 10)
    pdf.set_font("Arial", "B", 13)
    pdf.cell(60, 7, settings.get("judul_doc", "Jurnal Voucher"), border="TB", align="C", ln=1)

    # --- INFO KANAN ---
    pdf.set_font("Arial", "", 10)
    pdf.ln(5)
    label_w, colon_w, value_w = 35, 5, 50
    rows = [
        ("Nomor Voucher", data.get("nomor_voucher", "")),
        ("Tanggal", data.get("tanggal", "")),
        (settings.get("label_setelah_tanggal","Subjek"), data.get("subjek","")),
    ]
    for lbl, val in rows:
        pdf.set_x(130)
        pdf.cell(label_w, 6, lbl, align="L")
        pdf.cell(colon_w, 6, ":", align="C")
        pdf.multi_cell(value_w, 6, str(val), align="L")

    pdf.ln(4)

    # --- TABEL JURNAL ---
    col_widths = [30, 60, 40, 30, 30]
    headers = ["Akun Perkiraan", "Nama Akun", "Memo", "Debit", "Kredit"]

    pdf.set_font("Arial", "B", 10)
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 8, h, border=1, align="C")
    pdf.ln()

    pdf.set_font("Arial", "", 9)
    for row in data.get("jurnal", []):
        pdf.multi_cell(col_widths[0], 6, str(row[0]), border=1)  # wrap akun
        pdf.set_xy(pdf.get_x() + col_widths[0], pdf.get_y() - 6)
        pdf.cell(col_widths[1], 6, row[1], border=1)
        pdf.cell(col_widths[2], 6, row[2], border=1)
        pdf.cell(col_widths[3], 6, row[3], border=1, align="R")
        pdf.cell(col_widths[4], 6, row[4], border=1, align="R")
        pdf.ln()

    # total
    pdf.set_font("Arial", "B", 9)
    pdf.cell(sum(col_widths[:-2]), 7, "Total", border=1, align="R")
    pdf.cell(col_widths[-2], 7, data.get("total_debit",""), border=1, align="R")
    pdf.cell(col_widths[-1], 7, data.get("total_kredit",""), border=1, align="R")
    pdf.ln(10)

    # --- TERBILANG ---
    pdf.set_font("Arial", "I", 9)
    pdf.cell(25, 7, "Terbilang", border=1)
    pdf.multi_cell(0, 7, data.get("terbilang",""), border=1)
    pdf.ln(5)

    # --- KETERANGAN + TTD ---
    pdf.set_font("Arial", "", 9)
    pdf.cell(90, 6, "Keterangan", ln=0)
    pdf.cell(90, 6, "", ln=1)
    pdf.multi_cell(90, 6, data.get("keterangan",""), border="T")
    pdf.cell(90, 6, "-"*40, ln=0)  # garis putus2 dummy
    pdf.set_xy(110, pdf.get_y()-12)

    # tabel TTD
    jabatan_pejabat = settings.get("ttd_cols", [])
    if jabatan_pejabat:
        col_w = 180 / len(jabatan_pejabat)
        pdf.set_font("Arial", "B", 9)
        for jab, _ in jabatan_pejabat:
            pdf.cell(col_w, 7, jab, border=1, align="C")
        pdf.ln()
        pdf.set_font("Arial", "", 9)
        for _ in jabatan_pejabat:
            pdf.cell(col_w, 20, "", border=1, align="C")
        pdf.ln()
        for _, nm in jabatan_pejabat:
            pdf.cell(col_w, 7, nm, border=1, align="C")
        pdf.ln(10)

    return pdf

# --- STREAMLIT ---
st.sidebar.title("Pengaturan")
judul = st.sidebar.text_input("Judul Dokumen", "Jurnal Voucher")
label_subjek = st.sidebar.text_input("Label setelah Tanggal", "Subjek")
isi_subjek = st.sidebar.text_input("Isi Subjek/Pemberi/Penerima", "Putri")

jab1 = st.sidebar.text_input("Jabatan 1", "Finance")
nm1 = st.sidebar.text_input("Nama Pejabat 1", "")
jab2 = st.sidebar.text_input("Jabatan 2", "Disetujui")
nm2 = st.sidebar.text_input("Nama Pejabat 2", "")
jab3 = st.sidebar.text_input("Jabatan 3", "Penerima")
nm3 = st.sidebar.text_input("Nama Pejabat 3", "")

# contoh data
data = {
    "nomor_voucher": "Mandiri USD 6955/2024/001",
    "tanggal": "01 Jan 2024",
    "subjek": isi_subjek,
    "jurnal": [
        ["1004", "Bank Mandiri USD 122-00-0559695-5", "- Dept: BMA#02\n- Proyek: TRK", "10.000.000.000", "0"],
        ["1008", "Ayat Silang", "- Dept: BMA#02\n- Proyek: TRK", "0", "10.000.000.000"]
    ],
    "total_debit": "10.000.000.000",
    "total_kredit": "10.000.000.000",
    "terbilang": "Sepuluh Miliar Koma Nol Rupiah",
    "keterangan": "Sweep"
}

settings = {
    "judul_doc": judul,
    "label_setelah_tanggal": label_subjek,
    "subjek": isi_subjek,
    "ttd_cols": [(jab1, nm1), (jab2, nm2), (jab3, nm3)]
}

if st.button("Cetak"):
    pdf = buat_voucher(data, settings)
    st.download_button("Download PDF", data=pdf.output(dest="S").encode("latin-1"),
                       file_name="voucher.pdf", mime="application/pdf")
