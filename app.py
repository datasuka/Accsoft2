import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
from num2words import num2words
import zipfile
import calendar

# --- bersihkan data ---
def bersihkan_jurnal(df):
    df = df.rename(columns=lambda x: str(x).strip().lower())
    mapping = {
        "tanggal": "Tanggal",
        "nomor voucher jurnal": "Nomor Voucher Jurnal",
        "no akun": "No Akun",
        "akun": "Akun",
        "deskripsi": "Deskripsi",
        "debet": "Debet",
        "kredit": "Kredit",
        "departemen": "Departemen",
        "proyek": "Proyek",
        "subjek": "Subjek"
    }
    df = df.rename(columns={k.lower(): v for k,v in mapping.items() if k.lower() in df.columns})
    for col in ["Debet","Kredit"]:
        df[col] = pd.to_numeric(df.get(col,0), errors="coerce").fillna(0)
    return df

# format angka Indonesia
def fmt_num(val):
    try:
        return "{:,.0f}".format(float(val)).replace(",", ".")
    except:
        return "0"

# --- generate voucher ---
def buat_voucher(df, no_voucher, settings):
    pdf = FPDF("P", "mm", "A4")
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.add_page()

    # Header kiri (logo + perusahaan + alamat)
    if settings.get("logo"):
        logo_y = 10
        pdf.image(settings["logo"], 15, logo_y, settings.get("logo_size", 20))

    pdf.set_font("Arial", "B", 12)
    pdf.set_xy(40, 10)
    pdf.cell(70, 6, settings.get("perusahaan",""), ln=1)

    pdf.set_font("Arial", "", 9)
    pdf.set_x(40)
    pdf.multi_cell(70, 5, settings.get("alamat",""), align="L")

    # Header kanan: judul + info kotak
    judul = settings.get("judul_doc", "Jurnal Voucher")
    header_width = 90
    header_x = pdf.w - pdf.r_margin - header_width

    # Judul
    pdf.set_xy(header_x, 10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(header_width, 6, "", ln=1, align="R")
    pdf.set_x(header_x)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(header_width, 8, judul, ln=1, align="C")
    pdf.set_line_width(0.8)
    pdf.line(header_x, 10, pdf.w - pdf.r_margin, 10)
    pdf.line(header_x, 18, pdf.w - pdf.r_margin, 18)
    pdf.set_line_width(0.2)

    data = df[df["Nomor Voucher Jurnal"] == no_voucher]
    try:
        tgl = pd.to_datetime(data.iloc[0]["Tanggal"]).strftime("%d %b %Y")
    except:
        tgl = str(data.iloc[0]["Tanggal"])
    subjek = str(data.iloc[0].get("Subjek",""))

    pdf.set_xy(header_x, 22)
    pdf.set_font("Arial", "", 10)
    pdf.cell(35, 6, "Nomor Voucher", align="L")
    pdf.cell(header_width-35, 6, f": {no_voucher}", ln=1, align="L")

    pdf.set_x(header_x)
    pdf.cell(35, 6, "Tanggal", align="L")
    pdf.cell(header_width-35, 6, f": {tgl}", ln=1, align="L")

    pdf.set_x(header_x)
    label_subj = settings.get("label_subjek", "Subjek")
    pdf.cell(35, 6, label_subj, align="L")
    pdf.cell(header_width-35, 6, f": {subjek}", ln=1, align="L")

    pdf.ln(5)

    # tabel utama
    total_width = pdf.w - pdf.l_margin - pdf.r_margin
    base_col_widths = [25, 60, 50, 30, 30]
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

        # wrap rows
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

    # Terbilang row (tabel sendiri)
    terbilang = num2words(total_debit, lang='id')
    terbilang = " ".join([w.capitalize() for w in terbilang.split()])
    pdf.set_font("Arial","I",9)
    pdf.cell(25, 8, "Terbilang", border=1)
    pdf.cell(total_width-25, 8, terbilang + " Rupiah", border=1)
    pdf.ln()

    # Keterangan dengan garis putus
    if first_desc:
        pdf.set_font("Arial","",9)
        pdf.cell(total_width/2,8,"Keterangan",align="L")
        pdf.ln(5)
        pdf.multi_cell(total_width/2,6,first_desc)
        pdf.set_dash_pattern(1,1)
        pdf.line(pdf.l_margin, pdf.get_y(), pdf.l_margin+total_width/2, pdf.get_y())
        pdf.set_dash_pattern()

    pdf.ln(5)

    # tanda tangan
    pejabat = [p for p in settings.get("pejabat",[]) if p.strip()!=""]
    if not pejabat:
        pejabat = ["Finance","Disetujui","Penerima"]
    col_width = (pdf.w - pdf.l_margin - pdf.r_margin)/len(pejabat)
    pdf.set_font("Arial","B",10)
    for lbl in pejabat:
        pdf.cell(col_width,8,lbl,border=1,align="C")
    pdf.ln()
    pdf.set_font("Arial","",9)
    for _ in pejabat:
        pdf.cell(col_width,15,"",border=1)
    pdf.ln()
    for _ in pejabat:
        pdf.cell(col_width,8,"",border=1,align="C")
    pdf.ln(10)

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer

# --- Streamlit ---
st.set_page_config(page_title="Mini Akunting", layout="wide")
st.title("📑 Mini Akunting - Voucher")

# Sidebar
st.sidebar.header("⚙️ Pengaturan Perusahaan")
settings = {}
settings["perusahaan"] = st.sidebar.text_input("Nama Perusahaan")
settings["alamat"] = st.sidebar.text_area("Alamat Perusahaan")
settings["logo_size"] = st.sidebar.slider("Ukuran Logo (mm)", 10, 50, 20)
logo_file = st.sidebar.file_uploader("Upload Logo (PNG/JPG)", type=["png","jpg","jpeg"])
if logo_file:
    tmp = BytesIO(logo_file.read())
    settings["logo"] = tmp

settings["judul_doc"] = st.sidebar.text_input("Judul Dokumen", "Jurnal Voucher")
settings["label_subjek"] = st.sidebar.text_input("Label setelah Tanggal", "Subjek")
pejabat_cols = []
for i in range(1,4):
    pejabat_cols.append(st.sidebar.text_input(f"Kolom Tanda Tangan {i}", ""))
settings["pejabat"] = pejabat_cols

# Main content
file = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx","xls"])
if file:
    df = pd.read_excel(file)
    df = bersihkan_jurnal(df)
    st.dataframe(df.head())

    no_voucher = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())
    if st.button("Cetak"):
        pdf_file = buat_voucher(df, no_voucher, settings)
        st.download_button("⬇️ Download PDF", data=pdf_file, file_name=f"{no_voucher}.pdf")
