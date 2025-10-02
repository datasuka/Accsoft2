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
        "subjek": "Subjek",
        "debet": "Debet",
        "kredit": "Kredit",
        "departemen": "Departemen",
        "proyek": "Proyek"
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
def buat_voucher(df, no_voucher, settings, jenis_doc):
    pdf = FPDF("P", "mm", "A4")
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.add_page()

    # Header kiri (logo + perusahaan + alamat)
    if settings.get("logo"):
        pdf.image(settings["logo"], 15, 8, settings.get("logo_size", 20))

    pdf.set_font("Arial", "B", 12)
    pdf.set_xy(40, 10)
    pdf.multi_cell(70, 6, settings.get("perusahaan",""))  # max width 70
    pdf.set_font("Arial", "", 9)
    pdf.set_x(40)
    pdf.multi_cell(70, 5, settings.get("alamat",""), align="L")  # wrap max width 70

    # Ambil data voucher
    data = df[df["Nomor Voucher Jurnal"] == no_voucher]
    try:
        tgl = pd.to_datetime(data.iloc[0]["Tanggal"]).strftime("%d %b %Y")
    except:
        tgl = str(data.iloc[0]["Tanggal"])

    subjek_val = str(data.iloc[0].get("Subjek",""))

    # Judul rata tengah
    pdf.set_font("Arial", "B", 14)
    pdf.set_xy(0, 10)
    pdf.cell(pdf.w - 20, 10, "Jurnal Voucher", ln=1, align="R")
    pdf.set_font("Arial", "", 10)

    # Info Nomor Voucher / Tanggal / Subjek
    header_x = pdf.w - pdf.r_margin - 90
    label_w = 30
    value_w = 60

    pdf.set_xy(header_x, 20)
    pdf.cell(label_w, 7, "Nomor Voucher :", border=0)
    pdf.multi_cell(value_w, 7, no_voucher, border=0)

    pdf.set_x(header_x)
    pdf.cell(label_w, 7, "Tanggal :", border=0)
    pdf.multi_cell(value_w, 7, tgl, border=0)

    pdf.set_x(header_x)
    pdf.cell(label_w, 7, "Subjek :", border=0)
    pdf.multi_cell(value_w, 7, subjek_val, border=0)

    pdf.ln(5)

    # tabel utama
    total_width = pdf.w - pdf.l_margin - pdf.r_margin
    base_col_widths = [25, 60, 60, 35, 35]
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
        if str(row.get("Deskripsi","")).strip():
            if memo_text:
                memo_text += "\n"
            memo_text += str(row["Deskripsi"])

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

    # Terbilang row (sendiri)
    terbilang = num2words(total_debit, lang='id')
    terbilang = " ".join([w.capitalize() for w in terbilang.split()])
    pdf.set_font("Arial","I",9)
    pdf.cell(sum(col_widths),8,f"Terbilang : {terbilang} Rupiah",border=1,align="L")
    pdf.ln()

    # Keterangan
    if first_desc:
        pdf.set_font("Arial","",9)
        pdf.cell(sum(col_widths),8,f"Keterangan : {first_desc}",border=1,align="L")
        pdf.ln()

    # tanda tangan
    pdf.ln(10)
    pdf.set_font("Arial","",10)
    if jenis_doc == "Bukti Penerimaan Kas/Bank":
        ttd_labels = ["Finance","Disetujui","Pemberi"]
    else:
        ttd_labels = ["Finance","Disetujui","Penerima"]

    col_width = (pdf.w - pdf.l_margin - pdf.r_margin) / len(ttd_labels)
    for lbl in ttd_labels:
        pdf.cell(col_width, 10, lbl, border=1, align="C")
    pdf.ln()
    for _ in ttd_labels:
        pdf.cell(col_width, 25, "", border=1, align="C")
    pdf.ln()

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer

# --- Streamlit ---
st.set_page_config(page_title="Mini Akunting", layout="wide")
st.title("📑 Mini Akunting - Voucher Jurnal / Kas / Bank")

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

# Jenis dokumen
jenis_doc = st.radio("Pilih Jenis Dokumen", ["Jurnal Umum","Bukti Pengeluaran Kas/Bank","Bukti Penerimaan Kas/Bank"])

# Main content
file = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx","xls"])
if file:
    df = pd.read_excel(file)
    df = bersihkan_jurnal(df)
    st.dataframe(df.head())

    mode = st.radio("Pilih Mode Cetak", ["Single Voucher", "Per Bulan"])

    if mode == "Single Voucher":
        no_voucher = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())
        if st.button("Cetak"):
            pdf_file = buat_voucher(df, no_voucher, settings, jenis_doc)
            st.download_button("⬇️ Download PDF", data=pdf_file, file_name=f"{no_voucher}.pdf")

    else:
        df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
        df = df.dropna(subset=["Tanggal"])
        bulan = st.selectbox("Pilih Bulan", range(1,13), format_func=lambda x: calendar.month_name[x])

        if st.button("Cetak Semua Voucher Bulan Ini"):
            buffer_zip = BytesIO()
            with zipfile.ZipFile(buffer_zip, "w") as zf:
                for v in df[df["Tanggal"].dt.month==bulan]["Nomor Voucher Jurnal"].unique():
                    pdf_file = buat_voucher(df, v, settings, jenis_doc)
                    zf.writestr(f"{v}.pdf", pdf_file.getvalue())
            buffer_zip.seek(0)
            st.download_button("⬇️ Download ZIP", data=buffer_zip, file_name=f"voucher_{bulan}.zip", mime="application/zip")
