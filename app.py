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
        "kredit": "Kredit"
    }
    df = df.rename(columns={k.lower(): v for k,v in mapping.items() if k.lower() in df.columns})
    df["Debet"] = pd.to_numeric(df.get("Debet", 0), errors="coerce").fillna(0)
    df["Kredit"] = pd.to_numeric(df.get("Kredit", 0), errors="coerce").fillna(0)
    return df

# --- generate voucher ---
def buat_voucher(df, no_voucher, settings, jenis_doc):
    pdf = FPDF("P", "mm", "A4")
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.add_page()

    # Header perusahaan (logo kiri, info kanan)
    if settings.get("logo"):
        pdf.image(settings["logo"], 15, 8, settings.get("logo_size", 20))

    pdf.set_font("Arial", "B", 12)
    pdf.set_xy(40, 10)
    pdf.multi_cell(80, 6, settings.get("perusahaan",""))
    pdf.set_font("Arial", "", 9)
    pdf.set_x(40)
    pdf.multi_cell(80, 5, settings.get("alamat",""))

    # header kanan
    pdf.set_xy(140, 10)
    pdf.set_font("Arial", "B", 12)
    if jenis_doc == "Jurnal Umum":
        pdf.cell(0, 6, "Jurnal Voucher", ln=1, align="R")
    elif jenis_doc == "Bukti Pengeluaran Kas/Bank":
        pdf.cell(0, 6, "Bukti Pengeluaran Kas/Bank", ln=1, align="R")
    else:
        pdf.cell(0, 6, "Bukti Penerimaan Kas/Bank", ln=1, align="R")

    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, f"Nomor Voucher : {no_voucher}", ln=1, align="R")

    # ambil data voucher
    data = df[df["Nomor Voucher Jurnal"] == no_voucher]
    try:
        tgl = pd.to_datetime(data.iloc[0]["Tanggal"]).strftime("%d %b %Y")
    except:
        tgl = str(data.iloc[0]["Tanggal"])
    pdf.cell(0, 6, f"Tanggal       : {tgl}", ln=1, align="R")

    if jenis_doc == "Bukti Pengeluaran Kas/Bank":
        pdf.cell(0, 6, "Penerima     :", ln=1, align="R")
    elif jenis_doc == "Bukti Penerimaan Kas/Bank":
        pdf.cell(0, 6, "Pemberi      :", ln=1, align="R")

    pdf.ln(3)

    # table header
    col_widths = [28, 60, 60, 25, 25]
    headers = ["Akun Perkiraan","Nama Akun","Memo","Debit","Kredit"]

    pdf.set_font("Arial","B",9)
    for h,w in zip(headers,col_widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln()

    total_debit, total_kredit = 0,0
    pdf.set_font("Arial","",9)

    first_desc = ""
    for i, row in data.iterrows():
        debit_val = int(row["Debet"]) if pd.notna(row["Debet"]) else 0
        kredit_val = int(row["Kredit"]) if pd.notna(row["Kredit"]) else 0

        values = [
            str(row["No Akun"]),
            str(row["Akun"]),
            str(row.get("Deskripsi","")),
            f"{debit_val:,}".replace(",", "."),
            f"{kredit_val:,}".replace(",", ".")
        ]

        # tinggi baris wrap
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
    pdf.cell(col_widths[3],8,f"{total_debit:,}".replace(",", "."),border=1,align="R")
    pdf.cell(col_widths[4],8,f"{total_kredit:,}".replace(",", "."),border=1,align="R")
    pdf.ln()

    # Terbilang row (kapital tiap kata)
    terbilang = num2words(total_debit, lang='id')
    terbilang = " ".join([w.capitalize() for w in terbilang.split()])
    pdf.set_font("Arial","I",9)
    pdf.cell(sum(col_widths),8,f"Terbilang : {terbilang} Rupiah",border=1,align="L")
    pdf.ln()

    # Keterangan (pakai deskripsi)
    if first_desc:
        pdf.set_font("Arial","",9)
        pdf.cell(sum(col_widths),8,f"Keterangan : {first_desc}",border=1,align="L")
        pdf.ln()

    # tanda tangan
    pdf.ln(10)
    pdf.set_font("Arial","",10)

    if jenis_doc == "Jurnal Umum":
        ttd_labels = ["Finance","Disetujui","Penerima"]
    elif jenis_doc == "Bukti Pengeluaran Kas/Bank":
        ttd_labels = ["Finance","Disetujui","Penerima"]
    else:
        ttd_labels = ["Finance","Disetujui","Pemberi"]

    col_width = (pdf.w - pdf.l_margin - pdf.r_margin) / len(ttd_labels)
    for label in ttd_labels:
        pdf.cell(col_width,6,label,align="C",border=1)
    pdf.ln(20)
    for label in ttd_labels:
        pdf.cell(col_width,6,"(................)",align="C",border=1)
    pdf.ln(10)

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
