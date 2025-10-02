import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
from num2words import num2words
import zipfile
import calendar

# --- Bersihkan data ---
def bersihkan_jurnal(df):
    df = df.rename(columns=lambda x: str(x).strip().lower())
    mapping = {
        "tanggal": "Tanggal",
        "nomor voucher jurnal": "Nomor Voucher",
        "no akun": "No Akun",
        "akun": "Akun",
        "deskripsi": "Deskripsi",
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
def buat_voucher(df, no_voucher, settings):
    pdf = FPDF("P", "mm", "A4")
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.add_page()

    # Header kiri (logo + perusahaan + alamat)
    if settings.get("logo"):
        pdf.image(settings["logo"], 15, 8, settings.get("logo_size", 20))

    pdf.set_font("Arial", "B", 12)
    pdf.set_xy(40, 10)  # nama sejajar atas logo
    pdf.cell(60, 6, settings.get("perusahaan",""), ln=1)

    pdf.set_font("Arial", "", 9)
    pdf.set_x(40)
    pdf.multi_cell(60, 5, settings.get("alamat",""), align="L", max_line_height=5)

    # Header kanan (judul & info voucher)
    header_width = 100
    header_x = pdf.w - pdf.r_margin - header_width

    # Garis judul
    pdf.set_xy(header_x, 10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(header_width, 7, settings.get("judul_dokumen","Jurnal Voucher"), align="C", border="TB")
    pdf.ln(10)

    data = df[df["Nomor Voucher"] == no_voucher]
    try:
        tgl = pd.to_datetime(data.iloc[0]["Tanggal"]).strftime("%d %b %Y")
    except:
        tgl = str(data.iloc[0]["Tanggal"])

    # Info voucher
    info_items = [
        ("Nomor Voucher", no_voucher),
        ("Tanggal", tgl),
        (settings.get("label_tambahan","Subjek"), str(data.iloc[0].get("Subjek","")))
    ]
    pdf.set_font("Arial", "", 10)
    for label, val in info_items:
        pdf.set_x(header_x)
        pdf.cell(35, 6, f"{label}", align="L")
        pdf.cell(3, 6, ":", align="L")
        pdf.multi_cell(header_width-38, 6, val, align="L")

    pdf.ln(3)

    # Tabel utama
    total_width = pdf.w - pdf.l_margin - pdf.r_margin
    base_col_widths = [30, 55, 50, 30, 30]  
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
                lines = pdf.multi_cell(w, 5, val, split_only=True)
                line_counts.append(len(lines))
        max_lines = max(line_counts)
        row_height = max_lines * 5

        x_start = pdf.get_x()
        y_start = pdf.get_y()
        for i2, (val, w) in enumerate(zip(values, col_widths)):
            pdf.rect(x_start, y_start, w, row_height)
            pdf.set_xy(x_start, y_start)
            if i2 in [3,4]:
                pdf.cell(w, row_height, val, align="R")
            else:
                pdf.multi_cell(w, 5, val, align="L")
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
    pdf.ln(10)

    # Terbilang di tabel sendiri
    terbilang = num2words(total_debit, lang='id')
    terbilang = " ".join([w.capitalize() for w in terbilang.split()])
    pdf.set_font("Arial","I",9)
    pdf.cell(25,8,"Terbilang",border=1)
    pdf.cell(total_width-25,8,f"{terbilang} Rupiah",border=1)
    pdf.ln(10)

    # Keterangan + tanda tangan
    pdf.set_font("Arial","",9)
    ttd_labels = [settings.get("kolom1",""), settings.get("kolom2",""), settings.get("kolom3","")]
    pejabat_labels = [settings.get("pejabat1",""), settings.get("pejabat2",""), settings.get("pejabat3","")]

    left_w = total_width*0.4
    right_w = total_width-left_w
    pdf.cell(left_w,6,"Keterangan",border="T")
    pdf.cell(right_w,6,"",ln=1)

    pdf.multi_cell(left_w,6,first_desc, border="B")
    x_start = pdf.get_x()+left_w
    y_start = pdf.get_y()-6
    col_w = right_w/len(ttd_labels)
    pdf.set_xy(x_start,y_start)
    for lbl in ttd_labels:
        pdf.cell(col_w,6,lbl,border=1,align="C")
    pdf.ln()
    pdf.set_x(x_start)
    for _ in ttd_labels:
        pdf.cell(col_w,20,"",border=1,align="C")
    pdf.ln()
    pdf.set_x(x_start)
    for nm in pejabat_labels:
        pdf.cell(col_w,6,nm,border=1,align="C")

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

# Custom judul & label
settings["judul_dokumen"] = st.sidebar.text_input("Judul Dokumen","Jurnal Voucher")
settings["label_tambahan"] = st.sidebar.text_input("Label setelah Tanggal","Subjek")

# Kolom tanda tangan
settings["kolom1"] = st.sidebar.text_input("Jabatan 1","Finance")
settings["pejabat1"] = st.sidebar.text_input("Nama Pejabat 1","")
settings["kolom2"] = st.sidebar.text_input("Jabatan 2","Disetujui")
settings["pejabat2"] = st.sidebar.text_input("Nama Pejabat 2","")
settings["kolom3"] = st.sidebar.text_input("Jabatan 3","Penerima")
settings["pejabat3"] = st.sidebar.text_input("Nama Pejabat 3","")

# Main content
file = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx","xls"])
if file:
    df = pd.read_excel(file)
    df = bersihkan_jurnal(df)
    st.dataframe(df.head())

    mode = st.radio("Pilih Mode Cetak", ["Single Voucher", "Per Bulan"])

    if mode == "Single Voucher":
        no_voucher = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher"].unique())
        if st.button("Cetak"):
            pdf_file = buat_voucher(df, no_voucher, settings)
            st.download_button("⬇️ Download PDF", data=pdf_file, file_name=f"{no_voucher}.pdf")

    else:
        df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
        df = df.dropna(subset=["Tanggal"])
        bulan = st.selectbox("Pilih Bulan", range(1,13), format_func=lambda x: calendar.month_name[x])

        if st.button("Cetak Semua Voucher Bulan Ini"):
            buffer_zip = BytesIO()
            with zipfile.ZipFile(buffer_zip, "w") as zf:
                for v in df[df["Tanggal"].dt.month==bulan]["Nomor Voucher"].unique():
                    pdf_file = buat_voucher(df, v, settings)
                    zf.writestr(f"{v}.pdf", pdf_file.getvalue())
            buffer_zip.seek(0)
            st.download_button("⬇️ Download ZIP", data=buffer_zip, file_name=f"voucher_{bulan}.zip", mime="application/zip")
