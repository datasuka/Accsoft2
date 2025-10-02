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

# Format angka Indonesia
def fmt_num(val):
    try:
        return "{:,.0f}".format(float(val)).replace(",", ".")
    except:
        return "0"

# --- Generate Voucher ---
def buat_voucher(df, no_voucher, settings, pejabat_info, ttd_height=25):
    pdf = FPDF("P", "mm", "A4")
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.add_page()

    # Logo + Perusahaan
    if settings.get("logo"):
        pdf.image(settings["logo"], 15, 8, settings.get("logo_size", 20))

    pdf.set_font("Arial", "B", 12)
    pdf.set_xy(40, 10)
    pdf.cell(100, 6, settings.get("perusahaan",""), ln=1)
    pdf.set_font("Arial", "", 9)
    pdf.set_x(40)
    pdf.multi_cell(60, 5, settings.get("alamat",""), align="L")

    # Judul dokumen dengan garis atas–bawah
    judul = settings.get("judul_dokumen","Jurnal Voucher")
    pdf.set_font("Arial", "B", 12)
    pdf.set_xy(120, 10)
    pdf.line(120, 15, 200, 15)
    pdf.cell(80, 6, judul, align="C", ln=1)
    pdf.line(120, 21, 200, 21)

    # Data voucher
    data = df[df["Nomor Voucher Jurnal"] == no_voucher]
    tgl = pd.to_datetime(data.iloc[0]["Tanggal"], errors="coerce").strftime("%d %b %Y")
    subjek_val = str(data.iloc[0].get("Subjek",""))

    pdf.set_xy(120, 25)
    pdf.set_font("Arial", "", 10)
    pdf.cell(30, 6, "Nomor Voucher", align="L")
    pdf.cell(50, 6, f": {no_voucher}", ln=1)
    pdf.set_x(120); pdf.cell(30, 6, "Tanggal", align="L"); pdf.cell(50, 6, f": {tgl}", ln=1)
    pdf.set_x(120); pdf.cell(30, 6, settings.get("label_subjek","Subjek"), align="L"); pdf.cell(50, 6, f": {subjek_val}", ln=1)

    pdf.ln(5)

    # Tabel utama
    total_width = pdf.w - pdf.l_margin - pdf.r_margin
    col_widths = [25, 55, 55, 30, 30]
    headers = ["Akun Perkiraan","Nama Akun","Memo","Debit","Kredit"]

    pdf.set_font("Arial","B",9)
    for h,w in zip(headers,col_widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln()

    pdf.set_font("Arial","",9)
    total_debit, total_kredit, first_desc = 0,0,""

    for _, row in data.iterrows():
        debit_val, kredit_val = row["Debet"], row["Kredit"]
        memo_text = ""
        if row.get("Departemen"): memo_text += f"- Departemen : {row['Departemen']}\n"
        if row.get("Proyek"): memo_text += f"- Proyek     : {row['Proyek']}"

        values = [
            str(row.get("No Akun","")),
            str(row.get("Akun","")),
            memo_text,
            fmt_num(debit_val),
            fmt_num(kredit_val)
        ]

        # wrap isi cell
        heights = []
        for i, (val,w) in enumerate(zip(values,col_widths)):
            if i in [3,4]:
                heights.append(1)
            else:
                lines = pdf.multi_cell(w, 6, val, split_only=True)
                heights.append(len(lines))
        max_h = max(heights)*6
        x,y = pdf.get_x(), pdf.get_y()
        for i,(val,w) in enumerate(zip(values,col_widths)):
            pdf.rect(x,y,w,max_h)
            pdf.set_xy(x,y)
            if i in [3,4]: pdf.cell(w,max_h,val,align="R")
            else: pdf.multi_cell(w,6,val,align="L")
            x+=w
        pdf.set_y(y+max_h)

        total_debit += debit_val
        total_kredit += kredit_val
        if not first_desc and str(row.get("Deskripsi","")).strip():
            first_desc = str(row["Deskripsi"]).strip()

    # Total row
    pdf.set_font("Arial","B",9)
    pdf.cell(sum(col_widths[:-2]),8,"Total",border=1,align="R")
    pdf.cell(col_widths[3],8,fmt_num(total_debit),border=1,align="R")
    pdf.cell(col_widths[4],8,fmt_num(total_kredit),border=1,align="R")
    pdf.ln()

    # Terbilang
    pdf.set_font("Arial","I",9)
    angka = int(total_debit)
    terbilang = num2words(angka, lang="id")
    terbilang = " ".join([w.capitalize() for w in terbilang.split()])
    pdf.cell(total_width,8,f"Terbilang : {terbilang} Rupiah",border=1,ln=1)

    # Keterangan & TTD sejajar
    block_h = ttd_height
    ket_w, ttd_w = total_width*0.4, total_width*0.6
    y = pdf.get_y()

    # Keterangan
    pdf.rect(pdf.l_margin,y,ket_w,block_h)
    pdf.set_xy(pdf.l_margin+2,y+2)
    pdf.set_font("Arial","",9)
    pdf.cell(0,6,"Keterangan",ln=1)
    if first_desc:
        pdf.set_x(pdf.l_margin+2)
        pdf.multi_cell(ket_w-4,6,first_desc)
    pdf.dashed_line(pdf.l_margin+2,y+block_h-5,pdf.l_margin+ket_w-2,y+block_h-5,1,2)

    # TTD
    pdf.set_xy(pdf.l_margin+ket_w+5,y)
    col_w = (ttd_w-5)/max(1,len(pejabat_info))
    pdf.set_font("Arial","B",9)
    for jbt,_ in pejabat_info: pdf.cell(col_w,8,jbt or "",border=1,align="C")
    pdf.ln()
    pdf.set_x(pdf.l_margin+ket_w+5)
    for _ in pejabat_info: pdf.cell(col_w,block_h-16,"",border=1,align="C")
    pdf.ln()
    pdf.set_x(pdf.l_margin+ket_w+5)
    pdf.set_font("Arial","",9)
    for _,nm in pejabat_info: pdf.cell(col_w,8,nm or "",border=1,align="C")
    pdf.ln(15)

    buffer=BytesIO(); pdf.output(buffer); return buffer


# --- Streamlit App ---
st.set_page_config(page_title="Mini Akunting", layout="wide")
st.title("📑 Mini Akunting - Voucher Jurnal / Kas / Bank")

# Sidebar
st.sidebar.header("⚙️ Pengaturan Perusahaan")
settings = {}
settings["perusahaan"] = st.sidebar.text_input("Nama Perusahaan")
settings["alamat"] = st.sidebar.text_area("Alamat Perusahaan")
settings["logo_size"] = st.sidebar.slider("Ukuran Logo (mm)", 10, 50, 20)
settings["judul_dokumen"] = st.sidebar.text_input("Judul Dokumen", "Jurnal Voucher")
settings["label_subjek"] = st.sidebar.text_input("Label setelah Tanggal", "Subjek")

logo_file = st.sidebar.file_uploader("Upload Logo (PNG/JPG)", type=["png","jpg","jpeg"])
if logo_file:
    tmp = BytesIO(logo_file.read())
    settings["logo"] = tmp

st.sidebar.subheader("Kolom Tanda Tangan (max 3)")
pejabat_info = []
for i in range(1,4):
    jabatan = st.sidebar.text_input(f"Jabatan {i}", "")
    nama = st.sidebar.text_input(f"Nama Pejabat {i}", "")
    if jabatan or nama:
        pejabat_info.append((jabatan,nama))

ttd_height = st.sidebar.slider("Tinggi Kolom TTD (mm)", 20, 60, 35)

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
            pdf_file = buat_voucher(df, no_voucher, settings, pejabat_info, ttd_height)
            st.download_button("⬇️ Download PDF", data=pdf_file, file_name=f"{no_voucher}.pdf")

    else:
        df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
        df = df.dropna(subset=["Tanggal"])
        bulan = st.selectbox("Pilih Bulan", range(1,13), format_func=lambda x: calendar.month_name[x])

        if st.button("Cetak Semua Voucher Bulan Ini"):
            buffer_zip = BytesIO()
            with zipfile.ZipFile(buffer_zip, "w") as zf:
                for v in df[df["Tanggal"].dt.month==bulan]["Nomor Voucher Jurnal"].unique():
                    pdf_file = buat_voucher(df, v, settings, pejabat_info, ttd_height)
                    zf.writestr(f"{v}.pdf", pdf_file.getvalue())
            buffer_zip.seek(0)
            st.download_button("⬇️ Download ZIP", data=buffer_zip, file_name=f"voucher_{bulan}.zip", mime="application/zip")
