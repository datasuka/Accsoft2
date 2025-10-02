import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
from num2words import num2words
import zipfile
import calendar

# --- Format Angka ---
def fmt_num(val):
    try:
        return "{:,.0f}".format(float(val)).replace(",", ".")
    except:
        return "0"

# --- PDF Generator ---
class PDF(FPDF):
    def header(self): pass
    def footer(self): pass

def buat_voucher(df, no_voucher, settings):
    pdf = PDF("P", "mm", "A4")
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.add_page()

    # ========== Header ==========
    if settings.get("logo"):
        pdf.image(settings["logo"], 15, 10, settings.get("logo_size", 20))

    # Nama perusahaan
    pdf.set_font("Arial", "B", 12)
    pdf.set_xy(40, 10)
    pdf.cell(60, 6, settings.get("perusahaan",""), ln=1)

    # Alamat (max width 60mm biar gak nabrak kanan)
    pdf.set_font("Arial", "", 9)
    pdf.set_x(40)
    pdf.multi_cell(60, 5, settings.get("alamat",""), align="L")

    # Blok kanan judul + info voucher
    judul = settings.get("judul_doc","Jurnal Voucher")
    header_x = 120
    pdf.set_xy(header_x, 10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(80, 7, judul, ln=1, align="C")
    pdf.line(header_x, 10, 200, 10)   # garis atas
    pdf.line(header_x, 17, 200, 17)   # garis bawah

    data = df[df["Nomor Voucher Jurnal"]==no_voucher]
    try:
        tgl = pd.to_datetime(data.iloc[0]["Tanggal"]).strftime("%d %b %Y")
    except:
        tgl = str(data.iloc[0]["Tanggal"])
    subjek_label = settings.get("label_setelah_tanggal","Subjek")

    # Info voucher
    pdf.set_font("Arial", "", 10)
    pdf.set_x(header_x)
    pdf.cell(40, 6, f"Nomor Voucher", align="L")
    pdf.cell(5, 6, ":", align="L")
    pdf.multi_cell(35, 6, no_voucher, align="L")

    pdf.set_x(header_x)
    pdf.cell(40, 6, f"Tanggal", align="L")
    pdf.cell(5, 6, ":", align="L")
    pdf.multi_cell(35, 6, tgl, align="L")

    pdf.set_x(header_x)
    pdf.cell(40, 6, f"{subjek_label}", align="L")
    pdf.cell(5, 6, ":", align="L")
    pdf.multi_cell(35, 6, str(settings.get("subjek","")), align="L")
    pdf.ln(2)

    # ========== Tabel Utama ==========
    total_width = pdf.w - pdf.l_margin - pdf.r_margin
    col_widths = [25, 60, 50, 30, 30]
    headers = ["Akun Perkiraan","Nama Akun","Memo","Debit","Kredit"]

    pdf.set_font("Arial","B",9)
    for h,w in zip(headers,col_widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln()

    pdf.set_font("Arial","",9)
    total_debit, total_kredit = 0,0
    first_desc = ""

    for _, row in data.iterrows():
        debit_val = row["Debet"]
        kredit_val = row["Kredit"]

        memo_text = ""
        if row.get("Departemen"): memo_text += f"- Departemen : {row['Departemen']}\n"
        if row.get("Proyek"): memo_text += f"- Proyek     : {row['Proyek']}"

        values = [
            str(row.get("No Akun","")),
            str(row.get("Akun","")),
            memo_text.strip(),
            fmt_num(debit_val),
            fmt_num(kredit_val)
        ]

        # wrap rows
        line_counts = []
        for i2,(val,w) in enumerate(zip(values,col_widths)):
            if i2 in [3,4]:
                line_counts.append(1)
            else:
                lines = pdf.multi_cell(w,5,val,split_only=True)
                line_counts.append(len(lines))
        row_height = max(line_counts)*5

        x_start = pdf.get_x()
        y_start = pdf.get_y()
        for i2,(val,w) in enumerate(zip(values,col_widths)):
            pdf.rect(x_start,y_start,w,row_height)
            pdf.set_xy(x_start,y_start)
            if i2 in [3,4]:
                pdf.cell(w,row_height,val,align="R")
            else:
                pdf.multi_cell(w,5,val,align="L")
            x_start+=w
        pdf.set_y(y_start+row_height)

        total_debit += debit_val
        total_kredit += kredit_val
        if not first_desc and str(row.get("Deskripsi","")).strip():
            first_desc = str(row["Deskripsi"]).strip()

    # Total
    pdf.set_font("Arial","B",9)
    pdf.cell(sum(col_widths[:-2]),8,"Total",border=1,align="R")
    pdf.cell(col_widths[3],8,fmt_num(total_debit),border=1,align="R")
    pdf.cell(col_widths[4],8,fmt_num(total_kredit),border=1,align="R")
    pdf.ln()

    # ========== Terbilang ==========
    terbilang = num2words(total_debit, lang="id")
    terbilang = " ".join([w.capitalize() for w in terbilang.split()])
    pdf.ln(2)
    pdf.set_font("Arial","I",9)
    pdf.cell(25,8,"Terbilang",border=1,align="C")
    pdf.cell(total_width-25,8,terbilang+" Rupiah",border=1,align="L")
    pdf.ln(10)

    # ========== Keterangan + TTD ==========
    col_ket = (total_width*0.45)
    col_ttd = (total_width*0.55)

    y_start = pdf.get_y()
    pdf.set_font("Arial","",9)

    # blok keterangan
    pdf.multi_cell(col_ket,6,"Keterangan\n"+first_desc, border=1)
    pdf.line(pdf.get_x(), pdf.get_y(), pdf.get_x()+col_ket, pdf.get_y())  # garis putus bisa ditambahkan manual

    # blok tanda tangan
    pdf.set_xy(pdf.l_margin+col_ket, y_start)
    pdf.set_font("Arial","B",9)
    pdf.cell(col_ttd/3,8,settings.get("jab1",""),1,0,"C")
    pdf.cell(col_ttd/3,8,settings.get("jab2",""),1,0,"C")
    pdf.cell(col_ttd/3,8,settings.get("jab3",""),1,1,"C")

    pdf.set_font("Arial","",9)
    pdf.cell(col_ttd/3,20,"",1,0,"C")
    pdf.cell(col_ttd/3,20,"",1,0,"C")
    pdf.cell(col_ttd/3,20,"",1,1,"C")

    pdf.cell(col_ttd/3,8,settings.get("nm1",""),1,0,"C")
    pdf.cell(col_ttd/3,8,settings.get("nm2",""),1,0,"C")
    pdf.cell(col_ttd/3,8,settings.get("nm3",""),1,1,"C")

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer

# ================= STREAMLIT ==================
st.set_page_config(page_title="Voucher Generator", layout="wide")
st.title("📑 Voucher Generator")

settings = {}
settings["perusahaan"] = st.sidebar.text_input("Nama Perusahaan")
settings["alamat"] = st.sidebar.text_area("Alamat Perusahaan")
settings["logo_size"] = st.sidebar.slider("Ukuran Logo (mm)",10,50,20)
logo_file = st.sidebar.file_uploader("Upload Logo",type=["png","jpg","jpeg"])
if logo_file: settings["logo"]=BytesIO(logo_file.read())

# custom judul, subjek, pejabat
settings["judul_doc"] = st.text_input("Judul Dokumen","Jurnal Voucher")
settings["label_setelah_tanggal"] = st.text_input("Label setelah Tanggal","Subjek")
settings["subjek"] = st.text_input("Isi Subjek/Pemberi/Penerima","")

settings["jab1"] = st.text_input("Jabatan 1","Finance")
settings["nm1"]  = st.text_input("Nama Pejabat 1","")
settings["jab2"] = st.text_input("Jabatan 2","Disetujui")
settings["nm2"]  = st.text_input("Nama Pejabat 2","")
settings["jab3"] = st.text_input("Jabatan 3","Penerima")
settings["nm3"]  = st.text_input("Nama Pejabat 3","")

file = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx","xls"])
if file:
    df = pd.read_excel(file)
    df = df.rename(columns=lambda x:str(x).strip())
    no_voucher = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())
    if st.button("Cetak PDF"):
        pdf_file = buat_voucher(df,no_voucher,settings)
        st.download_button("⬇️ Download PDF",data=pdf_file,file_name=f"{no_voucher}.pdf")
