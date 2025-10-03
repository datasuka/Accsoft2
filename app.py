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

# format angka Indonesia
def fmt_num(val):
    try:
        return "{:,.0f}".format(float(val)).replace(",", ".")
    except:
        return "0"

# --- generate voucher ---
def buat_voucher(df, no_voucher, settings, pejabat, ttd_height):
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

    # --- Judul Dokumen ---
    judul = settings.get("judul_dokumen", "Bukti Jurnal")
    header_x = 120
    pdf.set_xy(header_x, 10)
    pdf.set_font("Arial", "B", 12)
    pdf.set_line_width(0.8)
    pdf.cell(80, 10, judul.upper(), border="TB", align="C", ln=1)
    pdf.set_line_width(0.2)

    # Info voucher
    data = df[df["Nomor Voucher Jurnal"] == no_voucher]
    try:
        tgl = pd.to_datetime(data.iloc[0]["Tanggal"]).strftime("%d %b %Y")
    except:
        tgl = str(data.iloc[0]["Tanggal"])
    subjek_val = str(data.iloc[0].get("Subjek",""))

    pdf.set_x(header_x)
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

    # --- Tabel utama ---
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

    # --- TERBILANG ---
    terbilang = num2words(total_debit, lang='id')
    terbilang = " ".join([w.capitalize() for w in terbilang.split()])
    if "Koma Nol" in terbilang:
        terbilang = terbilang.replace("Koma Nol", "")
    pdf.set_font("Arial", "I", 9)
    pdf.cell(total_width, 8, f"Terbilang : {terbilang} Rupiah", border=1, align="L")
    pdf.ln(10)

    # --- KETERANGAN & TTD ---
    ket_width = total_width * 0.4
    ttd_width = total_width * 0.6

    y_start = pdf.get_y()

    # Keterangan box
    pdf.set_font("Arial", "", 9)
    pdf.rect(pdf.l_margin, y_start, ket_width, ttd_height)
    pdf.set_xy(pdf.l_margin + 2, y_start + 2)
    pdf.cell(0, 6, "Keterangan", ln=1)
    if first_desc:
        pdf.set_x(pdf.l_margin + 2)
        pdf.multi_cell(ket_width - 4, 6, str(first_desc))
    # garis putus-putus
    y_dashed = y_start + ttd_height - 5
    pdf.dashed_line(pdf.l_margin + 2, y_dashed, pdf.l_margin + ket_width - 2, y_dashed, 1, 2)

    # TTD box
    pdf.set_xy(pdf.l_margin + ket_width + 5, y_start)
    col_width = (ttd_width - 5) / len(pejabat)

    # Header jabatan
    pdf.set_font("Arial", "B", 9)
    for jabatan, _ in pejabat:
        pdf.cell(col_width, 8, jabatan if jabatan else "", border=1, align="C")
    pdf.ln()

    # Kotak tanda tangan (custom tinggi)
    pdf.set_x(pdf.l_margin + ket_width + 5)
    for _ in pejabat:
        pdf.cell(col_width, ttd_height-16, "", border=1, align="C")
    pdf.ln()

    # Nama pejabat
    pdf.set_x(pdf.l_margin + ket_width + 5)
    pdf.set_font("Arial", "", 9)
    for _, nama in pejabat:
        pdf.cell(col_width, 8, nama if nama else "", border=1, align="C")
    pdf.ln(15)

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer

# --- Streamlit ---
st.set_page_config(page_title="Bukti Transaksi Akuntansi Generator", layout="wide")
st.title("📑 Bukti Transaksi Akuntansi Generator")
st.caption("By @zavibis")

st.markdown("""
Aplikasi ini digunakan untuk membuat **Bukti Jurnal / Kas / Bank** secara otomatis dari file Excel.  
✅ Tidak menyimpan data Anda di server/web  
✅ Semua pemrosesan terjadi di perangkat Anda.  

**Langkah penggunaan:**
1. Siapkan file Excel dengan format kolom:  
   `Tanggal | Nomor Voucher Jurnal | No Akun | Akun | Deskripsi | Debet | Kredit | Departemen | Proyek | Subjek`  
2. Pastikan nilai **Debet dan Kredit** seimbang.  
3. Upload Excel ke aplikasi ini.  
4. Pilih mode cetak **Single Voucher** atau **Per Bulan**.  
5. Download hasil PDF/ZIP.
---
""")

# Sidebar
st.sidebar.header("⚙️ Pengaturan Perusahaan")
settings = {}
settings["perusahaan"] = st.sidebar.text_input("Nama Perusahaan")
settings["alamat"] = st.sidebar.text_area("Alamat Perusahaan")
settings["logo_size"] = st.sidebar.slider("Ukuran Logo (mm)", 10, 50, 20)
settings["judul_dokumen"] = st.sidebar.text_input("Judul Dokumen", "Bukti Jurnal")
settings["label_subjek"] = st.sidebar.text_input("Label setelah Tanggal", "Subjek")

logo_file = st.sidebar.file_uploader("Upload Logo (PNG/JPG)", type=["png","jpg","jpeg"])
if logo_file:
    tmp = BytesIO(logo_file.read())
    settings["logo"] = tmp

st.sidebar.subheader("Kolom Tanda Tangan (max 3)")
pejabat = []
for i in range(1,4):
    jabatan = st.sidebar.text_input(f"Jabatan {i}", "")
    nama = st.sidebar.text_input(f"Nama Pejabat {i}", "")
    if jabatan or nama:
        pejabat.append((jabatan, nama))

# Custom tinggi TTD
ttd_height = st.sidebar.slider("Tinggi Kolom TTD (mm)", 30, 80, 40)

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
            pdf_file = buat_voucher(df, no_voucher, settings, pejabat, ttd_height)
            st.download_button("⬇️ Download PDF", data=pdf_file, file_name=f"{no_voucher}.pdf")

    else:
        df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
        df = df.dropna(subset=["Tanggal"])
        bulan = st.selectbox("Pilih Bulan", range(1,13), format_func=lambda x: calendar.month_name[x])

        if st.button("Cetak Semua Voucher Bulan Ini"):
            buffer_zip = BytesIO()
            with zipfile.ZipFile(buffer_zip, "w") as zf:
                for v in df[df["Tanggal"].dt.month==bulan]["Nomor Voucher Jurnal"].unique():
                    pdf_file = buat_voucher(df, v, settings, pejabat, ttd_height)
                    zf.writestr(f"{v}.pdf", pdf_file.getvalue())
            buffer_zip.seek(0)
            st.download_button("⬇️ Download ZIP", data=buffer_zip, file_name=f"voucher_{bulan}.zip", mime="application/zip")
