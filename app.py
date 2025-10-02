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
        "memo": "Memo",
        "departemen": "Departemen",
        "proyek": "Proyek",
        "deskripsi": "Deskripsi",
        "debet": "Debet",
        "kredit": "Kredit"
    }
    df = df.rename(columns={k.lower(): v for k, v in mapping.items() if k.lower() in df.columns})
    df["Debet"] = pd.to_numeric(df.get("Debet", 0), errors="coerce").fillna(0)
    df["Kredit"] = pd.to_numeric(df.get("Kredit", 0), errors="coerce").fillna(0)
    return df

# --- generate voucher ---
def buat_voucher(df, no_voucher, settings, jenis_doc="Jurnal Voucher"):
    pdf = FPDF("P", "mm", "A4")
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.add_page()

    # Logo
    if settings.get("logo"):
        pdf.image(settings["logo"], 15, 8, settings.get("logo_size", 20))

    # Nama perusahaan
    pdf.set_xy(40, 10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(90, 6, settings.get("perusahaan", ""), ln=1, align="L")

    # Alamat perusahaan (wrap max width 90mm)
    pdf.set_xy(40, 16)
    pdf.set_font("Arial", "", 9)
    pdf.multi_cell(90, 5, settings.get("alamat", ""), align="L")

    # Header kanan (judul + info)
    judul = jenis_doc
    header_width = 80
    header_x = pdf.w - pdf.r_margin - header_width
    pdf.set_xy(header_x, 10)
    pdf.set_font("Arial", "B", 13)
    pdf.cell(header_width, 8, judul, align="R")
    pdf.ln(8)

    # garis atas & bawah judul
    pdf.set_draw_color(0, 0, 0)
    pdf.line(header_x, 10, pdf.w - pdf.r_margin, 10)
    pdf.line(header_x, 18, pdf.w - pdf.r_margin, 18)

    # Ambil data voucher
    data = df[df["Nomor Voucher Jurnal"] == no_voucher]
    try:
        tgl = pd.to_datetime(data.iloc[0]["Tanggal"]).strftime("%d %b %Y")
    except:
        tgl = str(data.iloc[0]["Tanggal"])

    # Subjek label sesuai dokumen
    if jenis_doc == "Jurnal Voucher":
        subjek_label = "Subjek"
    elif jenis_doc == "Bukti Pengeluaran Kas/Bank":
        subjek_label = "Penerima"
    else:
        subjek_label = "Pemberi"

    # Info box
    label_w = 35
    value_w = header_width - label_w
    pdf.set_font("Arial", "", 10)

    pdf.set_x(header_x)
    pdf.cell(label_w, 7, "Nomor Voucher :", border=0)
    pdf.multi_cell(value_w, 7, no_voucher, border=0)

    pdf.set_x(header_x)
    pdf.cell(label_w, 7, "Tanggal :", border=0)
    pdf.multi_cell(value_w, 7, tgl, border=0)

    pdf.set_x(header_x)
    pdf.cell(label_w, 7, f"{subjek_label} :", border=0)
    pdf.multi_cell(value_w, 7, "", border=0)

    pdf.ln(3)

    # Table header
    base_col_widths = [35, 70, 55, 35, 35]  # auto adjust lebar
    headers = ["Akun Perkiraan", "Nama Akun", "Memo", "Debit", "Kredit"]

    pdf.set_font("Arial", "B", 9)
    for h, w in zip(headers, base_col_widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln()

    # Isi tabel
    total_debit, total_kredit = 0, 0
    pdf.set_font("Arial", "", 9)
    first_desc = ""

    for _, row in data.iterrows():
        debit_val = int(row["Debet"]) if pd.notna(row["Debet"]) else 0
        kredit_val = int(row["Kredit"]) if pd.notna(row["Kredit"]) else 0
        akun = str(row.get("No Akun", ""))
        nama_akun = str(row.get("Akun", ""))
        departemen = str(row.get("Departemen", ""))
        proyek = str(row.get("Proyek", ""))
        memo = f"- Departemen : {departemen}\n- Proyek : {proyek}\n{str(row.get('Memo',''))}".strip()

        values = [
            akun,
            nama_akun,
            memo,
            f"{debit_val:,.0f}".replace(",", "."),
            f"{kredit_val:,.0f}".replace(",", ".")
        ]

        # tinggi baris wrap
        line_counts = []
        for i2, (val, w) in enumerate(zip(values, base_col_widths)):
            if i2 in [3, 4]:
                line_counts.append(1)
            else:
                lines = pdf.multi_cell(w, 6, val, split_only=True)
                line_counts.append(len(lines))
        max_lines = max(line_counts)
        row_height = max_lines * 6

        x_start = pdf.get_x()
        y_start = pdf.get_y()
        for i2, (val, w) in enumerate(zip(values, base_col_widths)):
            pdf.rect(x_start, y_start, w, row_height)
            pdf.set_xy(x_start, y_start)
            if i2 in [3, 4]:
                pdf.cell(w, row_height, val, align="R")
            else:
                pdf.multi_cell(w, 6, val, align="L")
            x_start += w
        pdf.set_y(y_start + row_height)

        total_debit += debit_val
        total_kredit += kredit_val
        if first_desc == "" and str(row.get("Deskripsi", "")).strip():
            first_desc = str(row.get("Deskripsi")).strip()

    # Total row
    pdf.set_font("Arial", "B", 9)
    pdf.cell(sum(base_col_widths[:-2]), 8, "Total", border=1, align="R")
    pdf.cell(base_col_widths[3], 8, f"{total_debit:,.0f}".replace(",", "."), border=1, align="R")
    pdf.cell(base_col_widths[4], 8, f"{total_kredit:,.0f}".replace(",", "."), border=1, align="R")
    pdf.ln()

    # Terbilang
    pdf.set_font("Arial", "I", 9)
    pdf.cell(sum(base_col_widths), 8, f"Terbilang : {num2words(total_debit, lang='id').capitalize()} Rupiah", border=1, align="L")
    pdf.ln()

    # Deskripsi
    if first_desc:
        pdf.set_font("Arial", "", 9)
        pdf.multi_cell(sum(base_col_widths), 8, f"Keterangan : {first_desc}", border=1, align="L")
        pdf.ln()

    # Tanda tangan
    pdf.ln(5)
    pdf.set_font("Arial", "", 10)
    ttd_fields = ["Finance", "Disetujui", "Penerima"]
    col_width = (pdf.w - pdf.l_margin - pdf.r_margin) / len(ttd_fields)

    for role in ttd_fields:
        pdf.cell(col_width, 8, role, border=1, align="C")
    pdf.ln(20)

    for role in ttd_fields:
        pdf.cell(col_width, 8, "(................)", border=1, align="C")
    pdf.ln(10)

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer

# --- Streamlit ---
st.set_page_config(page_title="Mini Akunting", layout="wide")
st.title("📑 Mini Akunting - Voucher Jurnal/Kas/Bank")

# Sidebar
st.sidebar.header("⚙️ Pengaturan Perusahaan")
settings = {}
settings["perusahaan"] = st.sidebar.text_input("Nama Perusahaan")
settings["alamat"] = st.sidebar.text_area("Alamat Perusahaan")
settings["logo_size"] = st.sidebar.slider("Ukuran Logo (mm)", 10, 50, 20)
logo_file = st.sidebar.file_uploader("Upload Logo (PNG/JPG)", type=["png", "jpg", "jpeg"])
if logo_file:
    tmp = BytesIO(logo_file.read())
    settings["logo"] = tmp

# Main content
file = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx", "xls"])
if file:
    df = pd.read_excel(file)
    df = bersihkan_jurnal(df)
    st.dataframe(df.head())

    mode = st.radio("Pilih Mode Cetak", ["Single Voucher", "Per Bulan"])
    jenis_doc = st.selectbox("Jenis Dokumen", ["Jurnal Voucher", "Bukti Pengeluaran Kas/Bank", "Bukti Penerimaan Kas/Bank"])

    if mode == "Single Voucher":
        no_voucher = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())
        if st.button("Cetak Voucher"):
            pdf_file = buat_voucher(df, no_voucher, settings, jenis_doc)
            st.download_button("⬇️ Download PDF", data=pdf_file, file_name=f"{no_voucher}.pdf")

    else:  # per bulan
        df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
        df = df.dropna(subset=["Tanggal"])
        bulan = st.selectbox("Pilih Bulan", range(1, 13), format_func=lambda x: calendar.month_name[x])
        if st.button("Cetak Semua Voucher Bulan Ini"):
            buffer_zip = BytesIO()
            with zipfile.ZipFile(buffer_zip, "w") as zf:
                for v in df[df["Tanggal"].dt.month == bulan]["Nomor Voucher Jurnal"].unique():
                    pdf_file = buat_voucher(df, v, settings, jenis_doc)
                    zf.writestr(f"{v}.pdf", pdf_file.getvalue())
            buffer_zip.seek(0)
            st.download_button("⬇️ Download ZIP", data=buffer_zip, file_name=f"voucher_{bulan}.zip", mime="application/zip")
