import io
import os
import numpy as np
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.colors import blue, black, lightgrey
from datetime import datetime

# === Mapel per jenjang ===
# Daftar mata pelajaran untuk Kelas 8 dan Kelas 9 (Mengandung Seni Budaya)
mapel_kelas_8_9 = [
    "Pend. Agama dan Budi Pekerti",
    "Pendidikan Pancasila",
    "Bahasa Indonesia",
    "Matematika",
    "Ilmu Pengetahuan Alam",
    "Ilmu Pengetahuan Sosial",
    "Bahasa Inggris",
    "PJOK",
    "Informatika",
    "Seni Budaya",  # Khusus Kelas 8 & 9
    "Bahasa Jawa"
]

# Daftar mata pelajaran untuk Kelas 7 (Mengandung Prakarya)
mapel_kelas_7 = [
    "Pend. Agama dan Budi Pekerti",
    "Pendidikan Pancasila",
    "Bahasa Indonesia",
    "Matematika",
    "Ilmu Pengetahuan Alam",
    "Ilmu Pengetahuan Sosial",
    "Bahasa Inggris",
    "PJOK",
    "Informatika",
    "Prakarya",    # Khusus Kelas 7
    "Bahasa Jawa"
]

# Gabungan semua mapel (untuk template Excel)
mapel_semua = sorted(set(mapel_kelas_8_9) | set(mapel_kelas_7))

# Helper penentu daftar mapel berdasarkan nama kelas
def get_mapel_by_kelas(kelas_name):
    kelas_str = str(kelas_name).upper().strip()
    # Jika diawali "7" atau "VII", gunakan mapel Kelas 7 (Prakarya)
    if kelas_str.startswith("VII") or kelas_str.startswith("7"):
        return mapel_kelas_7
    # Untuk Kelas 8 dan Kelas 9, gunakan mapel Kelas 8 & 9 (Seni Budaya)
    else:
        return mapel_kelas_8_9

# Mapping bulan Indonesia
bulan_id = {
    "January": "Januari", "February": "Februari", "March": "Maret",
    "April": "April", "May": "Mei", "June": "Juni",
    "July": "Juli", "August": "Agustus", "September": "September",
    "October": "Oktober", "November": "November", "December": "Desember"
}

st.header("Laporan Hasil Asesmen")
st.markdown("---")

# --- Pilihan di Streamlit ---
asesmen_opsi = [
    "ASESMEN SUMATIF TENGAH SEMESTER GENAP",
    "ASESMEN SUMATIF AKHIR SEMESTER GANJIL",
    "ASESMEN SUMATIF TENGAH SEMESTER GANJIL",
    "ASESMEN SUMATIF AKHIR TAHUN SEMESTER GENAP"
]
sel_asesmen = st.selectbox("Pilih Jenis Asesmen", asesmen_opsi)

tahun_opsi = [f"{th}/{th+1}" for th in range(2025, 2036)]
sel_tahun = st.selectbox("Pilih Tahun Pelajaran", tahun_opsi, index=0)

sel_tgl_ttd = st.date_input("Tanggal Penulisan Tanda Tangan (di dokumen PDF)", datetime.now(), format="DD/MM/YYYY")

st.markdown("---")

# === Template Excel ===
def generate_template():
    cols = ["Kelas", "NIS", "Nama Siswa"] + mapel_semua
    df_template = pd.DataFrame(columns=cols)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_template.to_excel(writer, index=False, sheet_name="Nilai")
    buffer.seek(0)
    return buffer

st.download_button(
    "📥 Download Template Excel (Semua Kelas)",
    data=generate_template(),
    file_name="Template_Nilai.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Upload file Excel
uploaded = st.file_uploader("Unggah file Excel daftar nilai (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Silakan unggah file Excel nilai (menggunakan template).")
    st.stop()

try:
    df = pd.read_excel(uploaded, engine="openpyxl")
except Exception as e:
    st.error(f"Gagal membaca file Excel: {e}")
    st.stop()

df.columns = df.columns.str.strip()

if "Kelas" not in df.columns:
    st.error("Kolom 'Kelas' tidak ditemukan di file. Pastikan pakai template.")
    st.stop()

# Pastikan seluruh kolom mapel ada di DataFrame agar tidak KeyError
for m in mapel_semua:
    if m not in df.columns:
        df[m] = np.nan

# Bersihkan & konversi nilai numerik
for col in mapel_semua:
    df[col] = (
        df[col]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
        .str.strip()
    )
    df.loc[df[col] == "", col] = np.nan
    df[col] = pd.to_numeric(df[col], errors="coerce")

# Ambil daftar semua kelas unik dari file Excel
semua_kelas_di_excel = sorted(df["Kelas"].astype(str).unique())
sel_kelas = st.selectbox("Pilih Kelas", semua_kelas_di_excel)

semua_paralel = st.checkbox("Cetak semua kelas paralel?", value=False)

if semua_paralel:
    prefix = str(sel_kelas).strip()[0]
    df_kelas = df[df["Kelas"].astype(str).str.startswith(prefix)].copy()
else:
    df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()

# Pastikan kolom wajib dasar ada
expected_base = ["Kelas", "NIS", "Nama Siswa"]
missing_base = [c for c in expected_base if c not in df.columns]
if missing_base:
    st.error(f"Kolom wajib hilang: {missing_base}")
    st.stop()

siswa_list = df_kelas["Nama Siswa"].astype(str).tolist()
sel_siswa = st.selectbox("Pilih Siswa", ["-- Semua Siswa --"] + siswa_list)

def format_score(val):
    if pd.isna(val):
        return ""
    try:
        return f"{float(val):.2f}"
    except Exception:
        return str(val)

# Fungsi menggambar halaman PDF per siswa
def draw_student_page(c, row, sel_asesmen, sel_tahun, mapel_target, sel_tgl_ttd):
    width, height = A4

    margin_left   = 30 * mm
    margin_right  = 20 * mm
    margin_top    = 20 * mm
    margin_bottom = 20 * mm

    content_width = width - (margin_left + margin_right)
    y = height - margin_top

    # Logo Kiri
    logo_path = "assets/logo_kiri.png"
    if os.path.exists(logo_path):
        try:
            logo_w, logo_h = 30 * mm, 30 * mm
            x_logo = margin_left - 10 * mm
            y_logo = height - margin_top - (-10 * mm) - logo_h
            c.drawImage(logo_path, x_logo, y_logo, width=logo_w, height=logo_h, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    # Logo Kanan
    logo_kanan_path = "assets/logo_kanan.png"
    if os.path.exists(logo_kanan_path):
        try:
            logo_w, logo_h = 30 * mm, 30 * mm
            x_logo_kanan = width - margin_right - logo_w + 5 * mm
            y_logo_kanan = height - margin_top - (-10 * mm) - logo_h 
            c.drawImage(logo_kanan_path, x_logo_kanan, y_logo_kanan, width=logo_w, height=logo_h, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    # Kop Surat
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width / 2, y, "PEMERINTAH KABUPATEN BANTUL")
    y -= 5 * mm
    c.drawCentredString(width / 2, y, "DINAS PENDIDIKAN, KEPEMUDAAN, DAN OLAHRAGA")
    y -= 5 * mm
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2, y, "SMP NEGERI 2 BANGUNTAPAN")
    y -= 1 * mm

    aksara_path = "assets/aksara_jawa.jpg"
    if os.path.exists(aksara_path):
        try:
            aksara_w, aksara_h = 100 * mm, 10 * mm
            x_aksara = (width - aksara_w) / 2
            y_aksara = y - aksara_h
            c.drawImage(aksara_path, x_aksara, y_aksara, width=aksara_w, height=aksara_h, preserveAspectRatio=True, mask='auto')
            y = y_aksara - 2 * mm
        except Exception:
            y -= 4 * mm
    else:
        y -= 4 * mm

    c.setFont("Helvetica-Oblique", 10)
    c.drawCentredString(width / 2, y, "Jalan Karangsari, Banguntapan, Kabupaten Bantul, Yogyakarta 55198")
    y -= 5 * mm
    c.drawCentredString(width / 2, y, "Telp. (0274) 382754 382754")
    y -= 5 * mm
    c.setFont("Helvetica", 10)
    c.setFillColor(blue)
    c.drawCentredString(width / 2, y, "Laman : www.smpn2banguntapan.sch.id; Pos-el : smp2banguntapan@yahoo.com")
    c.setFillColor(black)
    y -= 3 * mm

    # Garis Pembatas Kop
    c.setLineWidth(1)
    c.line(margin_left, y, width - margin_right, y)
    y -= 1.5 * mm
    c.setLineWidth(0.5)
    c.line(margin_left, y, width - margin_right, y)
    y -= 10 * mm

    # Judul Dokumen
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width / 2, y, f"LAPORAN HASIL {sel_asesmen}")
    y -= 6 * mm
    c.drawCentredString(width / 2, y, f"TAHUN PELAJARAN {sel_tahun}")
    y -= 12 * mm

    # Identitas Siswa
    id_margin_left = margin_left + 10 * mm
    label_w = 20 * mm
    colon_x = id_margin_left + label_w
    value_x = colon_x + 5
    c.setFont("Helvetica", 12)

    c.drawString(id_margin_left, y, "Nama")
    c.drawString(colon_x, y, ":")
    c.drawString(value_x, y, " " + str(row.get("Nama Siswa", "")))
    y -= 6 * mm

    c.drawString(id_margin_left, y, "NIS")
    c.drawString(colon_x, y, ":")
    c.drawString(value_x, y, " " + str(row.get("NIS", "")))
    y -= 6 * mm

    c.drawString(id_margin_left, y, "Kelas")
    c.drawString(colon_x, y, ":")
    c.drawString(value_x, y, " " + str(row.get("Kelas", "")))
    y -= 10 * mm

# Pengambilan Nilai (Dengan penanganan otomatis jika salah isi kolom di Excel)
    nilai_list = []
    for subj in mapel_target:
        raw = row.get(subj, np.nan)
        
        # Jika Seni Budaya kosong di Excel tapi Prakarya ada isinya, ambil dari Prakarya
        if subj == "Seni Budaya" and (pd.isna(raw) or str(raw).strip() == ""):
            raw = row.get("Prakarya", np.nan)
            
        # Jika Prakarya kosong di Excel tapi Seni Budaya ada isinya, ambil dari Seni Budaya
        elif subj == "Prakarya" and (pd.isna(raw) or str(raw).strip() == ""):
            raw = row.get("Seni Budaya", np.nan)

        try:
            nilai_list.append(float(raw) if pd.notna(raw) else np.nan)
        except Exception:
            nilai_list.append(np.nan)

    nilai_series = pd.Series(nilai_list, index=mapel_target, dtype="float64")
    jumlah = float(nilai_series.sum(skipna=True))
    rata2 = float(nilai_series.mean(skipna=True)) if nilai_series.count() > 0 else 0.0

    # Format Tabel
    row_height = 7 * mm
    font_size = 11
    col_no_w, col_mapel_w, col_nilai_w = 15 * mm, 90 * mm, 25 * mm
    table_width = col_no_w + col_mapel_w + col_nilai_w

    x0 = margin_left + (content_width - table_width) / 2
    y0 = y
    nrows = len(mapel_target) + 3

    # Header Background
    c.setFillColor(lightgrey)
    c.rect(x0, y0 - row_height, table_width, row_height, stroke=0, fill=1)
    c.setFillColor(black)

    # Line Grid
    for r in range(nrows + 1):
        c.setLineWidth(0.5)
        c.line(x0, y0 - r * row_height, x0 + table_width, y0 - r * row_height)
    c.line(x0, y0, x0, y0 - nrows * row_height)
    c.line(x0 + col_no_w, y0, x0 + col_no_w, y0 - nrows * row_height)
    c.line(x0 + col_no_w + col_mapel_w, y0, x0 + col_no_w + col_mapel_w, y0 - nrows * row_height)
    c.line(x0 + table_width, y0, x0 + table_width, y0 - nrows * row_height)

    # Header Text
    c.setFont("Helvetica-Bold", font_size)
    adj_y = (y0 - row_height / 2) - (font_size / 3.5)
    c.drawCentredString(x0 + col_no_w / 2, adj_y, "No")
    c.drawCentredString(x0 + col_no_w + col_mapel_w / 2, adj_y, "Mata Pelajaran")
    c.drawCentredString(x0 + col_no_w + col_mapel_w + col_nilai_w / 2, adj_y, "Nilai")

    # Body Text
    c.setFont("Helvetica", font_size)
    y_text = y0 - row_height
    for i, subj in enumerate(mapel_target, start=1):
        adj_y = (y_text - row_height / 2) - (font_size / 3.5)
        val = nilai_series.get(subj, np.nan)
        c.drawCentredString(x0 + col_no_w / 2, adj_y, str(i))
        c.drawString(x0 + col_no_w + 2 * mm, adj_y, subj)
        c.drawCentredString(x0 + col_no_w + col_mapel_w + col_nilai_w / 2, adj_y, val_str := format_score(val))
        y_text -= row_height

    # Jumlah & Rata-rata
    adj_y = (y_text - row_height / 2) - (font_size / 3.5)
    c.setFont("Helvetica-Bold", font_size)
    c.drawString(x0 + col_no_w + 2 * mm, adj_y, "Jumlah")
    c.drawCentredString(x0 + col_no_w + col_mapel_w + col_nilai_w / 2, adj_y, format_score(jumlah))
    y_text -= row_height

    adj_y = (y_text - row_height / 2) - (font_size / 3.5)
    c.drawString(x0 + col_no_w + 2 * mm, adj_y, "Rata-rata")
    c.drawCentredString(x0 + col_no_w + col_mapel_w + col_nilai_w / 2, adj_y, format_score(rata2))
    y_text -= row_height + 20

    # Tanda tangan
    ttd_date = sel_tgl_ttd
    bulan_eng = ttd_date.strftime('%B')
    tgl = f"{ttd_date.day} {bulan_id.get(bulan_eng, bulan_eng)} {ttd_date.year}"

    x_ttd = width - margin_right - 70 * mm
    y_ttd_start = margin_bottom + 62 * mm

    c.setFont("Helvetica", 12)
    c.drawString(x_ttd, y_ttd_start, f"Banguntapan, {tgl}")
    y_ttd_start -= 8 * mm
    c.drawString(x_ttd, y_ttd_start, "Mengetahui,")
    y_ttd_start -= 5 * mm
    c.drawString(x_ttd, y_ttd_start, "Kepala Sekolah,")
    y_ttd_start -= -1 * mm

    ttd_path = "assets/ttd_kepsek.jpeg"
    if os.path.exists(ttd_path):
        try:
            c.drawImage(ttd_path, x_ttd, y_ttd_start - 22 * mm, width=40 * mm, height=20 * mm, mask="auto")
        except Exception:
            pass

    y_ttd_after = y_ttd_start - 25 * mm
    c.drawString(x_ttd, y_ttd_after, "Alina Fiftiyani Nurjannah, M.Pd.")
    y_ttd_after -= 6 * mm
    c.drawString(x_ttd, y_ttd_after, "NIP 198001052009032006")

# Fungsi Generator PDF
def make_pdf_for_student(row, sel_tgl_ttd):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    mapel_target = get_mapel_by_kelas(row["Kelas"])
    draw_student_page(c, row, sel_asesmen, sel_tahun, mapel_target, sel_tgl_ttd)
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

def make_pdf_for_class(df_k, sel_tgl_ttd):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    for _, row in df_k.iterrows():
        mapel_target = get_mapel_by_kelas(row["Kelas"])
        draw_student_page(c, row, sel_asesmen, sel_tahun, mapel_target, sel_tgl_ttd)
        c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

def make_pdf_for_all_classes(df_all, kelas_list_all, sel_tgl_ttd):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    for k in kelas_list_all:
        df_sel = df_all[df_all["Kelas"].astype(str) == str(k)]
        for _, row in df_sel.iterrows():
            mapel_target = get_mapel_by_kelas(row["Kelas"])
            draw_student_page(c, row, sel_asesmen, sel_tahun, mapel_target, sel_tgl_ttd)
            c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# --- Tombol Unduh ---
st.markdown("---")
st.subheader("Pilih Siswa & Unduh Laporan")

if not df_kelas.empty:
    if sel_siswa != "-- Semua Siswa --":
        row_siswa = df_kelas[df_kelas["Nama Siswa"] == sel_siswa].iloc[0]
        st.download_button(
            "📄 Download PDF (Per Siswa)",
            data=make_pdf_for_student(row_siswa, sel_tgl_ttd),
            file_name=f"Laporan_{row_siswa['Nama Siswa']}.pdf",
            mime="application/pdf"
        )

    st.download_button(
        "📄 Download PDF (Per Kelas)",
        data=make_pdf_for_class(df_kelas, sel_tgl_ttd),
        file_name=f"Laporan_{sel_kelas}.pdf",
        mime="application/pdf"
    )
else:
    st.warning("Tidak ada data siswa untuk kelas yang dipilih.")

# Tombol Download Semua Kelas diletakkan secara terpisah agar SELALU MUNCUL
st.markdown("---")
st.download_button(
    "📚 Download PDF (Semua Kelas di Excel)",
    data=make_pdf_for_all_classes(df, semua_kelas_di_excel, sel_tgl_ttd),
    file_name=f"Laporan_Semua_Kelas_{sel_tahun}.pdf",
    mime="application/pdf"
)
