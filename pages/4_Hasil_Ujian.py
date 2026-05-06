import io
import os  # <-- ditambahkan agar os.path.exists() bisa dipakai untuk tanda tangan
import numpy as np
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.colors import blue, black, lightgrey
from datetime import datetime

# Pastikan Anda sudah menginstal reportlab dan openpyxl:
# pip install streamlit numpy pandas reportlab openpyxl openpyxl

# === Mapel per jenjang ===
# Daftar mata pelajaran untuk Kelas 7 dan Kelas 8 (Mengandung Seni Budaya)
mapel_kelas_7_8 = [
    "Pend. Agama dan Budi Pekerti",
    "Pendidikan Pancasila",
    "Bahasa Indonesia",
    "Matematika",
    "Ilmu Pengetahuan Alam",
    "Ilmu Pengetahuan Sosial",
    "Bahasa Inggris",
    "PJOK",
    "Informatika",
    "Seni Budaya", # Khusus Kelas 7 & 8
    "Bahasa Jawa"
]

# Daftar mata pelajaran untuk Kelas 9 (Mengandung Prakarya)
mapel_kelas_9 = [
    "Pend. Agama dan Budi Pekerti",
    "Pendidikan Pancasila",
    "Bahasa Indonesia",
    "Matematika",
    "Ilmu Pengetahuan Alam",
    "Ilmu Pengetahuan Sosial",
    "Bahasa Inggris",
    "PJOK",
    "Informatika",
    "Prakarya", # Khusus Kelas 9
    "Bahasa Jawa"
]

# Gabungan semua mapel (untuk template Excel)
mapel_semua = sorted(set(mapel_kelas_7_8) | set(mapel_kelas_9))

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

# Pilihan asesmen
asesmen_opsi = [
    "ASESMEN SUMATIF TENGAH SEMESTER GENAP",
    "ASESMEN SUMATIF AKHIR SEMESTER GANJIL",
    "ASESMEN SUMATIF TENGAH SEMESTER GANJIL",
    "ASESMEN SUMATIF AKHIR TAHUN SEMESTER GENAP"
]
sel_asesmen = st.selectbox("Pilih Jenis Asesmen", asesmen_opsi)

# Tahun pelajaran
tahun_opsi = [f"{th}/{th+1}" for th in range(2025, 2036)]
sel_tahun = st.selectbox("Pilih Tahun Pelajaran", tahun_opsi, index=0)

# Input tanggal kustom untuk tanda tangan
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

# baca file
try:
    df = pd.read_excel(uploaded, engine="openpyxl")
except Exception as e:
    st.error(f"Gagal membaca file Excel: {e}")
    st.stop()

df.columns = df.columns.str.strip()

# Pilih kelas dari file
if "Kelas" not in df.columns:
    st.error("Kolom 'Kelas' tidak ditemukan di file. Pastikan pakai template.")
    st.stop()

kelas_list = sorted(df["Kelas"].astype(str).unique())
sel_kelas = st.selectbox("Pilih Kelas", kelas_list)
# Tambahan: opsi cetak semua paralel (tetap aman — tidak mengubah data lama)
semua_paralel = st.checkbox("Cetak semua kelas ?", value=False)

# Jika user minta semua paralel, override df_kelas menjadi semua kelas dengan prefix yang sama
if semua_paralel:
    # ambil prefix (angka depan), aman untuk format '9A' atau 'IX...' bila sekolah pakai angka
    prefix = str(sel_kelas).strip()[0]
    df_kelas = df[df["Kelas"].astype(str).str.startswith(prefix)].copy()
else:
    df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()

# Tentukan mapel sesuai jenjang
kelas_upper = str(sel_kelas).upper().strip()

# LOGIKA UTAMA: Cek apakah kelas adalah Kelas 9 (IX atau 9)
# Jika kelas dimulai dengan "IX" atau "9", gunakan mapel Kelas 9 (Prakarya)
if kelas_upper.startswith("IX") or kelas_upper.startswith("9"):
    # Gunakan daftar mapel Kelas 9 (Prakarya)
    mapel_urut = [m for m in mapel_kelas_9 if m in df.columns]
else:
    # Gunakan daftar mapel Kelas 7/8 (Seni Budaya)
    mapel_urut = [m for m in mapel_kelas_7_8 if m in df.columns]

# Pastikan kolom penting ada
expected_base = ["Kelas", "NIS", "Nama Siswa"]
missing_base = [c for c in expected_base if c not in df.columns]
if missing_base:
    st.error(f"Kolom wajib hilang: {missing_base}")
    st.stop()

# Pastikan mapel yang diperlukan ada di file (jika tidak ada, beri peringatan)
missing_mapel = [m for m in mapel_urut if m not in df.columns]
if missing_mapel:
    st.warning(f"Beberapa mata pelajaran tidak ditemukan di file dan akan diperlakukan kosong: {missing_mapel}")
    # tambahkan kolom kosong agar tidak error saat indexing
    for m in missing_mapel:
        df[m] = np.nan
    # recompute df_kelas (kalau perlu)
    if semua_paralel:
        prefix = str(sel_kelas).strip()[0]
        df_kelas = df[df["Kelas"].astype(str).str.startswith(prefix)].copy()
    else:
        df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()

# Bersihkan & konversi nilai pada semua kolom mapel yang ada di df
for col in [c for c in mapel_semua if c in df.columns]:
    # ubah ke string, ganti koma, hapus karakter non-digit kecuali . dan -, trim
    df[col] = (
        df[col]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
        .str.strip()
    )
    # kosongkan string hasil pembersihan yang menjadi ''
    df.loc[df[col] == "", col] = np.nan
    # konversi ke numeric
    df[col] = pd.to_numeric(df[col], errors="coerce")

# Pastikan ulang df_kelas kolom tersedia
df = df[["Kelas", "NIS", "Nama Siswa"] + [m for m in mapel_urut if m in df.columns]]
# recompute df_kelas again to align with df columns if needed
if semua_paralel:
    prefix = str(sel_kelas).strip()[0]
    df_kelas = df[df["Kelas"].astype(str).str.startswith(prefix)].copy()
else:
    df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()

# Pilih siswa (selalu definisikan, hindari NameError)
siswa_list = df_kelas["Nama Siswa"].astype(str).tolist()
sel_siswa = st.selectbox("Pilih Siswa", ["-- Semua Siswa --"] + siswa_list)

# helper: format skor aman
def format_score(val):
    if pd.isna(val):
        return ""
    try:
        num = float(val)
        # SELALU tampilkan dua angka di belakang koma
        return f"{num:.2f}"
    except Exception:
        # coba convert dari string yang mungkin masih berformat koma
        try:
            num = float(str(val).strip().replace(",", "."))
            return f"{num:.2f}"
        except Exception:
            return str(val)

# Fungsi gambar halaman siswa DENGAN MARGIN
def draw_student_page(c, row, sel_asesmen, sel_tahun, mapel_urut, sel_tgl_ttd):
    width, height = A4

    # === PENGATURAN MARGIN ===
    margin_left   = 30 * mm
    margin_right  = 20 * mm
    margin_top    = 20 * mm
    margin_bottom = 20 * mm

    # Hitung area kerja
    content_width  = width - (margin_left + margin_right)

    # Titik awal Y (di dalam margin atas)
    y = height - margin_top

    # (Optional) logo kiri atas
    logo_path = "assets/logo_kiri.png"
    try:
        logo_w = 30*mm
        logo_h = 30*mm
        # Posisikan logo di margin_left
        x_logo = margin_left - 10*mm
        # Posisi Y disesuaikan agar naik (penyesuaian dari -10*mm menjadi -5*mm)
        y_logo = height - margin_top - (-10*mm) - logo_h
        c.drawImage(logo_path, x_logo, y_logo, width=logo_w, height=logo_h, preserveAspectRatio=True, mask='auto')
    except Exception as e:
        # Menghilangkan error jika gambar tidak ditemukan
        pass

    # KOP sederhana (CentredString tidak dipengaruhi margin kiri/kanan)
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "PEMERINTAH KABUPATEN BANTUL")
    y -= 5*mm
    c.drawCentredString(width/2, y, "DINAS PENDIDIKAN, KEPEMUDAAN, DAN OLAHRAGA")
    y -= 5*mm
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width/2, y, "SMP NEGERI 2 BANGUNTAPAN")
    y -= 1*mm

    # (Optional) aksara jawa
    aksara_path = "assets/aksara_jawa.jpg"
    try:
        aksara_w = 100*mm
        aksara_h = 10*mm
        x_aksara = (width - aksara_w) / 2
        y_aksara = y - aksara_h
        c.drawImage(aksara_path, x_aksara, y_aksara, width=aksara_w, height=aksara_h, preserveAspectRatio=True, mask='auto')
        y = y_aksara - 2*mm
    except Exception as e:
        y -= 4*mm

    c.setFont("Helvetica-Oblique", 10)
    c.drawCentredString(width/2, y, "Jalan Karangsari, Banguntapan, Kabupaten Bantul, Yogyakarta 55198 Telp. 382754")
    y -= 5*mm
    c.setFont("Helvetica", 10)
    c.setFillColor(blue)
    c.drawCentredString(width/2, y, "Website : www.smpn2banguntapan.sch.id Email : smp2banguntapan@yahoo.com")
    c.setFillColor(black)
    y -= 3*mm

    # Garis separator (double)
    c.setLineWidth(1)
    c.line(margin_left, y, width - margin_right, y)
    y -= 1.5*mm
    c.setLineWidth(0.5)
    c.line(margin_left, y, width - margin_right, y)
    y -= 10*mm

    # Judul
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, f"LAPORAN HASIL {sel_asesmen}")
    y -= 6*mm
    c.drawCentredString(width/2, y, f"TAHUN PELAJARAN {sel_tahun}")
    y -= 12*mm

    # Identitas siswa (titik dua rata)
    id_margin_left = margin_left + 10*mm
    label_w = 20*mm
    colon_x = id_margin_left + label_w
    value_x = colon_x + 5
    c.setFont("Helvetica", 12)
    # Nama
    c.drawString(id_margin_left, y, "Nama")
    c.drawString(colon_x, y, ":")
    c.drawString(value_x, y, " " + str(row.get("Nama Siswa", "")))
    y -= 6*mm
    # NIS
    c.drawString(id_margin_left, y, "NIS")
    c.drawString(colon_x, y, ":")
    c.drawString(value_x, y, " " + str(row.get("NIS", "")))
    y -= 6*mm
    # Kelas
    c.drawString(id_margin_left, y, "Kelas")
    c.drawString(colon_x, y, ":")
    c.drawString(value_x, y, " " + str(row.get("Kelas", "")))
    y -= 10*mm

    # Ambil nilai mapel (pastikan numeric)
    nilai_list = []
    for subj in mapel_urut:
        raw = row.get(subj, np.nan)
        if pd.isna(raw):
            nilai_list.append(np.nan)
        else:
            try:
                nilai_list.append(float(raw))
            except Exception:
                try:
                    nilai_list.append(float(str(raw).strip().replace(",", ".")))
                except Exception:
                    nilai_list.append(np.nan)
    nilai_series = pd.Series(nilai_list, index=mapel_urut, dtype="float64")

    jumlah = float(nilai_series.sum(skipna=True))
    rata2 = float(nilai_series.mean(skipna=True)) if nilai_series.count() > 0 else 0.0

    # TABEL
    row_height = 7*mm
    font_size = 11
    col_no_w = 15*mm
    col_mapel_w = 90*mm
    col_nilai_w = 25*mm
    table_width = col_no_w + col_mapel_w + col_nilai_w
    # Posisikan tabel di tengah area konten
    x0 = margin_left + (content_width - table_width) / 2
    y0 = y
    nrows = len(mapel_urut) + 2 + 1  # header + mapel + jumlah + rata2

    # header background (gambar dulu)
    header_y_bottom = y0 - row_height
    c.setFillColor(lightgrey)
    c.rect(x0, header_y_bottom, table_width, row_height, stroke=0, fill=1)
    c.setFillColor(black)

    # gambar grid
    for r in range(nrows+1):
        c.setLineWidth(0.5)
        c.line(x0, y0 - r*row_height, x0 + table_width, y0 - r*row_height)
    c.line(x0, y0, x0, y0 - nrows*row_height)
    c.line(x0 + col_no_w, y0, x0 + col_no_w, y0 - nrows*row_height)
    c.line(x0 + col_no_w + col_mapel_w, y0, x0 + col_no_w + col_mapel_w, y0 - nrows*row_height)
    c.line(x0 + table_width, y0, x0 + table_width, y0 - nrows*row_height)

    # header teks (vertical center correction)
    c.setFont("Helvetica-Bold", font_size)
    header_center = y0 - row_height/2
    adj_y = header_center - (font_size/3.5)
    c.drawCentredString(x0 + col_no_w/2, adj_y, "No")
    c.drawCentredString(x0 + col_no_w + col_mapel_w/2, adj_y, "Mata Pelajaran")
    c.drawCentredString(x0 + col_no_w + col_mapel_w + col_nilai_w/2, adj_y, "Nilai")

    # isi tabel
    c.setFont("Helvetica", font_size)
    y_text = y0 - row_height
    for i, subj in enumerate(mapel_urut, start=1):
        cell_middle = y_text - row_height/2
        adj_y = cell_middle - (font_size/3.5)
        val = nilai_series.get(subj, np.nan)
        val_str = format_score(val)
        c.drawCentredString(x0 + col_no_w/2, adj_y, str(i))
        c.drawString(x0 + col_no_w + 2*mm, adj_y, subj)
        c.drawCentredString(x0 + col_no_w + col_mapel_w + col_nilai_w/2, adj_y, val_str)
        y_text -= row_height

    # Jumlah & Rata-rata (rata kiri teks label)
    cell_middle = y_text - row_height/2
    adj_y = cell_middle - (font_size/3.5)
    c.setFont("Helvetica-Bold", font_size)
    c.drawString(x0 + col_no_w + 2*mm, adj_y, "Jumlah")
    c.drawCentredString(x0 + col_no_w + col_mapel_w + col_nilai_w/2, adj_y, format_score(jumlah))
    y_text -= row_height

    cell_middle = y_text - row_height/2
    adj_y = cell_middle - (font_size/3.5)
    c.drawString(x0 + col_no_w + 2*mm, adj_y, "Rata-rata")
    c.drawCentredString(x0 + col_no_w + col_mapel_w + col_nilai_w/2, adj_y, format_score(rata2))
    y_text -= row_height + 20

    # tanda tangan
    # GUNAKAN TANGGAL PILIHAN DARI STREAMLIT (sel_tgl_ttd)
    ttd_date = sel_tgl_ttd
    bulan_eng = ttd_date.strftime('%B')
    tgl = f"{ttd_date.day} {bulan_id[bulan_eng]} {ttd_date.year}"

    # Posisikan tanda tangan berdasarkan margin_right dan margin_bottom
    x_ttd = width - margin_right - 70*mm
    y_ttd_start = margin_bottom + 67*mm

    c.setFont("Helvetica", 12)
    c.drawString(x_ttd, y_ttd_start, f"Banguntapan, {tgl}")
    y_ttd_start -= 8*mm
    c.drawString(x_ttd, y_ttd_start, "Mengetahui,")
    y_ttd_start -= 5*mm
    c.drawString(x_ttd, y_ttd_start, "Kepala Sekolah,")
    y_ttd_start -= -1*mm
    # Tambah gambar tanda tangan (cek file)
    ttd_path = "assets/ttd_kepsek.jpeg"
    if os.path.exists(ttd_path):
        try:
            c.drawImage(ttd_path, x_ttd, y_ttd_start - 22*mm, width=40*mm, height=20*mm, mask="auto")
        except Exception:
            pass

    # Nama & NIP (tetap ditampilkan)
    y_ttd_after = y_ttd_start - 25*mm
    c.drawString(x_ttd, y_ttd_after, "Alina Fiftiyani Nurjannah, M.Pd.")
    y_ttd_after -= 6*mm
    c.drawString(x_ttd, y_ttd_after, "NIP 198001052009032006")

# PDF generator
def make_pdf_for_student(row, mapel_urut, sel_tgl_ttd):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    # TERUSKAN sel_tgl_ttd KE draw_student_page
    draw_student_page(c, row, sel_asesmen, sel_tahun, mapel_urut, sel_tgl_ttd)
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

def make_pdf_for_class(df_kelas, mapel_urut, sel_tgl_ttd):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    for _, row in df_kelas.iterrows():
        # TERUSKAN sel_tgl_ttd KE draw_student_page
        draw_student_page(c, row, sel_asesmen, sel_tahun, mapel_urut, sel_tgl_ttd)
        c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

# Tambahan: fungsi untuk semua kelas paralel (jika ingin semua kelas di file)
def make_pdf_for_all_classes(df_all, kelas_list_all, mapel_kelas_7_8, mapel_kelas_9, sel_tgl_ttd):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    for kelas in kelas_list_all:
        kelas_upper = str(kelas).upper().strip()
        if kelas_upper.startswith("IX") or kelas_upper.startswith("9"):
            mapel_u = [m for m in mapel_kelas_9 if m in df_all.columns]
        else:
            mapel_u = [m for m in mapel_kelas_7_8 if m in df_all.columns]
        df_sel = df_all[df_all["Kelas"].astype(str) == str(kelas)]
        for _, row in df_sel.iterrows():
            draw_student_page(c, row, sel_asesmen, sel_tahun, mapel_u, sel_tgl_ttd)
            c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

st.markdown("---")
st.subheader("Pilih Siswa & Unduh Laporan")

# Buttons (hanya tampil jika df_kelas tidak kosong)
if df_kelas.empty:
    st.warning("Tidak ada data siswa untuk kelas ini.")
else:
    if sel_siswa != "-- Semua Siswa --":
        row = df_kelas[df_kelas["Nama Siswa"] == sel_siswa].iloc[0]
        # TERUSKAN sel_tgl_ttd
        st.download_button("📄 Download PDF (Per Siswa)",
                           data=make_pdf_for_student(row, mapel_urut, sel_tgl_ttd),
                           file_name=f"Laporan_{row['Nama Siswa']}.pdf",
                           mime="application/pdf")

    # TERUSKAN sel_tgl_ttd
    st.download_button("📄 Download PDF (Per Kelas)",
                       data=make_pdf_for_class(df_kelas, mapel_urut, sel_tgl_ttd),
                       file_name=f"Laporan_{sel_kelas}.pdf",
                       mime="application/pdf")

    # Jika user ingin seluruh kelas paralel sekaligus (tombol tambahan)
    if len(kelas_list) > 1:
        st.download_button(
            "📚 Download PDF (Semua kelas)",
            data=make_pdf_for_all_classes(df, kelas_list, mapel_kelas_7_8, mapel_kelas_9, sel_tgl_ttd),
            file_name=f"Laporan_Semua_Kelas_{sel_tahun}.pdf",
            mime="application/pdf"
        )
