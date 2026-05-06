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
mapel_kelas_7_8 = [
    "Pend. Agama dan Budi Pekerti", "Pendidikan Pancasila", "Bahasa Indonesia",
    "Matematika", "Ilmu Pengetahuan Alam", "Ilmu Pengetahuan Sosial",
    "Bahasa Inggris", "PJOK", "Informatika", "Seni Budaya", "Bahasa Jawa"
]

mapel_kelas_9 = [
    "Pend. Agama dan Budi Pekerti", "Pendidikan Pancasila", "Bahasa Indonesia",
    "Matematika", "Ilmu Pengetahuan Alam", "Ilmu Pengetahuan Sosial",
    "Bahasa Inggris", "PJOK", "Informatika", "Prakarya", "Bahasa Jawa"
]

mapel_semua = sorted(set(mapel_kelas_7_8) | set(mapel_kelas_9))

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
sel_tgl_ttd = st.date_input("Tanggal Tanda Tangan", datetime.now(), format="DD/MM/YYYY")

st.markdown("---")

# === Template Excel ===
def generate_template():
    cols = ["Kelas", "NIS", "Nama Siswa"] + mapel_semua + ["Literasi Budaya"]
    df_template = pd.DataFrame(columns=cols)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_template.to_excel(writer, index=False, sheet_name="Nilai")
    buffer.seek(0)
    return buffer

st.download_button("📥 Download Template Excel", data=generate_template(), file_name="Template_Nilai.xlsx")

uploaded = st.file_uploader("Unggah file Excel (.xlsx)", type=["xlsx"])
if not uploaded:
    st.stop()

df = pd.read_excel(uploaded, engine="openpyxl")
df.columns = df.columns.str.strip()
kelas_list = sorted(df["Kelas"].astype(str).unique())
sel_kelas = st.selectbox("Pilih Kelas", kelas_list)
semua_paralel = st.checkbox("Cetak semua kelas ?", value=False)

if semua_paralel:
    prefix = str(sel_kelas).strip()[0]
    df_kelas = df[df["Kelas"].astype(str).str.startswith(prefix)].copy()
else:
    df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()

kelas_upper = str(sel_kelas).upper().strip()
if kelas_upper.startswith("IX") or kelas_upper.startswith("9"):
    mapel_urut = [m for m in mapel_kelas_9 if m in df.columns]
else:
    mapel_urut = [m for m in mapel_kelas_7_8 if m in df.columns]

for col in [c for c in mapel_semua if c in df.columns]:
    df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")

siswa_list = ["-- Semua Siswa --"] + df_kelas["Nama Siswa"].astype(str).tolist()
sel_siswa = st.selectbox("Pilih Siswa", siswa_list)

def format_score(val):
    if pd.isna(val): return ""
    return f"{float(val):.2f}"

def draw_student_page(c, row, sel_asesmen, sel_tahun, mapel_urut, sel_tgl_ttd):
    width, height = A4
    margin_left, margin_right, margin_top, margin_bottom = 30*mm, 20*mm, 20*mm, 20*mm
    content_width = width - (margin_left + margin_right)
    y = height - margin_top

    # --- KOP & LOGO ---
    logo_path = "assets/logo_kiri.png"
    if os.path.exists(logo_path):
        c.drawImage(logo_path, margin_left - 5*mm, y - 21*mm, width=25*mm, height=25*mm, preserveAspectRatio=True, mask='auto')

    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "PEMERINTAH KABUPATEN BANTUL")
    y -= 5*mm
    c.drawCentredString(width/2, y, "DINAS PENDIDIKAN, KEPEMUDAAN, DAN OLAHRAGA")
    y -= 5*mm
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width/2, y, "SMP NEGERI 2 BANGUNTAPAN")
    
    # MODIFIKASI: Logo Aksara Jawa diturunkan satu baris (dari y-9 menjadi y-13) 
    # agar tidak menimpa teks nama sekolah di atasnya
    aksara_path = "assets/aksara_jawa.jpg"
    if os.path.exists(aksara_path):
        c.drawImage(aksara_path, (width-110*mm)/2, y - 13*mm, width=110*mm, height=11*mm, preserveAspectRatio=True, mask='auto')
        y -= 15*mm # Jarak ke baris alamat di bawah logo
    else:
        y -= 6*mm

    c.setFont("Helvetica-Oblique", 10)
    c.drawCentredString(width/2, y, "Jalan Karangsari, Banguntapan, Kabupaten Bantul, Yogyakarta 55198 Telp. 382754")
    y -= 4*mm
    c.setFont("Helvetica", 10)
    c.setFillColor(blue)
    c.drawCentredString(width/2, y, "Website : www.smpn2banguntapan.sch.id Email : smp2banguntapan@yahoo.com")
    c.setFillColor(black)
    y -= 3*mm
    c.setLineWidth(1); c.line(margin_left, y, width - margin_right, y)
    y -= 1*mm; c.setLineWidth(0.5); c.line(margin_left, y, width - margin_right, y)
    
    # Judul & Identitas
    y -= 8*mm
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, f"LAPORAN HASIL {sel_asesmen}")
    y -= 5*mm
    c.drawCentredString(width/2, y, f"TAHUN PELAJARAN {sel_tahun}")
    
    y -= 10*mm
    c.setFont("Helvetica", 11)
    idx = margin_left + 10*mm
    for label, key in [("Nama", "Nama Siswa"), ("NIS", "NIS"), ("Kelas", "Kelas")]:
        c.drawString(idx, y, label)
        c.drawString(idx + 25*mm, y, f": {row.get(key, '')}")
        y -= 6*mm

    # --- TABEL ---
    y -= 4*mm 
    row_h = 7*mm
    c_no, c_mp, c_ni = 12*mm, 95*mm, 23*mm
    tw = c_no + c_mp + c_ni
    x0 = margin_left + (content_width - tw) / 2
    
    c.setFillColor(lightgrey); c.rect(x0, y - row_h, tw, row_h, fill=1); c.setFillColor(black)
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(x0 + c_no/2, y - 5*mm, "No")
    c.drawCentredString(x0 + c_no + c_mp/2, y - 5*mm, "Mata Pelajaran")
    c.drawCentredString(x0 + c_no + c_mp + c_ni/2, y - 5*mm, "Nilai")
    
    y_table_top = y
    y_text = y - row_h
    c.setFont("Helvetica", 11)
    vals = []
    for i, m in enumerate(mapel_urut, 1):
        v = row.get(m, np.nan)
        c.drawCentredString(x0 + c_no/2, y_text - 5*mm, str(i))
        c.drawString(x0 + c_no + 2*mm, y_text - 5*mm, m)
        c.drawCentredString(x0 + c_no + c_mp + c_ni/2, y_text - 5*mm, format_score(v))
        if not pd.isna(v): vals.append(float(v))
        y_text -= row_h
        c.line(x0, y_text, x0 + tw, y_text)

    for label, res in [("Jumlah", sum(vals)), ("Rata-rata", sum(vals)/len(vals) if vals else 0)]:
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x0 + c_no + 2*mm, y_text - 5*mm, label)
        c.drawCentredString(x0 + c_no + c_mp + c_ni/2, y_text - 5*mm, format_score(res))
        y_text -= row_h
        c.line(x0, y_text, x0 + tw, y_text)

    c.line(x0, y_table_top, x0, y_text)
    c.line(x0 + c_no, y_table_top, x0 + c_no, y_text)
    c.line(x0 + c_no + c_mp, y_table_top, x0 + c_no + c_mp, y_text)
    c.line(x0 + tw, y_table_top, x0 + tw, y_text)
    c.line(x0, y_table_top, x0+tw, y_table_top)

    # --- KETERANGAN ---
    y_text -= 8*mm 
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(black)
    c.drawString(x0, y_text, "Ket :")
    y_text -= 5*mm
    lit_bud = str(row.get("Literasi Budaya", "Belum")).strip().capitalize()
    c.drawString(x0 + 12*mm, y_text, f"Tugas Laporan Literasi Budaya : {lit_bud}")

    # === TANDA TANGAN ===
    y_ttd = y_text - 13*mm 
    x_ttd = width - margin_right - 65*mm
    tgl_str = f"{sel_tgl_ttd.day} {bulan_id[sel_tgl_ttd.strftime('%B')]} {sel_tgl_ttd.year}"

    c.setFont("Helvetica", 12)
    c.drawString(x_ttd, y_ttd, f"Banguntapan, {tgl_str}")
    y_ttd -= 10*mm 
    c.drawString(x_ttd, y_ttd, "Mengetahui,")
    y_ttd -= 6*mm
    c.drawString(x_ttd, y_ttd, "Kepala Sekolah,")
    
    ttd_img = "assets/ttd_kepsek.jpeg"
    if os.path.exists(ttd_img):
        c.drawImage(ttd_img, x_ttd, y_ttd - 22*mm, width=35*mm, height=18*mm, mask='auto')

    c.setFont("Helvetica-Bold", 12)
    c.drawString(x_ttd, y_ttd - 30*mm, "Alina Fiftiyani Nurjannah, M.Pd.")
    c.setFont("Helvetica", 12)
    c.drawString(x_ttd, y_ttd - 36*mm, "NIP 198001052009032006")

def make_pdf(df_source, mapel_u, tgl):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    for _, r in df_source.iterrows():
        draw_student_page(c, r, sel_asesmen, sel_tahun, mapel_u, tgl)
        c.showPage()
    c.save(); buf.seek(0)
    return buf

if not df_kelas.empty:
    if sel_siswa != "-- Semua Siswa --":
        row_s = df_kelas[df_kelas["Nama Siswa"] == sel_siswa]
        st.download_button("📄 Download PDF Siswa", make_pdf(row_s, mapel_urut, sel_tgl_ttd), f"Laporan_{sel_siswa}.pdf")
    st.download_button("📄 Download PDF Satu Kelas", make_pdf(df_kelas, mapel_urut, sel_tgl_ttd), f"Laporan_{sel_kelas}.pdf")
