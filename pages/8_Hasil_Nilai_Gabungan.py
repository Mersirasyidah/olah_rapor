import io
import os
import numpy as np
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, lightgrey, blue, white
from datetime import datetime

# --- Konfigurasi Mapel & Semester ---
MAPEL_UTAMA = ["Bahasa Indonesia", "Matematika", "Bahasa Inggris", "IPA"]
# Tetap menggunakan kunci S1-S5 agar sinkron dengan data Excel
SEMESTER_LIST = ["S1", "S2", "S3", "S4", "S5"]

bulan_id = {
    "January": "Januari", "February": "Februari", "March": "Maret",
    "April": "April", "May": "Mei", "June": "Juni",
    "July": "Juli", "August": "Agustus", "September": "September",
    "October": "Oktober", "November": "November", "December": "Desember"
}

st.set_page_config(page_title="Generator Laporan Nilai", layout="centered")
st.header("Laporan Hasil Nilai Gabungan TKA/D")

# --- 1. Fungsi Template Excel ---
def generate_template():
    cols = ["Kelas", "NIS", "Nama Siswa"]
    for m in MAPEL_UTAMA:
        for s in SEMESTER_LIST:
            cols.append(f"{m}_{s}")
    for m in MAPEL_UTAMA:
        cols.append(f"{m}_TKAD")
    
    df_template = pd.DataFrame(columns=cols)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_template.to_excel(writer, index=False, sheet_name="Sheet1")
    buffer.seek(0)
    return buffer

st.download_button("📥 Download Template Excel", data=generate_template(), file_name="Template_Laporan.xlsx")

uploaded = st.file_uploader("Unggah file Excel Nilai", type=["xlsx"])

if not uploaded:
    st.info("Silakan unggah file Excel.")
    st.stop()

df = pd.read_excel(uploaded)

# Konversi angka & Proteksi Kolom
for m in MAPEL_UTAMA:
    for s in SEMESTER_LIST:
        col = f"{m}_{s}"
        if col not in df.columns: df[col] = 0
    col_t = f"{m}_TKAD"
    if col_t not in df.columns: df[col_t] = 0

for col in df.columns:
    if col not in ["Kelas", "NIS", "Nama Siswa"]:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce").fillna(0)

# --- 2. Fungsi Gambar PDF ---
def draw_kwarto_page(c, row, sel_tahun, tgl_ttd):
    width, height = LETTER
    margin_left = 25 * mm
    margin_right = 20 * mm
    y = height - 18 * mm 
    
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
    
    aksara_path = "assets/aksara_jawa.jpg"
    if os.path.exists(aksara_path):
        c.drawImage(aksara_path, (width-135*mm)/2, y - 13*mm, width=135*mm, height=11*mm, preserveAspectRatio=True, mask='auto')
        y -= 15*mm 
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
    
    # --- JUDUL & IDENTITAS ---
    y -= 10*mm
    c.setFont("Times-Bold", 14)
    c.drawCentredString(width/2, y, "LAPORAN HASIL NILAI GABUNGAN")
    y -= 5*mm
    c.setFont("Times-Bold", 12)
    c.drawCentredString(width/2, y, f"TAHUN PELAJARAN {sel_tahun}")
    
    y -= 12*mm 
    c.setFont("Times-Roman", 12)
    label_x = margin_left + 10*mm
    value_x = label_x + 40*mm
    
    identitas = [
        ("Nama Siswa", f": {row['Nama Siswa']}"),
        ("NIS", f": {row['NIS']}"),
        ("Kelas", f": {row['Kelas']}")
    ]
    for label, val in identitas:
        c.drawString(label_x, y, label)
        c.drawString(value_x, y, val)
        y -= 7*mm
    
    y -= 2*mm 

    # --- TABEL HEADER ---
    col_w = [10*mm, 50*mm, 12*mm, 12*mm, 12*mm, 12*mm, 12*mm, 27*mm, 27*mm]
    h_header = 18*mm 

    c.setLineWidth(0.5)
    c.setFillColor(lightgrey)
    c.rect(margin_left, y - h_header, sum(col_w), h_header, fill=1)
    c.setFillColor(black)
    c.setFont("Times-Bold", 12)

    # 1. No
    c.rect(margin_left, y - h_header, col_w[0], h_header, stroke=1)
    c.drawCentredString(margin_left + col_w[0]/2, y - (h_header/2) - 1*mm, "No")

    # 2. Mata Pelajaran
    c.rect(margin_left + col_w[0], y - h_header, col_w[1], h_header, stroke=1)
    c.drawCentredString(margin_left + col_w[0] + col_w[1]/2, y - (h_header/2) - 1*mm, "Mata Pelajaran")

    # 3. Nilai Rapor
    rapor_w = sum(col_w[2:7])
    c.rect(margin_left + sum(col_w[:2]), y - 8*mm, rapor_w, 8*mm, stroke=1)
    c.drawCentredString(margin_left + sum(col_w[:2]) + rapor_w/2, y - 5.5*mm, "Nilai Rapor Sem 1-5")

    # 4. Sub-header (Sem-1 dst)
    c.setFont("Times-Bold", 9)
    sub_x = margin_left + sum(col_w[:2])
    for i in range(1, 6):
        c.rect(sub_x, y - h_header, 12*mm, 10*mm, stroke=1)
        c.drawCentredString(sub_x + 6*mm, y - 14*mm, f"Sem-{i}")
        sub_x += 12*mm

    # 5. Rerata Sem 1-5 (TEKS DIBUNGKUS KE BAWAH)
    c.setFont("Times-Bold", 11) # Ukuran sedikit dikecilkan agar pas
    c.rect(margin_left + sum(col_w[:7]), y - h_header, col_w[7], h_header, stroke=1)
    # Gambar baris pertama teks
    c.drawCentredString(margin_left + sum(col_w[:7]) + col_w[7]/2, y - 7*mm, "Rerata")
    # Gambar baris kedua teks
    c.drawCentredString(margin_left + sum(col_w[:7]) + col_w[7]/2, y - 12*mm, "Sem 1-5")

    # 6. Nilai TKA/D
    c.setFont("Times-Bold", 12)
    c.rect(margin_left + sum(col_w[:8]), y - h_header, col_w[8], h_header, stroke=1)
    c.drawCentredString(margin_left + sum(col_w[:8]) + col_w[8]/2, y - (h_header/2) - 1*mm, "Nilai TKA/D")

    y -= h_header
    
    # --- ISI DATA (Logika Tetap Sama) ---
    total_rata_s15 = 0
    total_tkad = 0
    for idx, m in enumerate(MAPEL_UTAMA, 1):
        c.setFont("Times-Roman", 11)
        vals_sem = [row[f"{m}_{s}"] for s in SEMESTER_LIST]
        rata = sum(vals_sem) / 5
        tkad = row[f"{m}_TKAD"]
        total_rata_s15 += rata
        total_tkad += tkad
        
        curr_x = margin_left
        line_data = [str(idx), m, f"{vals_sem[0]:.0f}", f"{vals_sem[1]:.0f}", f"{vals_sem[2]:.0f}", f"{vals_sem[3]:.0f}", f"{vals_sem[4]:.0f}", f"{rata:.2f}", f"{tkad:.2f}"]
        for i, val in enumerate(line_data):
            c.rect(curr_x, y - 10*mm, col_w[i], 10*mm, stroke=1)
            if i == 1: c.drawString(curr_x + 2*mm, y - 6.5*mm, val)
            else: c.drawCentredString(curr_x + col_w[i]/2, y - 6.5*mm, val)
            curr_x += col_w[i]
        y -= 10*mm
        
    # --- FOOTER TABEL ---
    c.setFont("Times-Bold", 11)
    c.rect(margin_left, y - 10*mm, sum(col_w[:-2]), 10*mm, stroke=1)
    c.drawString(margin_left + 20*mm, y - 6.5*mm, "JUMLAH")
    c.rect(margin_left + sum(col_w[:-2]), y - 10*mm, col_w[-2], 10*mm, stroke=1); c.drawCentredString(margin_left + sum(col_w[:-2]) + col_w[-2]/2, y - 6.5*mm, f"{total_rata_s15:.2f}")
    c.rect(margin_left + sum(col_w[:-1]), y - 10*mm, col_w[-1], 10*mm, stroke=1); c.drawCentredString(margin_left + sum(col_w[:-1]) + col_w[-1]/2, y - 6.5*mm, f"{total_tkad:.2f}")
    y -= 10*mm

    nilai_gabungan = (total_rata_s15 * 0.4) + (total_tkad * 0.6)
    c.setFillColor(lightgrey); c.rect(margin_left, y - 12*mm, sum(col_w[:-1]), 12*mm, fill=1)
    c.setFillColor(black); c.setFont("Times-Bold", 12); c.drawCentredString(margin_left + sum(col_w[:-1])/2, y - 7.5*mm, "NILAI GABUNGAN")
    c.rect(margin_left + sum(col_w[:-1]), y - 12*mm, col_w[-1], 12*mm, stroke=1); c.setFont("Times-Bold", 14); c.drawCentredString(margin_left + sum(col_w[:-1]) + col_w[-1]/2, y - 7.5*mm, f"{nilai_gabungan:.2f}")
    
    # JARAK KE KETERANGAN
    y -= 22*mm 
    c.setFont("Times-Bold", 12); c.drawString(margin_left, y, "Ket :")
    y -= 6*mm
    c.setFont("Times-BoldItalic", 11); c.drawString(margin_left + 5*mm, y, "Rumus Nilai Gabungan = ((Nilai TKA + TKAD) x 60%) + (Jumlah Rerata Nilai Rapor Semester 1-5 x 40%)")
    
    # --- TANDA TANGAN ---
    y -= 15*mm 
    xttd = width - margin_right - 65*mm
    tgl_str = f"{tgl_ttd.day} {bulan_id[tgl_ttd.strftime('%B')]} {tgl_ttd.year}"
    
    c.setFont("Times-Roman", 12)
    c.drawString(xttd, y, f"Banguntapan, {tgl_str}")
    y -= 6*mm
    c.drawString(xttd, y, "Kepala Sekolah,")
    
    ttd_path = "assets/ttd_kepsek.jpeg" 
    if os.path.exists(ttd_path):
        c.drawImage(ttd_path, xttd + 5*mm, y - 22*mm, width=40*mm, height=20*mm, preserveAspectRatio=True, mask='auto')
    
    y -= 25*mm
    c.setFont("Times-Bold", 12); c.drawString(xttd, y, "Alina Fiftiyani Nurjannah, M.Pd.")
    y -= 5*mm
    c.setFont("Times-Roman", 12); c.drawString(xttd, y, "NIP 198001052009032006")

# --- Streamlit Execution ---
col1, col2 = st.columns(2)
with col1: sel_tahun = st.selectbox("Tahun Pelajaran", ["2024/2025", "2025/2026"])
with col2: sel_tgl = st.date_input("Tanggal TTD", datetime.now())

if not df.empty:
    sel_kelas = st.selectbox("Pilih Kelas", sorted(df["Kelas"].unique()))
    if st.button(f"🚀 Generate PDF Kelas {sel_kelas}"):
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=LETTER)
        for _, row in df[df["Kelas"] == sel_kelas].iterrows():
            draw_kwarto_page(c, row, sel_tahun, sel_tgl)
            c.showPage()
        c.save(); buffer.seek(0)
        st.download_button("📄 Unduh PDF", data=buffer, file_name=f"Laporan_{sel_kelas}.pdf", mime="application/pdf")
