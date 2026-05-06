import io
import numpy as np
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import blue, black, lightgrey
from reportlab.pdfgen import canvas
from datetime import datetime
import os

# ===============================================
# === KONFIGURASI ===
# ===============================================
mapel_tetap = ["Bahasa Indonesia", "Matematika", "Bahasa Inggris", "Ilmu Pengetahuan Alam"]
mapel_urut = mapel_tetap

bulan_id = {
    "January": "Januari", "February": "Februari", "March": "Maret",
    "April": "April", "May": "Mei", "June": "Juni",
    "July": "Juli", "August": "Agustus", "September": "September",
    "October": "Oktober", "November": "November", "December": "Desember"
}

st.header("Laporan Hasil Persiapan dan Pemantapan")

# --- Input Pilihan ---
asesmen_opsi = [
    "TKA/TKAD Dikpora Kab Bantul 1", "TKA/TKAD MKKS SMP Kab Bantul 1",
    "TKA/TKAD MKKS SMP Kab Bantul 2", "TKA/TKAD Forum MKKS SMP D.I.Yogyakarta",
    "TKA/TKAD Dikpora Kab Bantul 2"
]
sel_asesmen = st.selectbox("Pilih Jenis Asesmen", asesmen_opsi)
tahun_opsi = [f"{th}/{th+1}" for th in range(2025, 2036)]
sel_tahun = st.selectbox("Pilih Tahun Pelajaran", tahun_opsi, index=0)

tgl_kegiatan_opsi = [
    "Tanggal 20 - 23 Oktober 2025", "Tanggal 3 - 6 November 2025",
    "Tanggal 19 - 22 Januari 2026", "Tanggal 2 - 5 Februari 2026", "Tanggal 9 - 12 Maret 2026"
]
sel_tgl_kegiatan = st.selectbox("Pilih Tanggal Kegiatan", tgl_kegiatan_opsi)
sel_tgl_ttd = st.date_input("Tanggal Tanda Tangan", datetime.now())

# === Upload Data & Hitung Rank ===
uploaded = st.file_uploader("Unggah file Excel", type=["xlsx"])
if not uploaded: st.stop()

df = pd.read_excel(uploaded).fillna(np.nan)
df.columns = df.columns.str.strip()

# Perhitungan Ranking Per Sesi
for i in range(1, 6):
    cols_sesi = [f"{m}_TKAD{i}" for m in mapel_tetap]
    available = [c for c in cols_sesi if c in df.columns]
    if available:
        df[f"Total_Sesi_{i}"] = df[available].sum(axis=1, skipna=True)
        mask = df[available].notna().any(axis=1)
        df.loc[mask, f"Peringkat_TKAD{i}"] = df[mask].groupby("Kelas")[f"Total_Sesi_{i}"].rank(ascending=False, method="min")
    else:
        df[f"Peringkat_TKAD{i}"] = np.nan

kelas_list = sorted(df["Kelas"].astype(str).unique())
sel_kelas = st.selectbox("Pilih Kelas", kelas_list)
df_kelas = df[df["Kelas"].astype(str) == str(sel_kelas)].copy()
siswa_list = df_kelas["Nama Siswa"].astype(str).tolist()
sel_siswa = st.selectbox("Pilih Siswa", ["-- Semua Siswa --"] + siswa_list)

# ===============================================
# === PDF GENERATOR ===
# ===============================================
def format_val(val):
    if pd.isna(val) or val == "": return ""
    try:
        num = float(val)
        return f"{num:.2f}" if num % 1 != 0 else str(int(num))
    except: return str(val)

def draw_student_page(c, row):
    width, height = A4
    ml, mr, mt = 30*mm, 15*mm, 25*mm
    cw = width - ml - mr
    y = height - mt

    # --- 1. KOP SURAT ---
    logo_path = "assets/logo_kiri.png"
    if os.path.exists(logo_path):
        c.drawImage(logo_path, ml - 10*mm, y - 18*mm, width=28*mm, height=28*mm, preserveAspectRatio=True, mask='auto')
    
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, y, "PEMERINTAH KABUPATEN BANTUL")
    y -= 5*mm
    c.drawCentredString(width/2, y, "DINAS PENDIDIKAN, KEPEMUDAAN, DAN OLAHRAGA")
    y -= 5*mm
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width/2, y, "SMP NEGERI 2 BANGUNTAPAN")
    y -= 1*mm

    # Aksara Jawa (Tetap Ukuran Besar)
    aksara_path = "assets/aksara_jawa.jpg"
    if os.path.exists(aksara_path):
        c.drawImage(aksara_path, (width-125*mm)/2, y - 13*mm, width=125*mm, height=12*mm, preserveAspectRatio=True, mask='auto')
        y -= 14*mm
    else:
        y -= 4*mm

    c.setFont("Helvetica-Oblique", 10)
    c.drawCentredString(width/2, y, "Jalan Karangsari, Banguntapan, Bantul, Yogyakarta 55198 Telp. 382754")
    y -= 5*mm
    c.setFont("Helvetica", 10); c.setFillColor(blue)
    c.drawCentredString(width/2, y, "Website : www.smpn2banguntapan.sch.id Email : smp2banguntapan@yahoo.com")
    c.setFillColor(black); y -= 3*mm
    
    c.setLineWidth(1.2); c.line(ml, y, width-mr, y)
    y -= 10*mm # KEMBALI KE SEMULA (Sebelumnya 7mm)

    # --- 2. JUDUL ---
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(width/2, y, "LAPORAN HASIL PERSIAPAN DAN PEMANTAPAN")
    y -= 6*mm
    c.drawCentredString(width/2, y, sel_asesmen.upper())
    y -= 6*mm
    c.drawCentredString(width/2, y, f"TAHUN PELAJARAN {sel_tahun}")
    y -= 6*mm
    c.drawCentredString(width/2, y, sel_tgl_kegiatan)
    y -= 12*mm # KEMBALI KE SEMULA (Sebelumnya 8mm)

    # --- 3. IDENTITAS ---
    c.setFont("Helvetica-Bold", 10)
    for lbl, k in [("Nama", "Nama Siswa"), ("NIS", "NIS"), ("Kelas", "Kelas")]:
        c.drawString(ml+10*mm, y, lbl)
        c.drawString(ml+35*mm, y, f": {row.get(k, '')}")
        y -= 6*mm # KEMBALI KE SEMULA (Sebelumnya 5mm)
    y -= 5*mm # KEMBALI KE SEMULA (Sebelumnya 3mm)

    # --- 4. TABEL ---
    row_h, col_no, col_m, col_s = 7*mm, 12*mm, 65*mm, 17*mm
    tw = col_no + col_m + (col_s * 5)
    xs = ml + (cw - tw)/2

    # Background Header
    c.setFillColor(lightgrey); c.rect(xs, y-2*row_h, tw, 2*row_h, fill=1); c.setFillColor(black)
    
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(xs+col_no/2, y-row_h-2*mm, "No")
    c.drawCentredString(xs+col_no+col_m/2, y-row_h-2*mm, "Mata Pelajaran")
    c.drawCentredString(xs+col_no+col_m+(col_s*5)/2, y-row_h/2-2*mm, "Nilai TKA/TKAD")
    for i in range(5):
        c.drawCentredString(xs+col_no+col_m+(i*col_s)+col_s/2, y-1.5*row_h-2*mm, str(i+1))

    # Data
    y_row = y - 2*row_h
    c.setFont("Helvetica", 10)
    sums = [[] for _ in range(5)]
    for i, m in enumerate(mapel_urut, 1):
        c.drawCentredString(xs+col_no/2, y_row-row_h/2-2*mm, str(i))
        c.drawString(xs+col_no+2*mm, y_row-row_h/2-2*mm, m)
        for j in range(5):
            val = row.get(f"{m}_TKAD{j+1}", np.nan)
            c.drawCentredString(xs+col_no+col_m+(j*col_s)+col_s/2, y_row-row_h/2-2*mm, format_val(val))
            if not pd.isna(val): sums[j].append(float(val))
        y_row -= row_h

    # Footer
    c.setFont("Helvetica-Bold", 10)
    footers = [
        ("Jumlah", [sum(x) if x else np.nan for x in sums]),
        ("Rata-rata", [sum(x)/len(x) if x else np.nan for x in sums]),
        ("Peringkat Kelas ke -", [row.get(f"Peringkat_TKAD{j+1}", np.nan) for j in range(5)])
    ]
    for lbl, data in footers:
        c.drawString(xs+col_no+2*mm, y_row-row_h/2-2*mm, lbl)
        for j, v in enumerate(data):
            c.drawCentredString(xs+col_no+col_m+(j*col_s)+col_s/2, y_row-row_h/2-2*mm, format_val(v))
        y_row -= row_h

    # --- GRID TABEL (Garis Tengah Header Tetap Muncul) ---
    c.setLineWidth(0.5)
    num_grid_rows = len(mapel_urut) + 2 + 3
    for i in range(num_grid_rows + 1):
        if i == 1:
            c.line(xs+col_no+col_m, y-row_h, xs+tw, y-row_h)
        else:
            c.line(xs, y - i*row_h, xs + tw, y - i*row_h)
    
    c.line(xs, y, xs, y_row) 
    c.line(xs+col_no, y, xs+col_no, y_row) 
    c.line(xs+col_no+col_m, y, xs+col_no+col_m, y_row) 
    for i in range(1, 6):
        c.line(xs+col_no+col_m+i*col_s, y-row_h, xs+col_no+col_m+i*col_s, y_row)
    c.line(xs+tw, y, xs+tw, y_row) 

    # --- 5. KETERANGAN ---
    y_ket = y_row - 8*mm
    c.setFont("Helvetica-Bold", 9); c.drawString(ml, y_ket, "Keterangan:")
    c.setFont("Helvetica", 8)
    for i, item in enumerate(asesmen_opsi, 1):
        y_ket -= 4*mm
        tgl_item = tgl_kegiatan_opsi[i-1] if i-1 < len(tgl_kegiatan_opsi) else "-"
        c.drawString(ml+5*mm, y_ket, f"{i}. {item} ({tgl_item})")

    # --- 6. TANDA TANGAN & NIP ---
    y_sign = y_ket - 10*mm
    tgl_str = f"{sel_tgl_ttd.day} {bulan_id[sel_tgl_ttd.strftime('%B')]} {sel_tgl_ttd.year}"
    x_ttd = width - mr - 65*mm
    
    c.setFont("Helvetica", 10)
    c.drawString(x_ttd, y_sign, f"Banguntapan, {tgl_str}")
    y_sign -= 5*mm
    c.drawString(x_ttd, y_sign, "Mengetahui,")
    y_sign -= 5*mm
    c.drawString(x_ttd, y_sign, "Kepala Sekolah,")
    
    # Gambar Tanda Tangan
    ttd_path = "assets/ttd_kepsek.jpeg"
    if os.path.exists(ttd_path):
        c.drawImage(ttd_path, x_ttd, y_sign - 18*mm, width=35*mm, height=15*mm, mask='auto')
    
    # Nama & NIP
    y_name = y_sign - 25*mm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_ttd, y_name, "Alina Fiftiyani Nurjannah, M.Pd.")
    y_name -= 5*mm
    c.setFont("Helvetica", 10)
    c.drawString(x_ttd, y_name, "NIP 198001052009032006")

def make_pdf(data_rows):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    for _, r in data_rows.iterrows():
        draw_student_page(c, r)
        c.showPage()
    c.save(); buf.seek(0)
    return buf

# Tombol Unduh
col1, col2 = st.columns(2)
with col1:
    if sel_siswa != "-- Semua Siswa --":
        row_sel = df_kelas[df_kelas["Nama Siswa"] == sel_siswa]
        st.download_button("📄 PDF Per Siswa", make_pdf(row_sel), f"Laporan_{sel_siswa}.pdf")
with col2:
    st.download_button("📄 PDF Satu Kelas", make_pdf(df_kelas), f"Laporan_Kelas_{sel_kelas}.pdf")
