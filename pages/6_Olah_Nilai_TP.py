import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np

# --- KONFIGURASI AWAL ---
st.set_page_config(layout="wide", page_title="Editor Nilai Rapor")

# Variabel Global
COMMON_SUBJECTS = ["Matematika", "Bahasa Inggris", "IPA", "IPS", "Bahasa Indonesia", "Seni Budaya", "P.Pancasila", "Pendidikan Agama", "PJOK", "Informatika"]
YEAR_OPTIONS = ["2025/2026", "2026/2027", "2027/2028"]
KKM = 80
TP_COLS = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']

# --- FUNGSI LOGIKA (PYTHON) ---
def apply_logic(df):
    """Menghitung Nilai Rapor dan Status TP (R minimal 1)"""
    df = df.copy()
    
    # 1. Pastikan kolom numerik
    for col in TP_COLS + ['PTS', 'SAS']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 2. Hitung Status TK (T/R) dengan syarat R minimal 1 jika semua >= 80
    def check_tk(row):
        # Ambil nilai TP yang tidak 0
        active_tps = [tp for tp in TP_COLS if row[tp] > 0]
        
        # Default status berdasarkan KKM
        for tp in TP_COLS:
            row[f'TK_{tp}'] = "" if row[tp] <= 0 else ("T" if row[tp] >= KKM else "R")
        
        # Validasi: Jika semua tuntas (>=80), ambil 1 yang terkecil jadi R
        if active_tps and all(row[tp] >= KKM for tp in active_tps):
            min_val = min([row[tp] for tp in active_tps])
            # Cari TP pertama yang nilainya sama dengan nilai terkecil
            for tp in active_tps:
                if row[tp] == min_val:
                    row[f'TK_{tp}'] = "R"
                    break
        return row

    df = df.apply(check_tk, axis=1)
    
    # 3. Hitung Rata-rata & NR (Sederhana)
    df['NR'] = df[TP_COLS + ['PTS', 'SAS']].mean(axis=1).round(0).astype(int)
    return df

# --- TAMPILAN UTAMA ---
st.title("Sistem Input Nilai Rapor")

# --- BAGIAN PENGATURAN (IDENTITAS) ---
# Saya pindahkan ke sidebar agar selalu terlihat
with st.sidebar:
    st.header("📌 Identitas Rapor")
    mapel = st.selectbox("Mata Pelajaran", COMMON_SUBJECTS)
    tahun = st.selectbox("Tahun Pelajaran", YEAR_OPTIONS)
    semester = st.radio("Semester", ["Ganjil", "Genap"])
    guru = st.text_input("Nama Guru")
    
    st.divider()
    uploaded_file = st.file_uploader("Upload Daftar Siswa (CSV)", type="csv")

# --- DATA LOADING ---
if uploaded_file:
    df_raw = pd.read_csv(uploaded_file)
else:
    # Dummy data jika tidak ada file
    df_raw = pd.DataFrame({
        'NIS': ['1001', '1002', '1003'],
        'Nama': ['Andi', 'Budi', 'Caca'],
        'Kelas': ['7A', '7A', '7A']
    })

# Inisialisasi kolom nilai jika belum ada
for c in TP_COLS + ['PTS', 'SAS']:
    if c not in df_raw.columns:
        df_raw[c] = 0.0

# --- EDITOR NILAI ---
st.subheader(f"Edit Nilai: {mapel} ({semester})")
edited_df = st.data_editor(df_raw, hide_index=True, use_container_width=True)

# Tombol Proses & Download
if st.button("🔥 Proses & Siapkan Download"):
    processed_df = apply_logic(edited_df)
    
    # Tampilkan preview singkat
    st.success("Logika berhasil diterapkan! (R minimal 1 sudah dicek)")
    st.dataframe(processed_df[['Nama'] + [f'TK_{tp}' for tp in TP_COLS] + ['NR']].head())

    # --- EXCEL EXPORT ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Nilai")
        
        # Format
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
        locked_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'border': 1})
        
        # Tulis Header Identitas di Excel
        worksheet.write(0, 0, f"Mapel: {mapel}")
        worksheet.write(1, 0, f"Tahun: {tahun}")
        worksheet.write(2, 0, f"Guru: {guru}")
        
        # Tulis Data
        cols = processed_df.columns.tolist()
        for i, col in enumerate(cols):
            worksheet.write(4, i, col, header_fmt)
            for r_idx, val in enumerate(processed_df[col]):
                worksheet.write(5 + r_idx, i, val)

    st.download_button(
        label="📥 Download Excel",
        data=output.getvalue(),
        file_name=f"Nilai_{mapel}_{semester}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
