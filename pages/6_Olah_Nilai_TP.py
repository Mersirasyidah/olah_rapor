import csv
import os
import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np
from typing import List

# =========================================================
# KONFIGURASI
# =========================================================

COMMON_SUBJECTS = ["Matematika", "Bahasa Inggris", "IPA", "IPS", "Bahasa Indonesia", "Seni Budaya", "P.Pancasila", "Pendidikan Agama", "Bahasa Jawa", "PJOK", "Informatika", "Prakarya"]
CLASS_OPTIONS = ["7A", "7B", "7C", "7D", "8A", "8B", "9A", "9B"]
YEAR_OPTIONS = ["2025/2026", "2026/2027", "2027/2028", "2028/2029", "2029/2030"]

SCORE_COLUMNS = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5', 'LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5', 'PTS', 'SAS', 'NR']
INPUT_SCORE_COLS = [c for c in SCORE_COLUMNS if c != 'NR']

COLUMN_DISPLAY_MAP = {
    'TP1': 'TP-1', 'TP2': 'TP-2', 'TP3': 'TP-3', 'TP4': 'TP-4', 'TP5': 'TP-5',
    'LM_1': 'LM-1', 'LM_2': 'LM-2', 'LM_3': 'LM-3', 'LM_4': 'LM-4', 'LM_5': 'LM-5',
    'PTS': 'PTS', 'SAS': 'SAS/SAT', 'NR': 'NR',
    'Avg_TP': 'Rata-rata TP', 'Avg_LM': 'Rata-rata LM', 'Avg_PSA': 'Rata-rata PSA',
    'Deskripsi_NR': 'Deskripsi Rapor'
}
KKM = 80

# =========================================================
# FUNGSI PERHITUNGAN (LOGIKA PYTHON)
# =========================================================

def calculate_nr(df_input: pd.DataFrame) -> pd.DataFrame:
    df = df_input.copy()
    tp_cols = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    lm_cols = ['LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5']
    
    df['Avg_TP'] = df[tp_cols].replace(0.0, np.nan).mean(axis=1).round(2)
    df['Avg_LM'] = df[lm_cols].replace(0.0, np.nan).mean(axis=1).round(2)
    df['Avg_PSA'] = np.where((df['PTS'] + df['SAS']) > 0.0, (df['PTS'] + df['SAS']) / 2, 0.0).round(2)

    # Bobot: TP(1), LM(1), PSA(2)
    def row_nr(row):
        comps = []
        w = []
        if pd.notna(row['Avg_TP']) and row['Avg_TP'] > 0: comps.append(row['Avg_TP']); w.append(1)
        if pd.notna(row['Avg_LM']) and row['Avg_LM'] > 0: comps.append(row['Avg_LM']); w.append(1)
        if row['Avg_PSA'] > 0: comps.append(row['Avg_PSA']); w.append(2)
        
        if not comps: return 0
        return round(sum(c * weight for c, weight in zip(comps, w)) / sum(w))

    df['NR'] = df.apply(row_nr, axis=1).astype(int)
    return df

def calculate_tk_status(df_input: pd.DataFrame) -> pd.DataFrame:
    df = df_input.copy()
    tp_cols = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    
    def apply_validation(row):
        filled_tps = [tp for tp in tp_cols if pd.notna(row[tp]) and row[tp] > 0]
        # Inisialisasi awal
        for tp in tp_cols:
            row[f'TK_{tp}'] = "" if row[tp] <= 0 else ("T" if row[tp] >= KKM else "R")
        
        if filled_tps:
            # Jika semua yang diisi nilainya >= KKM (Tuntas semua)
            if all(row[tp] >= KKM for tp in filled_tps):
                # Cari nilai terkecil
                min_val = min([row[tp] for tp in filled_tps])
                # Cari TP pertama yang memiliki nilai terkecil tersebut
                for tp in filled_tps:
                    if row[tp] == min_val:
                        row[f'TK_{tp}'] = "R"
                        break # Hanya satu yang dijadikan R
        return row

    return df.apply(apply_validation, axis=1)

def generate_nr_description(df_input: pd.DataFrame) -> pd.DataFrame:
    df = df_input.copy()
    tp_cols = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    
    def get_desc(row):
        remidi = [f"TP-{i+1}" for i, tp in enumerate(tp_cols) if row.get(f'TK_{tp}') == 'R']
        if remidi:
            tp_str = f"{', '.join(remidi[:-1])}, dan {remidi[-1]}" if len(remidi) > 1 else remidi[0]
            return f"Ananda perlu meningkatkan pemahaman pada materi di {tp_str}."
        if row[tp_cols].sum() == 0: return "Nilai belum diinput."
        return "Ananda telah menunjukkan penguasaan materi yang sangat baik pada seluruh TP."

    df['Deskripsi_NR'] = df.apply(get_desc, axis=1)
    return df

# =========================================================
# EXCEL GENERATOR (DENGAN RUMUS MATCH-MIN)
# =========================================================

def col_idx_to_excel(col_idx):
    col_idx += 1
    letters = ""
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

def write_form_nilai_sheet(df, mapel, semester, kelas, tp, guru, writer, sheet_name):
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)

    # Formats
    header_fmt = workbook.add_format({'border': 1, 'bold': True, 'bg_color': '#D9E1F2', 'align': 'center', 'valign': 'vcenter'})
    border_fmt = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.0'})
    text_fmt = workbook.add_format({'border': 1, 'align': 'left'})
    protected_fmt = workbook.add_format({'border': 1, 'bg_color': '#FFF2CC', 'align': 'center', 'num_format': '0.0', 'locked': True})
    int_fmt = workbook.add_format({'border': 1, 'bg_color': '#FFF2CC', 'align': 'center', 'num_format': '0', 'locked': True})

    # Header Labels
    TP_LIST = ['TP-1', 'TP-2', 'TP-3', 'TP-4', 'TP-5']
    CORE_LABELS = TP_LIST + ['LM-1', 'LM-2', 'LM-3', 'LM-4', 'LM-5', 'PTS', 'SAS/SAT', 'Rata-rata TP', 'Rata-rata LM', 'Rata-rata PSA', 'NR']
    STATUS_LABELS = [f"Status {t}" for t in TP_LIST]
    ALL_HEADERS = ['NIS', 'NAMA SISWA', 'KELAS'] + CORE_LABELS + STATUS_LABELS + ['Deskripsi Rapor']

    # Header Info
    worksheet.write(0, 0, f"Mata Pelajaran: {mapel}")
    worksheet.write(1, 0, f"Kelas: {kelas} | Semester: {semester}")
    worksheet.write(2, 0, f"Tahun: {tp} | KKTP: {KKM}")

    START_ROW = 5
    for i, h in enumerate(ALL_HEADERS):
        worksheet.write(START_ROW, i, h, header_fmt)

    # Column Indexes
    idx_tp_start = 3
    idx_tp_end = 7
    idx_avg_tp = 15
    idx_avg_lm = 16
    idx_avg_psa = 17
    idx_nr = 18
    idx_status_start = 19
    idx_desc = 24

    for r_idx, (_, row_data) in enumerate(df.iterrows()):
        r = START_ROW + 1 + r_idx
        ex_r = r + 1
        
        # Identitas
        worksheet.write(r, 0, row_data['NIS'], text_fmt)
        worksheet.write(r, 1, row_data['Nama'], text_fmt)
        worksheet.write(r, 2, row_data['Kelas'], text_fmt)

        # Nilai Input (TP, LM, PTS, SAS)
        raw_cols = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5', 'LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5', 'PTS', 'SAS']
        for i, col in enumerate(raw_cols):
            val = row_data[col]
            worksheet.write(r, 3 + i, val if val > 0 else "", border_fmt)

        # Rumus Rata-rata
        tp_rng = f"{col_idx_to_excel(idx_tp_start)}{ex_r}:{col_idx_to_excel(idx_tp_end)}{ex_r}"
        worksheet.write_formula(r, idx_avg_tp, f'=IFERROR(AVERAGEIF({tp_rng},">0"),0)', protected_fmt)
        
        lm_rng = f"{col_idx_to_excel(8)}{ex_r}:{col_idx_to_excel(12)}{ex_r}"
        worksheet.write_formula(r, idx_avg_lm, f'=IFERROR(AVERAGEIF({lm_rng},">0"),0)', protected_fmt)
        
        psa_sum = f"({col_idx_to_excel(13)}{ex_r}+{col_idx_to_excel(14)}{ex_r})"
        worksheet.write_formula(r, idx_avg_psa, f'=IF({psa_sum}>0,{psa_sum}/2,0)', protected_fmt)

        # Rumus NR
        c_tp, c_lm, c_psa = col_idx_to_excel(idx_avg_tp), col_idx_to_excel(idx_avg_lm), col_idx_to_excel(idx_avg_psa)
        denom = f'(({c_tp}{ex_r}>0)*1+({c_lm}{ex_r}>0)*1+({c_psa}{ex_r}>0)*2)'
        worksheet.write_formula(r, idx_nr, f'=IF({c_psa}{ex_r}>0,ROUND(({c_tp}{ex_r}+{c_lm}{ex_r}+2*{c_psa}{ex_r})/MAX(1,{denom}),0),0)', int_fmt)

        # LOGIKA VALIDASI STATUS TP (Paling Krusial)
        for i in range(5):
            curr_col = col_idx_to_excel(idx_tp_start + i)
            # Rumus: 
            # 1. Jika TP < KKM -> R
            # 2. Jika semua TP >= KKM, cari index terkecil pertama menggunakan MATCH(MIN)
            # 3. Jika index kolom saat ini == hasil MATCH, maka R.
            f_status = (
                f'=IF({curr_col}{ex_r}=0,"",'
                f'IF({curr_col}{ex_r}<{KKM},"R",'
                f'IF(AND(MIN({tp_rng})>={KKM}, MATCH(MIN({tp_rng}),{tp_rng},0)={i+1}),"R","T")))'
            )
            worksheet.write_formula(r, idx_status_start + i, f_status, protected_fmt)

        worksheet.write(r, idx_desc, row_data['Deskripsi_NR'], text_fmt)

    worksheet.set_column(1, 1, 30)
    worksheet.set_column(idx_desc, idx_desc, 60)

# =========================================================
# STREAMLIT APP
# =========================================================

st.set_page_config(layout="wide", page_title="Sistem Olah Nilai")

# Load data base
if 'base_data' not in st.session_state:
    st.session_state.base_data = pd.DataFrame({
        'NIS': [str(100 + i) for i in range(5)],
        'Nama': ['Siswa A', 'Siswa B', 'Siswa C', 'Siswa D', 'Siswa E'],
        'Kelas': ['7A']*5
    })
    for c in INPUT_SCORE_COLS: st.session_state.base_data[c] = 0.0

st.title("Griya Rapor: Editor Nilai Kurikulum Merdeka")

with st.sidebar:
    st.header("Pengaturan")
    mapel = st.selectbox("Mata Pelajaran", COMMON_SUBJECTS)
    kls_list = st.multiselect("Pilih Kelas", ["7A", "7B", "8A", "8B"], default=["7A"])
    tahun = st.selectbox("Tahun Pelajaran", YEAR_OPTIONS)
    guru = st.text_input("Nama Guru")

if kls_list:
    df_editor = st.session_state.base_data.copy()
    
    # Pre-calculate untuk preview
    df_view = calculate_nr(df_editor)
    df_view = calculate_tk_status(df_view)
    df_view = generate_nr_description(df_view)

    # Konfigurasi Tampilan
    st.subheader("Tabel Input Nilai")
    st.info("Input nilai (0-100). Gunakan titik atau koma untuk desimal.")
    
    res_df = st.data_editor(
        df_view,
        column_config={
            "NIS": st.column_config.TextColumn(disabled=True),
            "Nama": st.column_config.TextColumn(disabled=True),
            "Kelas": st.column_config.TextColumn(disabled=True),
            "NR": st.column_config.NumberColumn(disabled=True),
            "Deskripsi_NR": st.column_config.TextColumn(disabled=True, width="large"),
        },
        hide_index=True
    )

    if st.button("Generate & Unduh Excel"):
        # Proses Data dari Editor
        final_df = res_df.copy()
        for col in INPUT_SCORE_COLS:
            if final_df[col].dtype == object:
                final_df[col] = final_df[col].astype(str).str.replace(',', '.')
            final_df[col] = pd.to_numeric(final_df[col], errors='coerce').fillna(0.0)
        
        # Recalculate agar deskripsi di excel sesuai input terakhir
        final_df = calculate_nr(final_df)
        final_df = calculate_tk_status(final_df)
        final_df = generate_nr_description(final_df)

        output = BytesIO()
