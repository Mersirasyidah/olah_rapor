import csv
import os
import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np
from typing import List, Dict, Any, Union

# =========================================================
# KONFIGURASI DAN DATA LOADING
# =========================================================

COMMON_SUBJECTS = ["Matematika", "Bahasa Inggris", "IPA", "IPS", "Bahasa Indonesia","Seni Budaya", "P.Pancasila", "Pendidikan Agama", "Bahasa Jawa", "PJOK", "Informatika", "Prakarya"]
CLASS_OPTIONS = ["7A", "7B", "7C", "7D", "8A", "8B", "9A", "9B"]
YEAR_OPTIONS = ["2025/2026", "2026/2027", "2027/2028", "2028/2029", "2029/2030"]

SCORE_COLUMNS = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5', 'LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5', 'PTS', 'SAS', 'NR']
INPUT_SCORE_COLS = [c for c in SCORE_COLUMNS if c != 'NR']

COLUMN_DISPLAY_MAP = {
    'TP1': 'TP-1', 'TP2': 'TP-2', 'TP3': 'TP-3', 'TP4': 'TP-4', 'TP5': 'TP-5',
    'LM_1': 'LM-1', 'LM_2': 'LM-2', 'LM_3': 'LM-3', 'LM_4': 'LM-4', 'LM_5': 'LM-5',
    'PTS': 'PTS', 'SAS': 'SAS/SAT', 'NR': 'NR',
    'Avg_TP': 'Rata-rata TP',
    'Avg_LM': 'Rata-rata LM',
    'Avg_PSA': 'Rata-rata PSA',
    'Deskripsi_NR': 'Deskripsi Rapor'
}
KKM = 80

@st.cache_data
def load_dummy_data():
    data = {
        'NIS': [1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008],
        'Nama': ['Budi Santoso', 'Citra Dewi', 'Doni Pratama', 'Eka Fitriani', 'Fajar Nur', 'Gita Cahyani', 'Hendra Wijaya', 'Irma Suryani'],
        'Kelas': ['7A', '7A', '7A', '7A', '7B', '7B', '7C', '7C'],
    }
    df = pd.DataFrame(data)
    for col in INPUT_SCORE_COLS:
        df[col] = np.random.uniform(65.0, 95.0, size=len(df)).round(1)
    df['NR'] = 0 
    df['NIS'] = df['NIS'].astype(str)
    return df

@st.cache_data
def load_base_student_data():
    file_path = "daftar_siswa.csv"
    try:
        if os.path.exists(file_path):
            df = pd.read_csv(file_path)
            column_rename_map = {'NISN': 'NIS', 'NAMA': 'Nama', 'KLAS': 'Kelas', 'KLS': 'Kelas'}
            df.columns = [col.strip() for col in df.columns]
            df = df.rename(columns=column_rename_map, errors='ignore')
            required_cols = ['NIS', 'Nama', 'Kelas']
            if all(col in df.columns for col in required_cols):
                for col in INPUT_SCORE_COLS:
                    if col not in df.columns:
                        df[col] = 0.0
                df['NIS'] = df['NIS'].astype(str).str.strip()
                df['Nama'] = df['Nama'].astype(str).str.strip()
                df['Kelas'] = df['Kelas'].astype(str).str.strip().str.upper()
                for col in INPUT_SCORE_COLS:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
                return df
        return load_dummy_data()
    except Exception:
        return load_dummy_data()

# =========================================================
# FUNGSI PERHITUNGAN DAN LOGIKA VALIDASI
# =========================================================

def calculate_nr(df_input: pd.DataFrame) -> pd.DataFrame:
    df = df_input.copy()
    tp_cols = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    df['Avg_TP'] = df[tp_cols].replace(0.0, np.nan).mean(axis=1).round(2)

    lm_cols = ['LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5']
    df['Avg_LM'] = df[lm_cols].replace(0.0, np.nan).mean(axis=1).round(2)

    df['Avg_PSA'] = np.where((df['PTS'] + df['SAS']) > 0.0, (df['PTS'] + df['SAS']) / 2, 0.0).round(2)

    nr_components = df[['Avg_TP', 'Avg_LM', 'Avg_PSA']].copy()
    nr_components['Avg_TP_weighted'] = nr_components['Avg_TP'].fillna(0.0) * 1
    nr_components['Avg_LM_weighted'] = nr_components['Avg_LM'].fillna(0.0) * 1
    nr_components['Avg_PSA_weighted'] = nr_components['Avg_PSA'].fillna(0.0) * 2
    
    sum_components = nr_components['Avg_TP_weighted'] + nr_components['Avg_LM_weighted'] + nr_components['Avg_PSA_weighted']
    count_components = nr_components.apply(lambda row: sum([
        1 if pd.notna(row['Avg_TP']) and row['Avg_TP'] > 0.0 else 0,
        1 if pd.notna(row['Avg_LM']) and row['Avg_LM'] > 0.0 else 0,
        2 if pd.notna(row['Avg_PSA']) and row['Avg_PSA'] > 0.0 else 0
    ]), axis=1)

    df['NR_FLOAT'] = np.where(count_components > 0, sum_components / count_components, 0.0)
    df['NR'] = df['NR_FLOAT'].round(0).astype(int)
    return df.drop(columns=['NR_FLOAT'])

def calculate_tk_status(df_input: pd.DataFrame) -> pd.DataFrame:
    df = df_input.copy()
    tp_cols = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    threshold = float(KKM)

    for tp in tp_cols:
        tk = f'TK_{tp}'
        df[tk] = df[tp].apply(lambda x: "" if x <= 0.0 or pd.isna(x) else "T" if x >= threshold else "R")

    def apply_validation_rule(row):
        filled_tps = {tp: row[tp] for tp in tp_cols if pd.notna(row[tp]) and row[tp] > 0.0}
        if len(filled_tps) > 0:
            all_t = all(v >= threshold for v in filled_tps.values())
            if all_t:
                smallest_tp = min(filled_tps, key=filled_tps.get)
                row[f'TK_{smallest_tp}'] = 'R'
        return row

    return df.apply(apply_validation_rule, axis=1)

def generate_nr_description(df_input: pd.DataFrame) -> pd.DataFrame:
    df = df_input.copy()
    tp_cols = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    tk_cols = [f'TK_{col}' for col in tp_cols]
    descriptions = []

    for _, row in df.iterrows():
        remidi_tps = [i + 1 for i, col in enumerate(tk_cols) if row.get(col) == 'R']
        if remidi_tps:
            tp_list = [f"TP-{i}" for i in remidi_tps]
            tp_list_str = f"{', '.join(tp_list[:-1])}, dan {tp_list[-1]}" if len(tp_list) > 1 else tp_list[0]
            description = f"Ananda perlu meningkatkan pemahaman pada materi di {tp_list_str}."
        elif row[tp_cols].sum() == 0.0:
            description = "Nilai belum diinput."
        else:
            description = "Ananda telah menunjukkan penguasaan materi yang sangat baik pada seluruh TP."
        descriptions.append(description)

    df['Deskripsi_NR'] = descriptions
    return df

# =========================================================
# HELPERS EXCEL & RUMUS VALIDASI
# =========================================================

def col_idx_to_excel(col_idx):
    col_idx += 1
    letters = ""
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

def write_form_nilai_sheet(df, mapel, semester, kelas, tp, guru, nip, writer, sheet_name):
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)

    # Formats
    fmt_header = workbook.add_format({'border': 1, 'bold': True, 'bg_color': '#D9E1F2', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    fmt_border = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.0'})
    fmt_text = workbook.add_format({'border': 1, 'align': 'left'})
    fmt_protected = workbook.add_format({'border': 1, 'bg_color': '#FFF2CC', 'align': 'center', 'num_format': '0.0', 'locked': True})
    fmt_int = workbook.add_format({'border': 1, 'bg_color': '#FFF2CC', 'align': 'center', 'num_format': '0', 'locked': True})

    # Columns Setup
    TP_COLS = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    LM_COLS = ['LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5']
    CORE_COLS = TP_COLS + LM_COLS + ['PTS', 'SAS', 'Avg_TP', 'Avg_LM', 'Avg_PSA', 'NR']
    TK_COLS = [f'TK_{c}' for c in TP_COLS]
    
    HEADER_LABELS = ['NIS', 'NAMA SISWA', 'KELAS'] + \
                   [COLUMN_DISPLAY_MAP.get(c, c) for c in CORE_COLS] + \
                   [f"Status TP-{i+1}" for i in range(5)] + ['Deskripsi Rapor']

    # Header Info
    worksheet.write(0, 0, 'Mata Pelajaran: ' + str(mapel))
    worksheet.write(1, 0, 'Kelas: ' + str(kelas))
    worksheet.write(2, 0, 'Semester: ' + str(semester))
    worksheet.write(3, 0, 'KKTP: ' + str(KKM))

    # Table Header
    START_ROW = 6
    for i, label in enumerate(HEADER_LABELS):
        worksheet.write(START_ROW, i, label, fmt_header)

    # Identifikasi Index Kolom untuk Rumus
    idx_tp_start = 3 
    idx_tp_end = 7
    idx_lm_start = 8
    idx_lm_end = 12
    idx_pts = 13
    idx_sas = 14
    idx_avg_tp = 15
    idx_avg_lm = 16
    idx_avg_psa = 17
    idx_nr = 18
    idx_status_start = 19
    idx_desc = 24

    for r_idx in range(len(df)):
        row = START_ROW + 1 + r_idx
        excel_row = row + 1
        
        # Data Statis
        worksheet.write(row, 0, df.iloc[r_idx]['NIS'], fmt_text)
        worksheet.write(row, 1, df.iloc[r_idx]['Nama'], fmt_text)
        worksheet.write(row, 2, df.iloc[r_idx]['Kelas'], fmt_text)

        # Input Nilai (TP & LM & PTS/SAS)
        for i, col_name in enumerate(TP_COLS + LM_COLS + ['PTS', 'SAS']):
            val = df.iloc[r_idx][col_name]
            worksheet.write(row, 3 + i, val if val > 0 else "", fmt_border)

        # RUMUS: Rata-rata TP
        col_tp_s = col_idx_to_excel(idx_tp_start)
        col_tp_e = col_idx_to_excel(idx_tp_end)
        worksheet.write_formula(row, idx_avg_tp, f'=IFERROR(AVERAGEIF({col_tp_s}{excel_row}:{col_tp_e}{excel_row},">0"),0)', fmt_protected)

        # RUMUS: Rata-rata LM
        col_lm_s = col_idx_to_excel(idx_lm_start)
        col_lm_e = col_idx_to_excel(idx_lm_end)
        worksheet.write_formula(row, idx_avg_lm, f'=IFERROR(AVERAGEIF({col_lm_s}{excel_row}:{col_lm_e}{excel_row},">0"),0)', fmt_protected)

        # RUMUS: Rata-rata PSA
        c_pts = col_idx_to_excel(idx_pts)
        c_sas = col_idx_to_excel(idx_sas)
        worksheet.write_formula(row, idx_avg_psa, f'=IF(({c_pts}{excel_row}+{c_sas}{excel_row})>0,({c_pts}{excel_row}+{c_sas}{excel_row})/2,0)', fmt_protected)

        # RUMUS: NR
        c_atp = col_idx_to_excel(idx_avg_tp)
        c_alm = col_idx_to_excel(idx_avg_lm)
        c_apsa = col_idx_to_excel(idx_avg_psa)
        denom = f'(({c_atp}{excel_row}>0)*1+({c_alm}{excel_row}>0)*1+({c_apsa}{excel_row}>0)*2)'
        worksheet.write_formula(row, idx_nr, f'=IF({c_apsa}{excel_row}>0,ROUND(({c_atp}{excel_row}+{c_alm}{excel_row}+2*{c_apsa}{excel_row})/MAX(1,{denom}),0),0)', fmt_int)

        # LOGIKA VALIDASI: Status TP (T/R) - Minimal 1 R jika semua >= 80
        for i in range(5):
            curr_tp_col = col_idx_to_excel(idx_tp_start + i)
            # Rumus: Jika TP < KKM -> R. Jika Semua TP >= KKM AND Nilai ini adalah Nilai Terkecil -> R. Selain itu T.
            formula_status = (
                f'=IF({curr_tp_col}{excel_row}=0,"",'
                f'IF({curr_tp_col}{excel_row}<{KKM},"R",'
                f'IF(AND(MIN({col_tp_s}{excel_row}:{col_tp_e}{excel_row})>={KKM},'
                f'{curr_tp_col}{excel_row}=MIN({col_tp_s}{excel_row}:{col_tp_e}{excel_row})),"R","T")))'
            )
            worksheet.write_formula(row, idx_status_start + i, formula_status, fmt_protected)

        # Deskripsi
        worksheet.write(row, idx_desc, df.iloc[r_idx]['Deskripsi_NR'], fmt_text)

    # Column Width
    worksheet.set_column(1, 1, 25)
    worksheet.set_column(idx_desc, idx_desc, 60)

# =========================================================
# STREAMLIT UI
# =========================================================

st.set_page_config(layout="wide", page_title="Editor Nilai Rapor")

# Muat data
df_all_students_base = load_base_student_data()
df_all_students = df_all_students_base

st.title("Olah Nilai Rapor dengan Validasi Otomatis")

col_set1, col_set2 = st.columns(2)
with col_set1:
    mapel = st.selectbox("Mata Pelajaran", COMMON_SUBJECTS)
    available_classes = sorted(df_all_students['Kelas'].unique().tolist())
    kelas_terpilih = st.multiselect("Pilih Kelas", available_classes, default=available_classes[:1])
with col_set2:
    tp_tahun = st.selectbox("Tahun Pelajaran", YEAR_OPTIONS)
    guru = st.text_input("Guru Pengampu")

if kelas_terpilih:
    df_filtered = df_all_students[df_all_students['Kelas'].isin(kelas_terpilih)].copy()
    
    # Hitung data untuk tampilan preview
    df_filtered = calculate_nr(df_filtered)
    df_filtered = calculate_tk_status(df_filtered)
    df_filtered = generate_nr_description(df_filtered)

    st.subheader("Data Editor")
    # Mapping untuk editor
    df_editor = df_filtered.copy()
    for col in INPUT_SCORE_COLS:
        df_editor[col] = df_editor[col].replace(0.0, "")

    edited_data = st.data_editor(df_editor, hide_index=True, use_container_width=True)

    # Ekspor ke Excel
    if st.button("Generate File Excel"):
        # Pastikan data yang diunduh adalah hasil input terbaru dari editor
        df_final = edited_data.copy()
        for col in INPUT_SCORE_COLS:
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0.0)
        
        # Hitung ulang logika Python agar deskripsi di excel sinkron
        df_final = calculate_nr(df_final)
        df_final = calculate_tk_status(df_final)
        df_final = generate_nr_description(df_final)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for kls in kelas_terpilih:
                df_kls = df_final[df_final['Kelas'] == kls]
                write_form_nilai_sheet(df_kls, mapel, "Ganjil", kls, tp_tahun, guru, "-", writer, f"Nilai - {kls}")
        
        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name=f"Nilai_Rapor_{mapel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Silakan pilih kelas terlebih dahulu.")
