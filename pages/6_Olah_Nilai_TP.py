import csv
import os
import streamlit as st
import pandas as pd
from io import BytesIO
import numpy as np

# =========================================================
# KONFIGURASI DAN DATA LOADING
# =========================================================

COMMON_SUBJECTS = ["Matematika", "Bahasa Inggris", "IPA", "IPS", "Bahasa Indonesia","Seni Budaya", "P.Pancasila", "Pendidikan Agama", "Bahasa Jawa", "PJOK", "Informatika", "Prakarya"]
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

@st.cache_data
def load_dummy_data():
    data = {
        'NIS': ['1001', '1002', '1003', '1004'],
        'Nama': ['Budi Santoso', 'Citra Dewi', 'Doni Pratama', 'Eka Fitriani'],
        'Kelas': ['7A', '7A', '7A', '7A'],
    }
    df = pd.DataFrame(data)
    for col in INPUT_SCORE_COLS:
        df[col] = 0.0
    return df

# =========================================================
# FUNGSI PERHITUNGAN DAN LOGIKA VALIDASI
# =========================================================

def calculate_nr(df_input: pd.DataFrame) -> pd.DataFrame:
    df = df_input.copy()
    tp_cols = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    lm_cols = ['LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5']
    
    df['Avg_TP'] = df[tp_cols].replace(0.0, np.nan).mean(axis=1).round(2)
    df['Avg_LM'] = df[lm_cols].replace(0.0, np.nan).mean(axis=1).round(2)
    df['Avg_PSA'] = np.where((df['PTS'] + df['SAS']) > 0.0, (df['PTS'] + df['SAS']) / 2, 0.0).round(2)

    def row_nr(row):
        comps, w = [], []
        if pd.notna(row['Avg_TP']) and row['Avg_TP'] > 0: comps.append(row['Avg_TP']); w.append(1)
        if pd.notna(row['Avg_LM']) and row['Avg_LM'] > 0: comps.append(row['Avg_LM']); w.append(1)
        if row['Avg_PSA'] > 0: comps.append(row['Avg_PSA']); w.append(2)
        if not comps: return 0
        return int(round(sum(c * weight for c, weight in zip(comps, w)) / sum(w)))

    df['NR'] = df.apply(row_nr, axis=1)
    return df

def calculate_tk_status(df_input: pd.DataFrame) -> pd.DataFrame:
    df = df_input.copy()
    tp_cols = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    threshold = float(KKM)

    def apply_validation_rule(row):
        # Inisialisasi status dasar
        for tp in tp_cols:
            row[f'TK_{tp}'] = "" if row[tp] <= 0 else ("T" if row[tp] >= threshold else "R")
            
        filled_tps = [tp for tp in tp_cols if pd.notna(row[tp]) and row[tp] > 0.0]
        
        if len(filled_tps) > 0:
            # Cek apakah semua nilai yang diisi >= 80
            all_t = all(row[tp] >= threshold for tp in filled_tps)
            
            if all_t:
                # Cari nilai terkecil
                min_val = min([row[tp] for tp in filled_tps])
                # Cari TP mana saja yang punya nilai terkecil itu
                candidates = [tp for tp in filled_tps if row[tp] == min_val]
                # AMBIL HANYA SATU (paling pertama muncul) untuk jadi R
                target_tp = candidates[0]
                
                # Reset semua yang diisi jadi T dulu
                for tp in filled_tps:
                    row[f'TK_{tp}'] = "T"
                # Set satu yang terpilih jadi R
                row[f'TK_{target_tp}'] = "R"
        return row

    return df.apply(apply_validation_rule, axis=1)

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
# HELPERS EXCEL & RUMUS VALIDASI
# =========================================================

def col_idx_to_excel(col_idx):
    col_idx += 1
    letters = ""
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

def write_form_nilai_sheet(df, mapel, semester, kelas, tp_tahun, guru, writer, sheet_name):
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)

    # Formats
    fmt_header = workbook.add_format({'border': 1, 'bold': True, 'bg_color': '#D9E1F2', 'align': 'center', 'valign': 'vcenter'})
    fmt_border = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.0'})
    fmt_text = workbook.add_format({'border': 1, 'align': 'left'})
    fmt_protected = workbook.add_format({'border': 1, 'bg_color': '#FFF2CC', 'align': 'center', 'num_format': '0.0'})
    fmt_int = workbook.add_format({'border': 1, 'bg_color': '#FFF2CC', 'align': 'center', 'num_format': '0'})

    # Header Labels
    TP_COLS = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    CORE_COLS = TP_COLS + ['LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5', 'PTS', 'SAS', 'Avg_TP', 'Avg_LM', 'Avg_PSA', 'NR']
    HEADER_LABELS = ['NIS', 'NAMA SISWA', 'KELAS'] + [COLUMN_DISPLAY_MAP.get(c, c) for c in CORE_COLS] + \
                   [f"Status TP-{i+1}" for i in range(5)] + ['Deskripsi Rapor']

    # Header Info
    worksheet.write(0, 0, f'Mata Pelajaran: {mapel}')
    worksheet.write(1, 0, f'Kelas: {kelas} | Semester: {semester}')
    worksheet.write(2, 0, f'Tahun Pelajaran: {tp_tahun} | KKTP: {KKM}')

    START_ROW = 5
    for i, label in enumerate(HEADER_LABELS):
        worksheet.write(START_ROW, i, label, fmt_header)

    # Column Index Mapping
    idx_tp_s, idx_tp_e = 3, 7
    idx_avg_tp, idx_avg_lm, idx_avg_psa, idx_nr = 15, 16, 17, 18
    idx_status_start, idx_desc = 19, 24

    for r_idx, (_, row_data) in enumerate(df.iterrows()):
        row = START_ROW + 1 + r_idx
        ex_r = row + 1
        
        worksheet.write(row, 0, row_data['NIS'], fmt_text)
        worksheet.write(row, 1, row_data['Nama'], fmt_text)
        worksheet.write(row, 2, row_data['Kelas'], fmt_text)

        # Input Nilai
        for i, col_name in enumerate(CORE_COLS[:12]):
            val = row_data[col_name]
            worksheet.write(row, 3 + i, val if val > 0 else "", fmt_border)

        # Rumus Rata-rata & NR
        tp_range = f"{col_idx_to_excel(idx_tp_s)}{ex_r}:{col_idx_to_excel(idx_tp_e)}{ex_r}"
        worksheet.write_formula(row, idx_avg_tp, f'=IFERROR(AVERAGEIF({tp_range},">0"),0)', fmt_protected)
        
        lm_range = f"{col_idx_to_excel(8)}{ex_r}:{col_idx_to_excel(12)}{ex_r}"
        worksheet.write_formula(row, idx_avg_lm, f'=IFERROR(AVERAGEIF({lm_range},">0"),0)', fmt_protected)
        
        psa_sum = f"({col_idx_to_excel(13)}{ex_r}+{col_idx_to_excel(14)}{ex_r})"
        worksheet.write_formula(row, idx_avg_psa, f'=IF({psa_sum}>0,{psa_sum}/2,0)', fmt_protected)

        # Rumus Status TP (Sama dengan logika Python: MATCH-MIN)
        for i in range(5):
            curr_c = col_idx_to_excel(idx_tp_s + i)
            f_status = (
                f'=IF({curr_c}{ex_r}=0,"",'
                f'IF({curr_c}{ex_r}<{KKM},"R",'
                f'IF(AND(MIN({tp_range})>={KKM}, MATCH(MIN({tp_range}),{tp_range},0)={i+1}),"R","T")))'
            )
            worksheet.write_formula(row, idx_status_start + i, f_status, fmt_protected)

        worksheet.write(row, idx_desc, row_data['Deskripsi_NR'], fmt_text)

    worksheet.set_column(1, 1, 25)
    worksheet.set_column(idx_desc, idx_desc, 60)

# =========================================================
# STREAMLIT UI
# =========================================================

st.set_page_config(layout="wide", page_title="Editor Nilai Rapor")

with st.sidebar:
    st.header("📌 Identitas Rapor")
    mapel = st.selectbox("Mata Pelajaran", COMMON_SUBJECTS)
    tahun = st.selectbox("Tahun Pelajaran", YEAR_OPTIONS)
    semester = st.radio("Semester", ["Ganjil", "Genap"])
    kelas_pilihan = st.multiselect("Pilih Kelas", ["7A", "7B", "8A", "8B"], default=["7A"])
    guru = st.text_input("Nama Guru")

df_base = load_dummy_data()

if kelas_pilihan:
    st.title(f"Editor Nilai: {mapel}")
    
    # Pre-process preview
    df_view = calculate_nr(df_base)
    df_view = calculate_tk_status(df_view)
    df_view = generate_nr_description(df_view)

    # Data Editor
    edited_df = st.data_editor(df_view, hide_index=True, use_container_width=True)

    if st.button("Generate File Excel"):
        # Final Calculation sebelum export
        final_df = edited_df.copy()
        for col in INPUT_SCORE_COLS:
            final_df[col] = pd.to_numeric(final_df[col], errors='coerce').fillna(0.0)
        
        final_df = calculate_nr(final_df)
        final_df = calculate_tk_status(final_df)
        final_df = generate_nr_description(final_df)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for k in kelas_pilihan:
                write_form_nilai_sheet(final_df, mapel, semester, k, tahun, guru, writer, f"Nilai {k}")
        
        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name=f"Nilai_{mapel}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
