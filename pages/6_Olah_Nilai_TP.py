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

# Variabel konfigurasi
COMMON_SUBJECTS = ["Matematika", "Bahasa Inggris", "IPA", "IPS", "Bahasa Indonesia","Seni Budaya", "P.Pancasila", "Pendidikan Agama", "Bahasa Jawa", "PJOK", "Informatika", "Prakarya"]
CLASS_OPTIONS = ["7A", "7B", "7C", "7D", "8A", "8B", "9A", "9B"]
YEAR_OPTIONS = ["2025/2026", "2026/2027", "2027/2028", "2028/2029", "2029/2030"]

# Kolom-kolom nilai yang akan diinput/dihitung
SCORE_COLUMNS = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5', 'LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5', 'PTS', 'SAS', 'NR']
INPUT_SCORE_COLS = [c for c in SCORE_COLUMNS if c != 'NR']

# Peta untuk nama tampilan kolom
COLUMN_DISPLAY_MAP = {
    'TP1': 'TP-1', 'TP2': 'TP-2', 'TP3': 'TP-3', 'TP4': 'TP-4', 'TP5': 'TP-5',
    'LM_1': 'LM-1', 'LM_2': 'LM-2', 'LM_3': 'LM-3', 'LM_4': 'LM-4', 'LM_5': 'LM-5',
    'PTS': 'PTS', 'SAS': 'SAS/SAT', 'NR': 'NR',
    'Avg_TP': 'Rata-rata TP',
    'Avg_LM': 'Rata-rata LM',
    'Avg_PSA': 'Rata-rata PSA',
    'Deskripsi_NR': 'Deskripsi Rapor'
}
# Threshold KKM/Batas Ketuntasan untuk menentukan Tingkat Ketercapaian (TK)
KKM = 80

@st.cache_data
def load_dummy_data():
    """Membuat DataFrame dummy untuk semua siswa (sebagai fallback)."""
    data = {
        'NIS': [1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008],
        'Nama': ['Budi Santoso', 'Citra Dewi', 'Doni Pratama', 'Eka Fitriani', 'Fajar Nur', 'Gita Cahyani', 'Hendra Wijaya', 'Irma Suryani'],
        'Kelas': ['7A', '7A', '7A', '7A', '7B', '7B', '7C', '7C'],
    }
    df = pd.DataFrame(data)
    # Inisialisasi kolom nilai input dengan angka acak yang realistis (sekarang FLOAT)
    for col in INPUT_SCORE_COLS:
        df[col] = np.random.uniform(65.0, 95.0, size=len(df)).round(1)
    df['NR'] = 0 # Initialize NR as integer (hasil akhir)
    df['NIS'] = df['NIS'].astype(str) # Pastikan NIS adalah string
    return df

@st.cache_data
def load_base_student_data():
    """
    Mencoba memuat data siswa dari daftar_siswa.csv secara default.
    """
    file_path = "daftar_siswa.csv"
    try:
        if os.path.exists(file_path):
            df = pd.read_csv(file_path)

            column_rename_map = {
                'NISN': 'NIS', 'NAMA': 'Nama', 'KLAS': 'Kelas', 'KLS': 'Kelas',
            }

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

                df = df.drop(columns=['Nilai Rata-rata'], errors='ignore')

                return df
            else:
                 return load_dummy_data()
        else:
            return load_dummy_data()

    except Exception as e:
        return load_dummy_data()


# Muat data dasar (prioritas CSV, fallback dummy)
df_all_students_base = load_base_student_data()

# =========================================================
# FUNGSI PERHITUNGAN DAN DESKRIPSI
# =========================================================

def calculate_nr(df_input: pd.DataFrame) -> pd.DataFrame:
    """
    Menghitung Nilai Rapor (NR) dan nilai perantara (Avg_TP, Avg_LM, Avg_PSA).
    """
    df = df_input.copy()

    # 1. Hitung Rata-rata TP
    tp_cols = [c for c in df.columns if c.startswith('TP') and len(c) == 3]
    df['Avg_TP'] = df[tp_cols].replace(0.0, np.nan).mean(axis=1).round(2)

    # 2. Hitung Rata-rata LM
    lm_cols = [c for c in df.columns if c.startswith('LM_') and len(c) == 4]
    df['Avg_LM'] = df[lm_cols].replace(0.0, np.nan).mean(axis=1).round(2)

    # 3. Hitung Rata-rata Penilaian Sumatif Akhir (PSA)
    df['Avg_PSA'] = np.where(
        (df['PTS'] + df['SAS']) > 0.0,
        (df['PTS'] + df['SAS']) / 2,
        0.0
    ).round(2)

    # 4. Hitung NR (Nilai Rapor)
    nr_components = df[['Avg_TP', 'Avg_LM', 'Avg_PSA']].copy()
    nr_components['Avg_TP_weighted'] = nr_components['Avg_TP'].fillna(0.0) * 1
    nr_components['Avg_LM_weighted'] = nr_components['Avg_LM'].fillna(0.0) * 1
    nr_components['Avg_PSA_weighted'] = nr_components['Avg_PSA'].fillna(0.0) * 2
    sum_components = (
        nr_components['Avg_TP_weighted'] +
        nr_components['Avg_LM_weighted'] +
        nr_components['Avg_PSA_weighted']
    )
    count_components = nr_components.apply(lambda row: sum([1 if pd.notna(row['Avg_TP']) and row['Avg_TP'] > 0.0 else 0,
                                            1 if pd.notna(row['Avg_LM']) and row['Avg_LM'] > 0.0 else 0,
                                                           2 if pd.notna(row['Avg_PSA']) and row['Avg_PSA'] > 0.0 else 0]), axis=1)

    df['NR_FLOAT'] = np.where(count_components > 0, sum_components / count_components, 0.0)
    df['NR_FLOAT'] = np.where(df['Avg_PSA'] > 0.0, df['NR_FLOAT'], 0.0)
    df['NR'] = df['NR_FLOAT'].round(0).astype(int)
    df = df.drop(columns=['NR_FLOAT'])

    return df

def calculate_tk_status(df_input: pd.DataFrame) -> pd.DataFrame:
    """Menentukan Status TP1–TP5 (T/R) termasuk validasi tambahan."""
    df = df_input.copy()
    tp_cols = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    tk_cols = [f'TK_{tp}' for tp in tp_cols]
    threshold = float(KKM)

    for tp in tp_cols:
        tk = f'TK_{tp}'
        df[tk] = df[tp].apply(
            lambda x: "" if x <= 0.0 or pd.isna(x)
            else "T" if x >= threshold
            else "R"
        )

    def apply_validation_rule(row):
        filled_tps = {tp: row[tp] for tp in tp_cols if pd.notna(row[tp]) and row[tp] > 0.0}

        if len(filled_tps) < 2:
            return row

        all_t = all(v >= threshold for v in filled_tps.values())

        if all_t:
            smallest_tp = min(filled_tps, key=filled_tps.get)
            smallest_tk = f'TK_{smallest_tp}'
            row[smallest_tk] = 'R'
        return row

    df = df.apply(apply_validation_rule, axis=1)

    return df

def generate_nr_description(df_input: pd.DataFrame) -> pd.DataFrame:
    """Membuat Deskripsi Naratif Nilai Rapor (Deskripsi_NR)."""
    df = df_input.copy()
    tp_cols_prefix = [c for c in df.columns if c.startswith('TP') and len(c) == 3]
    tk_cols = [f'TK_{col}' for col in tp_cols_prefix]
    descriptions = []

    for index, row in df.iterrows():
        remidi_tps = [i + 1 for i, col in enumerate(tk_cols) if row.get(col) == 'R']
        tp_scores_sum = row[tp_cols_prefix].sum()

        if remidi_tps:
            tp_list = [f"TP-{i}" for i in remidi_tps]
            if len(tp_list) > 1:
                tp_list_str = f"{', '.join(tp_list[:-1])}, dan {tp_list[-1]}"
            else:
                tp_list_str = tp_list[0]
            description = f"Ananda perlu meningkatkan pemahaman dan penguasaan pada materi di {tp_list_str}."

        elif tp_scores_sum == 0.0:
             description = "Nilai Tujuan Pembelajaran belum diinput."
        else:
             description = "Ananda telah menunjukkan penguasaan materi yang sangat baik dan tuntas pada seluruh Tujuan Pembelajaran."

        descriptions.append(description)

    df['Deskripsi_NR'] = descriptions
    return df

# =========================================================
# HELPERS UNTUK EXCEL (Tidak Berubah)
# =========================================================

def col_idx_to_excel(col_idx):
    """Convert 0-based column index to Excel column letters (0->A)."""
    col_idx += 1
    letters = ""
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        letters = chr(65 + remainder) + letters
        col_idx = int(col_idx)
    return letters

def write_form_nilai_sheet(df, mapel, semester, kelas, tp, guru, nip, writer, sheet_name):
    """Menulis satu sheet Form Nilai Siswa (Laporan Lengkap) ke writer yang sudah ada."""
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)

    # Definisi Format (Menggunakan format numerik 1 desimal untuk nilai input)
    border_format_float = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.0'
    })
    header_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bold': True, 'fg_color': '#D9E1F2', 'text_wrap': True
    })
    text_format = workbook.add_format({
        'border': 1, 'align': 'left', 'valign': 'vcenter'
    })
    status_tp_protected_format = workbook.add_format({
        'bg_color': '#FFF2CC', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'locked': True
    })
    formula_protected_format_float = workbook.add_format({
        'bg_color': '#FFF2CC', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'locked': True, 'num_format': '0.0'
    })
    formula_protected_format_int = workbook.add_format({
        'bg_color': '#FFF2CC', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'locked': True, 'num_format': '0'
    })
    header_info_format = workbook.add_format({
        'align': 'left', 'valign': 'vcenter'
    })

    LM_COLS_EXPORT = ['LM_1', 'LM_2', 'LM_3', 'LM_4', 'LM_5']
    TK_COLS_EXPORT = [c for c in df.columns if c.startswith('TK_TP') and len(c) == 6]
    TP_SCORE_COLS = ['TP1', 'TP2', 'TP3', 'TP4', 'TP5']
    INPUT_SCORE_COLS_ALL = TP_SCORE_COLS + LM_COLS_EXPORT + ['PTS', 'SAS']
    CORE_SCORE_COLS = TP_SCORE_COLS + LM_COLS_EXPORT + ['PTS', 'SAS', 'Avg_TP', 'Avg_LM', 'Avg_PSA', 'NR']
    FINAL_COLS_ORDER_DATA = ['NIS', 'Nama', 'Kelas'] + \
                            [c for c in CORE_SCORE_COLS if c in df.columns] + \
                            ['Deskripsi_NR']

    df_export = df[FINAL_COLS_ORDER_DATA].copy()

    for col in INPUT_SCORE_COLS_ALL:
        if col in df_export.columns:
            df_export[col] = pd.to_numeric(df_export[col], errors='coerce').fillna(0.0)
    if 'NR' in df_export.columns:
        df_export['NR'] = df_export['NR'].fillna(0).astype(int)

    tk_display_map = {c: c.replace('TK_', 'Status ') for c in TK_COLS_EXPORT}
    HEADER_COLS_ORDER = ['NIS', 'NAMA SISWA', 'KELAS'] + \
                  [COLUMN_DISPLAY_MAP.get(c, c) for c in CORE_SCORE_COLS if c in df.columns] + \
                  [tk_display_map.get(c, c) for c in TK_COLS_EXPORT] + \
                  [COLUMN_DISPLAY_MAP.get('Deskripsi_NR', 'Deskripsi Rapor')]

    # 2. MENULIS HEADER INFORMASI
    START_ROW_INFO = 0
    INFO_COL_START = 6
    worksheet.write(0 + START_ROW_INFO, 0, 'Mata Pelajaran', header_info_format)
    worksheet.write(0 + START_ROW_INFO, 2, ': ' + str(mapel), header_info_format)
    worksheet.write(1 + START_ROW_INFO, 0, 'Kelas', header_info_format)
    worksheet.write(1 + START_ROW_INFO, 2, ': ' + str(kelas), header_info_format)
    worksheet.write(2 + START_ROW_INFO, 0, 'Semester', header_info_format)
    worksheet.write(2 + START_ROW_INFO, 2, ': ' + str(semester), header_info_format)
    worksheet.write(3 + START_ROW_INFO, 0, 'KKTP', header_info_format)
    worksheet.write(3 + START_ROW_INFO, 2, f": {KKM}", header_info_format)

    worksheet.write(0 + START_ROW_INFO, INFO_COL_START, 'Tahun Pelajaran', header_info_format)
    worksheet.write(0 + START_ROW_INFO, INFO_COL_START + 2, ': ' + str(tp), header_info_format)
    worksheet.write(1 + START_ROW_INFO, INFO_COL_START, 'Guru Mata Pelelajaran', header_info_format)
    worksheet.write(1 + START_ROW_INFO, INFO_COL_START + 2, ': ' + str(guru), header_info_format)
    worksheet.write(2 + START_ROW_INFO, INFO_COL_START, 'NIP Guru', header_info_format)
    worksheet.write(2 + START_ROW_INFO, INFO_COL_START + 2, ': ' + str(nip), header_info_format)

    worksheet.set_column(0, 0, 20)
    worksheet.set_column(2, 2, 30)
    worksheet.set_column(INFO_COL_START, INFO_COL_START, 20)
    worksheet.set_column(INFO_COL_START + 2, INFO_COL_START + 2, 30)

    # 3. MENULIS DATA NILAI SISWA
    START_ROW_DATA = 6
    COL_OFFSET = 0

    for col_num, value in enumerate(HEADER_COLS_ORDER):
        worksheet.write(START_ROW_DATA, col_num + COL_OFFSET, value, header_format)

    cols = list(HEADER_COLS_ORDER)
    def idx_of(title): return cols.index(title) if title in cols else None
    idx_avg_tp = idx_of('Rata-rata TP')
    idx_avg_lm = idx_of('Rata-rata LM')
    idx_avg_psa = idx_of('Rata-rata PSA')
    idx_nr = idx_of('NR')
    status_tp_map = {f'Status TP{i}': f'TP-{i}' for i in range(1, 6)}
    status_tp_indices = {
        idx_of(status_name): idx_of(tp_score_name)
        for status_name, tp_score_name in status_tp_map.items()
        if idx_of(status_name) is not None and idx_of(tp_score_name) is not None
    }
    tp_score_display_titles = ['TP-1','TP-2','TP-3','TP-4','TP-5']
    lm_display_titles = ['LM-1','LM-2','LM-3','LM-4','LM-5']
    tp_score_indices = [cols.index(t) for t in tp_score_display_titles if t in cols]
    lm_indices = [cols.index(t) for t in lm_display_titles if t in cols]
    idx_pts = cols.index('PTS') if 'PTS' in cols else None
    idx_sas = cols.index('SAS/SAT') if 'SAS/SAT' in cols else None

    num_rows = len(df_export)
    REVERSE_COLUMN_MAP = {v: k for k, v in COLUMN_DISPLAY_MAP.items()}
    REVERSE_COLUMN_MAP['NAMA SISWA'] = 'Nama'
    REVERSE_COLUMN_MAP['KELAS'] = 'Kelas'

    for r_i in range(num_rows):
        row_num = START_ROW_DATA + 1 + r_i
        excel_row = row_num + 1

        for col_num, column_name in enumerate(HEADER_COLS_ORDER):
            if column_name.startswith('Status TP'):
                KKM_VALUE = KKM
                tp_score_idx = status_tp_indices.get(col_num)
                if tp_score_idx is not None:
                    col_tp_score = col_idx_to_excel(tp_score_idx + COL_OFFSET)
                    formula_tk = (
                        f'=IF({col_tp_score}{excel_row}=0,"",'
                        f'IF({col_tp_score}{excel_row}>={KKM_VALUE},"T","R"))'
                    )
                    worksheet.write_formula(row_num, col_num + COL_OFFSET, formula_tk, status_tp_protected_format)
                continue

            if column_name in ['Rata-rata TP', 'Rata-rata LM', 'Rata-rata PSA', 'NR']:
                continue

            data_column_name = REVERSE_COLUMN_MAP.get(column_name, column_name)
            if data_column_name not in df_export.columns:
                 continue

            cell_value = df_export.iloc[r_i, df_export.columns.get_loc(data_column_name)]
            current_format = border_format_float

            is_input_score = data_column_name in INPUT_SCORE_COLS_ALL
            if is_input_score and cell_value == 0.0:
                cell_value = ""

            if data_column_name in ['NIS', 'Nama', 'Kelas', 'Deskripsi_NR']:
                if isinstance(cell_value, (int, float)):
                    cell_value = str(cell_value)
                current_format = text_format
            else:
                current_format = border_format_float

            worksheet.write(row_num, col_num + COL_OFFSET, cell_value, current_format)

        # TULIS RUMUS AVG dan NR
        if idx_avg_tp is not None and tp_score_indices:
            first_tp_col = col_idx_to_excel(tp_score_indices[0] + COL_OFFSET)
            last_tp_col = col_idx_to_excel(tp_score_indices[-1] + COL_OFFSET)
            formula_avg_tp = f"=AVERAGEIF({first_tp_col}{excel_row}:{last_tp_col}{excel_row},\">0\")"
            worksheet.write_formula(row_num, idx_avg_tp + COL_OFFSET, formula_avg_tp, formula_protected_format_float)

        if idx_avg_lm is not None and lm_indices:
            first_lm_col = col_idx_to_excel(lm_indices[0] + COL_OFFSET)
            last_lm_col = col_idx_to_excel(lm_indices[-1] + COL_OFFSET)
            formula_avg_lm = f"=AVERAGEIF({first_lm_col}{excel_row}:{last_lm_col}{excel_row},\">0\")"
            worksheet.write_formula(row_num, idx_avg_lm + COL_OFFSET, formula_avg_lm, formula_protected_format_float)

        if idx_avg_psa is not None and idx_pts is not None and idx_sas is not None:
            col_pts = col_idx_to_excel(idx_pts + COL_OFFSET)
            col_sas = col_idx_to_excel(idx_sas + COL_OFFSET)
            formula_avg_psa = (
                f"=IF({col_pts}{excel_row}+{col_sas}{excel_row}>0, "
                f"({col_pts}{excel_row}+{col_sas}{excel_row})/2, 0)"
            )
            worksheet.write_formula(row_num, idx_avg_psa + COL_OFFSET, formula_avg_psa, formula_protected_format_float)

        if idx_nr is not None and idx_avg_tp is not None and idx_avg_lm is not None and idx_avg_psa is not None:
            col_avg_tp = col_idx_to_excel(idx_avg_tp + COL_OFFSET)
            col_avg_lm = col_idx_to_excel(idx_avg_lm + COL_OFFSET)
            col_avg_psa = col_idx_to_excel(idx_avg_psa + COL_OFFSET)

            calculation_denominator = (
                f"(({col_avg_tp}{excel_row}>0)*1+({col_avg_lm}{excel_row}>0)*1+({col_avg_psa}{excel_row}>0)*2)"
            )
            calculation_core = (
                f"({col_avg_tp}{excel_row}+{col_avg_lm}{excel_row}+2*{col_avg_psa}{excel_row})/"
                f"IF({calculation_denominator}=0,1,{calculation_denominator})"
            )
            formula_nr = (
                f"=IF({col_avg_psa}{excel_row}>0, "
                f"IFERROR(ROUND({calculation_core},0),0),0)"
            )
            worksheet.write_formula(row_num, idx_nr + COL_OFFSET, formula_nr, formula_protected_format_int)

    worksheet.set_column(0 + COL_OFFSET, 0 + COL_OFFSET, 5)
    worksheet.set_column(1 + COL_OFFSET, 1 + COL_OFFSET, 25)
    worksheet.set_column(2 + COL_OFFSET, 2 + COL_OFFSET, 7)
    idx_desc = idx_of('Deskripsi Rapor')
    end_col_numeric = idx_desc - 1 + COL_OFFSET if idx_desc is not None else len(HEADER_COLS_ORDER) - 1 + COL_OFFSET
    worksheet.set_column(3 + COL_OFFSET, end_col_numeric, 8)
    if idx_desc is not None:
        worksheet.set_column(idx_desc + COL_OFFSET, idx_desc + COL_OFFSET, 60)
    worksheet.freeze_panes(START_ROW_DATA + 2, 2)


def write_report_tk_sheet(df, mapel, kelas, tp, writer, sheet_name):
    """Menulis satu sheet Laporan TK ke writer yang sudah ada."""
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)

    border_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    header_format = workbook.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bold': True, 'fg_color': '#D9E1F2'
    })
    text_format = workbook.add_format({
        'border': 1, 'align': 'left', 'valign': 'vcenter'
    })
    header_info_format = workbook.add_format({
        'align': 'left', 'valign': 'vcenter'
    })

    tk_cols = [c for c in df.columns if c.startswith('TK_TP') and len(c) == 6]
    df_export = df[['NIS', 'Nama', 'Kelas', 'NR'] + tk_cols].copy()

    tk_headers_short = [f'KTP-{i}' for i in range(1, len(tk_cols) + 1)]
    df_export.columns = ['NIS', 'NAMA SISWA', 'KELAS', 'NR'] + tk_headers_short

    df_export['NR'] = df_export['NR'].fillna(0).astype(int)
    for col in [c for c in df_export.columns if c.startswith('KTP-')]:
        df_export[col] = df_export[col].astype(str).replace('nan', '')

    df_export['NR'] = df_export['NR'].apply(lambda x: '' if x == 0 else x)


    KKTP = KKM
    header_data = {
        'Keterangan': [
            'Mata Pelajaran', 'Kelas', 'Tahun Pelajaran', 'Batas Ketuntasan (KKTP)'
        ],
        'Nilai_Isian': [
            ': ' + str(mapel),
            ': ' + str(kelas),
            ': ' + str(tp),
            ': ' + str(KKTP)
        ]
    }
    combined_header_df = pd.DataFrame(header_data)

    for r_idx, row in combined_header_df.iterrows():
        worksheet.write(r_idx, 0, row['Keterangan'], header_info_format)
        worksheet.write(r_idx, 2, row['Nilai_Isian'], header_info_format)

    START_ROW_DATA = 6
    for col_num, value in enumerate(df_export.columns.values):
        worksheet.write(START_ROW_DATA, col_num, value, header_format)

    num_rows = len(df_export)
    num_cols = len(df_export.columns)
    for row_num in range(START_ROW_DATA + 1, START_ROW_DATA + num_rows + 1):
        for col_num in range(num_cols):
            cell_value = df_export.iloc[row_num - (START_ROW_DATA + 1), col_num]
            column_name = df_export.columns[col_num]
            if column_name in ['NIS', 'NAMA SISWA', 'KELAS']:
                if isinstance(cell_value, (int, float)):
                    cell_value = str(cell_value)
                worksheet.write(row_num, col_num, cell_value, text_format)
            elif column_name.startswith('KTP-'):
                worksheet.write_string(row_num, col_num, str(cell_value), border_format)
            else:
                worksheet.write(row_num, col_num, cell_value, border_format)

    worksheet.set_column(1, 1, 25)
    worksheet.set_column(0, num_cols-1, 8)
    worksheet.freeze_panes(START_ROW_DATA + 1, 2)


def export_multisheet_form_nilai(df_all: pd.DataFrame, classes: List[str], mapel, semester, tp, guru, nip):
    """Menghasilkan file Excel multisheet untuk Form Nilai."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for kelas in classes:
            df_kelas = df_all[df_all['Kelas'] == kelas].reset_index(drop=True)
            if not df_kelas.empty:
                sheet_name = f"{kelas} - Form Nilai"
                write_form_nilai_sheet(df_kelas, mapel, semester, kelas, tp, guru, nip, writer, sheet_name)
    return output.getvalue()

def export_multisheet_report_tk(df_all: pd.DataFrame, classes: List[str], mapel, tp):
    """Menghasilkan file Excel multisheet untuk Laporan TK."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for kelas in classes:
            df_kelas = df_all[df_all['Kelas'] == kelas].reset_index(drop=True)
            if not df_kelas.empty:
                sheet_name = f"{kelas} - Laporan TK"
                write_report_tk_sheet(df_kelas, mapel, kelas, tp, writer, sheet_name)
    return output.getvalue()

# =========================================================
# APLIKASI STREAMLIT UTAMA
# =========================================================
st.set_page_config(layout="wide", page_title="Editor Nilai Kurikulum Merdeka (LM x 5)")

# Injeksi CSS (Tidak Berubah)
st.markdown(
    """
    <style>
    /* Buat tabel st.dataframe bisa scroll horz & vert */
    div[data-testid="stDataFrame"] {
        width: 100%;
        overflow: auto;
        height: 60vh;
    }
    /* Freeze Header Row */
    div[data-testid="stDataFrame"] thead tr:first-child th {
        position: sticky;
        top: 0;
        background-color: #ffffff;
        z-index: 5;
        border-bottom: 1px solid #ddd;
    }
    /* Freeze kolom NIS (kolom ke-1) */
    div[data-testid="stDataFrame"] tbody tr td:nth-child(1), div[data-testid="stDataFrame"] thead tr th:nth-child(1) {
        position: sticky;
        left: 0;
        background-color: #ffffff;
        z-index: 4;
    }
    /* Freeze kolom Nama (kolom ke-2) */
    div[data-testid="stDataFrame"] tbody tr td:nth-child(2), div[data-testid="stDataFrame"] thead tr th:nth-child(2) {
        position: sticky;
        left: 80px;
        background-color: #ffffff;
        z-index: 3;
    }
    /* Hilangkan index di Streamlit Data Editor */
    div[data-testid="stDataFrame"] .st-bd {
        margin-left: -5px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("Olah Nilai Rapor / Griya Rapor")
st.write("---")

# --- Bagian Sidebar untuk Unggah Data Siswa ---
st.sidebar.header("Data Siswa")
uploaded_file = st.sidebar.file_uploader(
    "Unggah File CSV Data Siswa (Kolom wajib: NIS, Nama, Kelas)",
    type="csv"
)

# Menentukan data yang akan digunakan (Uploaded > Base CSV > Dummy)
if uploaded_file is not None:
    try:
        df_all_students_raw = pd.read_csv(uploaded_file)
        df_all_students_raw.columns = [col.strip() for col in df_all_students_raw.columns]
        df_all_students = df_all_students_raw.rename(columns={
            'NISN': 'NIS', 'NAMA': 'Nama', 'KLAS': 'Kelas', 'KLS': 'Kelas'
        }, errors='ignore')

        required_cols = ['NIS', 'Nama', 'Kelas']
        if all(col in df_all_students.columns for col in required_cols):
            df_all_students['NIS'] = df_all_students['NIS'].astype(str).str.strip()
            df_all_students['Nama'] = df_all_students['Nama'].astype(str).str.strip()
            df_all_students['Kelas'] = df_all_students['Kelas'].astype(str).str.strip().str.upper()

            for col in INPUT_SCORE_COLS:
                if col not in df_all_students.columns:
                    df_all_students[col] = 0.0

            for col in INPUT_SCORE_COLS:
                df_all_students[col] = pd.to_numeric(df_all_students[col], errors='coerce').fillna(0.0)

            df_all_students = df_all_students.drop(columns=['Nilai Rata-rata'], errors='ignore')

        else:
            st.sidebar.error("CSV yang diunggah harus memiliki kolom 'NIS', 'Nama', dan 'Kelas'. Menggunakan data dasar.")
            df_all_students = df_all_students_base
    except Exception as e:
        st.sidebar.error(f"Terjadi kesalahan saat memuat CSV yang diunggah: {e}. Menggunakan data dasar.")
        df_all_students = df_all_students_base
else:
    df_all_students = df_all_students_base

if len(df_all_students) > 0 and 'Nama' in df_all_students.columns:
    st.sidebar.info(f"Total data siswa yang dimuat: **{len(df_all_students)}**.")


st.header("⚙️ Pengaturan Form Nilai")

# --- Form Input Data Guru dan Kelas ---
col1, col2, col3 = st.columns(3)
with col1:
    mapel_terpilih = st.selectbox("Mata Pelajaran", COMMON_SUBJECTS)
with col2:
    available_classes = sorted(df_all_students['Kelas'].unique().tolist())
    if len(available_classes) == 0:
        available_classes = CLASS_OPTIONS

    kelas_input_list = st.multiselect("Pilih Kelas", available_classes, default=available_classes[:1])

with col3:
    semester_input = st.selectbox("Semester", ["Ganjil", "Genap"])

col4, col5, col6 = st.columns(3)
with col4:
    tahun_pelajaran_input = st.selectbox("Tahun Pelajaran", YEAR_OPTIONS, index=0)
with col5:
    guru_input = st.text_input("Nama Guru Mata Pelajaran", "")
with col6:
    nip_guru_input = st.text_input("NIP Guru", "")

st.write("---")
st.header("📝 Editor Nilai")

if not kelas_input_list:
    st.warning("Silakan pilih minimal satu kelas untuk memulai editor nilai.")
else:
    df_selected_classes = df_all_students[df_all_students['Kelas'].isin(kelas_input_list)].reset_index(drop=True)

    # 1. Hitung NR, Avg_TP, Avg_LM, Avg_PSA (Perhitungan Python)
    df_calculated = calculate_nr(df_selected_classes)
    # 2. Hitung Status TK (T/R)
    df_calculated = calculate_tk_status(df_calculated)
    # 3. Hitung Deskripsi NR
    df_calculated = generate_nr_description(df_calculated)

    DISPLAY_COLS = ['NIS', 'Nama', 'Kelas'] + INPUT_SCORE_COLS
    editor_display_cols = DISPLAY_COLS + [
        'Avg_TP', 'Avg_LM', 'Avg_PSA', 'NR', 'Deskripsi_NR'
    ] + [c for c in df_calculated.columns if c.startswith('TK_TP')]

    df_display = df_calculated[editor_display_cols].rename(columns=COLUMN_DISPLAY_MAP)

    float_display_cols = ['Rata-rata TP', 'Rata-rata LM', 'Rata-rata PSA'] + [COLUMN_DISPLAY_MAP.get(c, c) for c in INPUT_SCORE_COLS]
    for col in float_display_cols:
        if col in df_display.columns:
            # Nilai output di editor ditampilkan sebagai string 1 desimal
            df_display[col] = df_display[col].apply(lambda x: '' if x == 0.0 else f"{x:.1f}")

    column_config_map = {
        'NIS': st.column_config.TextColumn("NIS", disabled=True),
        'Nama': st.column_config.TextColumn("NAMA SISWA", disabled=True),
        'Kelas': st.column_config.TextColumn("KELAS", disabled=True),
    }

    # KONFIGURASI KOLOM INPUT DENGAN TEXTCOLUMN UNTUK MEMUNGKINKAN INPUT KOMA
    for col in INPUT_SCORE_COLS:
        column_config_map[COLUMN_DISPLAY_MAP.get(col, col)] = st.column_config.TextColumn( # <-- UBAH KE TEXTCOLUMN
            COLUMN_DISPLAY_MAP.get(col, col),
            help="Nilai antara 0-100 (Dapat diinput pecahan, koma akan otomatis diubah ke titik)",
            default=""
        )

    # Konfigurasi kolom hasil (disabled)
    for col in ['Rata-rata TP', 'Rata-rata LM', 'Rata-rata PSA', 'NR', 'Deskripsi Rapor'] + [COLUMN_DISPLAY_MAP.get(c, c) for c in df_calculated.columns if c.startswith('TK_TP')]:
        column_config_map[col] = st.column_config.TextColumn(col, disabled=True)

    st.subheader("Tabel Olah Nilai")
    st.info("Nilai di kolom **TP-n**, **LM-n**, **PTS**, dan **SAS/SAT** dapat diubah langsung. **Input koma (',') akan otomatis dikonversi menjadi titik ('.')** saat perhitungan.")

    edited_df = st.data_editor(
        df_display,
        column_config=column_config_map,
        hide_index=True,
        key="data_editor_nilai"
    )

    # --- Bagian Ekspor ---
    df_export_edited = edited_df.rename(columns={v: k for k, v in COLUMN_DISPLAY_MAP.items()})

    # Konversi kolom input kembali ke float, DENGAN MENGGANTI KOMA (,) MENJADI TITIK (.)
    for col in INPUT_SCORE_COLS:
        if col in df_export_edited.columns:
            # 1. Pastikan data adalah string, dan ganti koma dengan titik
            temp_series = df_export_edited[col].astype(str).str.replace(',', '.', regex=False)

            # 2. Konversi ke float, mengabaikan error (diisi 0.0)
            df_export_edited[col] = pd.to_numeric(temp_series, errors='coerce').fillna(0.0)

    # Recalculate based on edited data
    df_export_calculated = calculate_nr(df_export_edited)
    df_export_calculated = calculate_tk_status(df_export_calculated)
    df_export_calculated = generate_nr_description(df_export_calculated)

    col_form, col_tk = st.columns(2)

    with col_form:
        excel_buffer_nilai = export_multisheet_form_nilai(
            df_export_calculated,
            kelas_input_list,
            mapel_terpilih,
            semester_input,
            tahun_pelajaran_input,
            guru_input,
            nip_guru_input
        )
        st.download_button(
            label="⬇️ Unduh **Form Nilai/Omah Rapor** (bisa pilih banyak kelas)",
            file_name=f"Form_Nilai_Rapor_Multi_{mapel_terpilih}_{semester_input}.xlsx",
            data=excel_buffer_nilai,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_nilai"
        )

    with col_tk:
        excel_buffer_tk = export_multisheet_report_tk(
            df_export_calculated,
            kelas_input_list,
            mapel_terpilih,
            tahun_pelajaran_input
        )
        st.download_button(
            label="⬇️ Unduh **Nilai Rapor dan TK** (untuk diupload ke e-rapor)",
            file_name=f"Laporan_TK_Multi_{mapel_terpilih}_{semester_input}.xlsx",
            data=excel_buffer_tk,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_tk"
        )
