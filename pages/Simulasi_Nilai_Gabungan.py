import io
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from datetime import datetime

# --- 1. KONFIGURASI HALAMAN & CSS ---
st.set_page_config(page_title="Portal Akademik SMPN 2 Banguntapan", layout="wide")

def local_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Quicksand:wght@500;700&display=swap');

    .running-text {
        font-family: 'Quicksand', sans-serif;
        font-size: 14px; color: #3E584A; background-color: #E8F0E8;
        padding: 8px 0; font-weight: bold; margin-top: -50px;
        margin-bottom: 20px; border-bottom: 1px solid #D1DBD1;
    }

    .stApp { background-color: #F7F9F7; color: #34495E; }
    [data-testid="stSidebar"] { background-color: #E8F0E8 !important; border-right: 1px solid #D1DBD1; }
    
    [data-testid="stVerticalBlock"] > div:has(div.element-container) {
        background: white; border-radius: 12px; padding: 15px 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03); border: 1px solid #E0E7E0; margin-bottom: 8px;
    }

    h1, h2, h3 { color: #3E584A !important; }
    
    /* Warna Metric (Poin) */
    [data-testid="stMetricValue"] {
        color: #4F7942 !important;
        font-size: 24px !important;
        font-weight: 700;
    }

    .stButton>button { 
        border-radius: 8px; background-color: #6B8E7B; color: white; 
        font-weight: 600; border: none; width: 100%;
    }
    </style>
    """, unsafe_allow_html=True)

local_css()

# --- FUNGSI GENERATE PDF ---
def create_pdf(user, detail_data, nilai_akhir):
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=LETTER)
    w, h = LETTER
    p.setFont("Helvetica-Bold", 16)
    p.drawCentredString(w/2, h - 20*mm, "LAPORAN SIMULASI NILAI GABUNGAN")
    p.setFont("Helvetica", 12)
    p.drawCentredString(w/2, h - 26*mm, "SMP NEGERI 2 BANGUNTAPAN")
    p.line(20*mm, h - 32*mm, w - 20*mm, h - 32*mm)
    p.setFont("Helvetica-Bold", 11)
    p.drawString(20*mm, h - 45*mm, f"Nama: {user['Nama Siswa']}")
    p.drawString(20*mm, h - 51*mm, f"NIS: {user['NIS']}")
    p.setFont("Helvetica", 10)
    y = h - 70*mm
    for d in detail_data:
        p.drawString(22*mm, y, f"{d['Mata Pelajaran']} - Rerata: {d['Rerata']}, TKA/D: {d['TKA/D']}")
        y -= 7*mm
    p.setFont("Helvetica-Bold", 14)
    p.drawCentredString(w/2, y - 10*mm, f"NILAI AKHIR: {nilai_akhir:.2f}")
    p.showPage()
    p.save()
    buffer.seek(0)
    return buffer

# --- RUNNING TEXT ---
st.markdown("""<div class="running-text"><marquee scrollamount="8">✨ Rumus Nilai Gabungan = ((Nilai TKA + TKAD) x 60%) + (Jumlah Rerata Nilai Rapor Semester 1-5 x 40%) ✨</marquee></div>""", unsafe_allow_html=True)

# --- SESSION STATE ---
if 'db_siswa' not in st.session_state: st.session_state.db_siswa = None
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
MAPEL_UTAMA = ["Bahasa Indonesia", "Matematika", "Bahasa Inggris", "IPA"]

# --- NAVIGASI ---
menu = st.sidebar.selectbox("📂 MENU UTAMA", ["Home / Login", "Admin Upload"])

if menu == "Admin Upload":
    st.title("📂 Admin Control")
    pwd = st.text_input("Password", type="password")
    if pwd == "admin123":
        uploaded = st.file_uploader("Upload Excel", type=["xlsx"])
        if uploaded:
            df = pd.read_excel(uploaded)
            df.columns = df.columns.str.strip()
            df["NIS"] = df["NIS"].astype(str).str.strip()
            st.session_state.db_siswa = df
            st.success("Database Terupdate!")
else:
    if not st.session_state.logged_in:
        st.title("🏛️ Portal Simulasi")
        nis_in = st.text_input("MASUKKAN NIS")
        if st.button("LOGIN"):
            if st.session_state.db_siswa is not None:
                match = st.session_state.db_siswa[st.session_state.db_siswa["NIS"] == nis_in.strip()]
                if not match.empty:
                    st.session_state.logged_in = True
                    st.session_state.user_data = match.iloc[0].to_dict()
                    st.rerun()
                else: st.error("NIS tidak ditemukan.")
    else:
        user = st.session_state.user_data
        st.title(f"🏫 Profil: {user['Nama Siswa']}")
        if st.sidebar.button("Logout"):
            st.session_state.logged_in = False
            st.rerun()

        col_in, col_res = st.columns([1, 2])
        sim_tkad = {}
        with col_in:
            st.subheader("📝 Input TKA/D")
            for m in MAPEL_UTAMA:
                sim_tkad[f"{m}_TKAD"] = st.number_input(f"{m}", 0.0, 100.0, 0.0, step=0.01, format="%.2f", key=f"in_{m}")

        with col_res:
            total_rerata = 0
            detail_data = []
            for m in MAPEL_UTAMA:
                v = [user[f"{m}_S{i}"] for i in range(1,6)]
                avg = sum(v)/5
                total_rerata += avg
                detail_data.append({
                    "Mata Pelajaran": m, "S1":int(v[0]), "S2":int(v[1]), "S3":int(v[2]), "S4":int(v[3]), "S5":int(v[4]),
                    "Rerata": f"{avg:.2f}", "TKA/D": f"{sim_tkad[f'{m}_TKAD']:.2f}"
                })
            
            total_tkad = sum(sim_tkad.values())
            
            # --- BAGIAN POIN YANG KEMBALI DITAMBAHKAN ---
            poin_rapor = total_rerata * 0.4
            poin_tkad = total_tkad * 0.6
            nilai_akhir = poin_rapor + poin_tkad

            m1, m2 = st.columns(2)
            m1.metric("Poin Rapor (40%)", f"{poin_rapor:.2f}")
            m2.metric("Poin TKA/D (60%)", f"{poin_tkad:.2f}")
            # --------------------------------------------

            st.markdown(f"""<div style="background:#E8F5E9;padding:15px;border-radius:12px;border:1px solid #A5D6A7;text-align:center;margin-bottom:10px;">
                <p style="margin:0; font-size:12px; font-weight:bold; color:#2E7D32;">TOTAL NILAI AKHIR GABUNGAN</p>
                <h1 style="font-size:50px !important;color:#1B5E20 !important;margin:0;">{nilai_akhir:.2f}</h1></div>""", unsafe_allow_html=True)
            
            with st.expander("🔍 Rincian Nilai", expanded=True):
                st.table(pd.DataFrame(detail_data))

            pdf_file = create_pdf(user, detail_data, nilai_akhir)
            st.download_button(label="🖨️ UNDUH LAPORAN (PDF)", data=pdf_file, file_name=f"Simulasi_{user['NIS']}.pdf", mime="application/pdf")
