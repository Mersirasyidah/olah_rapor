import streamlit as st

# ================= KONFIGURASI HALAMAN =================
st.set_page_config(
    page_title="Sistem Informasi Sekolah",
    layout="wide",
    page_icon="🎓"
)

# ================= CSS KUSTOM =================
st.markdown("""
    <style>
    /* Global Styling */
    .stApp {
        background: linear-gradient(135deg, #87CEFA, #e0f7ff); /* Sky Blue Gradient */
        font-family: "Segoe UI", sans-serif;
    }

    /* Header Styling */
    .title {
        text-align: center;
        font-size: 30px;
        font-weight: 800;
        color: #003366;
        margin-bottom: -5px;
        letter-spacing: 0.5px;
        font-family: 'Oswald', Bebas Neue;
    }
    .subtitle {
        text-align: center;
        font-size: 16px;
        font-weight: 500;
        color: #555;
        margin-bottom: 25px;
    }

    /* Hide original st.page_link label */
    div[data-testid="stPageLink-container"] p {
        display: none;
    }

    /* Menu Card Styling */
    .menu-card-container {
        display: flex;
        justify-content: center; /* Memastikan kartu berada di tengah */
        gap: 30px; /* Jarak antar kartu */
    }
    .menu-card {
        background: white;
        width: 200px;
        padding: 20px 15px;
        border-radius: 14px;
        text-align: center;
        font-size: 15px;
        font-weight: 600;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        transition: transform 0.25s ease, box-shadow 0.25s ease;
        cursor: pointer;
        display: block; /* Penting untuk tautan */
        text-decoration: none; /* Menghilangkan underline pada teks tautan */
        color: inherit;
    }
    .menu-card:hover {
        transform: translateY(-6px);
        box-shadow: 0 10px 18px rgba(0,0,0,0.2);
    }
    .icon-circle {
        width: 48px;
        height: 48px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto 10px auto;
        font-size: 22px;
        color: white;
    }
    .blue { background: #4facfe; }
    .green { background: #43e97b; }
    .orange { background: #f6d365; }
    .purple { background: #a18cd1; }
    .gold { background: #FFD700; }
    .red { background: #FF4B4B; }

    /* Info Box Styling (Untuk teks menarik) */
    .info-box {
        background-color: rgba(255, 255, 255, 0.9);
        padding: 25px;
        border-radius: 12px;
        margin: 30px 0;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        text-align: center;
        max-width: 800px;
        margin-left: auto;
        margin-right: auto;
        border-left: 5px solid #0056b3;
    }
    .info-box h2 {
        color: #0056b3;
        margin-top: 0;
        font-size: 24px;
        font-weight: 700;
    }
    .info-box p {
        color: #333;
        font-size: 16px;
        line-height: 1.6;
    }

    /* Footer Styling */
    .footer {
        text-align: center;
        font-size: 13px;
        margin-top: 20px;
        padding-top: 15px;
        border-top: 1px solid #ddd;
        color: #666;
    }
    </style>
""", unsafe_allow_html=True)

# ================= HEADER =================
# Menggunakan kolom untuk tata letak header dengan logo
col1, col2, col3 = st.columns([10, 100, 10])

with col1:
    # Ganti dengan path ke logo kiri Anda
    st.image("assets/logo_kiri.png", width=70)

with col2:
    st.markdown("<h1 class='title'>SISTEM INFORMASI SEKOLAH</h1>", unsafe_allow_html=True)
    st.markdown("<h3 class='subtitle'>SMP NEGERI 2 BANGUNTAPAN</h3>", unsafe_allow_html=True)

with col3:
    # Ganti dengan path ke logo kanan Anda
    st.image("assets/logo_kanan.png", width=70)

# ================= TEKS PEMBUKA MENARIK =================
st.markdown("""
    <div class="info-box">
        <h2>🎓 Sistem Informasi Sekolah: Pusat Data Terpadu</h2>
        <p>
            Selamat datang di <b>Dashboard Utama</b> SMP Negeri 2 Banguntapan.<br> Sistem ini adalah portal digital, untuk mengelola dan memantau seluruh aktivitas akademik siswa secara <br><b>efektif dan efisien</b>. Dengan sistem ini, dapat menyamakan dan memudahkan proses pengeditan serta pengelolaan data siswa untuk keperluan informasi akademik. Gunakan kartu menu di bawah ini untuk mengakses data <b>Daftar Nama, Nilai, Absensi, dan Hasil Ujian</b> dengan cepat.
        </p>
    </div>
""", unsafe_allow_html=True)

st.write("")

# ================= MENU HORIZONTAL (2 Baris, 3 Kolom per Baris) =================

# ------------------- Baris Pertama (3 Kolom) -------------------
col1, col2, col3 = st.columns(3)

# Menu 1: Daftar Nama Siswa (Masuk ke col1)
with col1:
    st.page_link("pages/1_Daftar_Nama.py", label="📋 Daftar Nama Siswa")
    st.markdown("""
        <div class='menu-card'>
            <div class='icon-circle blue'>📋</div>
            Daftar Nama Siswa
        </div>
    """, unsafe_allow_html=True)

# Menu 2: Daftar Nilai Siswa (Masuk ke col2)
with col2:
    st.page_link("pages/2_Daftar_Nilai.py", label="📝 Daftar Nilai Siswa")
    st.markdown("""
        <div class='menu-card'>
            <div class='icon-circle green'>📝</div>
            Daftar Nilai Siswa
        </div>
    """, unsafe_allow_html=True)

# Menu 3: Daftar Absensi (Masuk ke col3)
with col3:
    st.page_link("pages/3_Daftar_Absensi.py", label="📊 Daftar Absensi")
    st.markdown("""
        <div class='menu-card'>
            <div class='icon-circle orange'>📊</div>
            Daftar Absensi
        </div>
    """, unsafe_allow_html=True)

# ==================== MENAMBAHKAN JARAK (SPASI) ANTAR BARIS ====================
st.markdown("<br><br>", unsafe_allow_html=True) # Menggunakan HTML <br> untuk 2 baris spasi
# Atau bisa juga:
# st.write("")
# st.write("")
# Atau pemisah visual yang lebih kuat:
# st.markdown("---")
# ==============================================================================

# ------------------- Baris Kedua (3 Kolom) -------------------
# Panggilan st.columns(3) yang baru akan membuat baris baru setelah spasi di atas.
col4, col5, col6 = st.columns(3)

# Menu 4: Nilai Hasil Ujian (1) (Masuk ke col4)
with col4:
    st.page_link("pages/4_Hasil_Ujian.py", label="🎓 Nilai Hasil Ujian")
    st.markdown("""
        <div class='menu-card'>
            <div class='icon-circle gold'>🎓</div>
            Nilai Hasil Ujian
        </div>
    """, unsafe_allow_html=True)

# Menu 5: Nilai Hasil Ujian (2) (Masuk ke col5)
with col5:
    st.page_link("pages/5_Hasil_TO.py", label="📚 Nilai Hasil TO TKA/TKAD")
    st.markdown("""
        <div class='menu-card'>
            <div class='icon-circle purple'>📚</div>
            Nilai Hasil TO TKA/TKAD
        </div>
    """, unsafe_allow_html=True)

# Menu 6: Nilai Hasil Ujian (3) (Masuk ke col6)
with col6:
    st.page_link("pages/6_Olah_Nilai_TP.py", label="🏆 Nilai Hasil Olah Rapor")
    st.markdown("""
        <div class='menu-card'>
            <div class='icon-circle red'>🏆</div>
            Nilai Hasil Olah Rapor
        </div>
    """, unsafe_allow_html=True)


# ================= FOOTER DI TENGAH DENGAN JARAK ATAS YANG BESAR =================
st.markdown(
    """
    <div style="
        text-align: right;
        padding: 10px;
        font-size: 14px;
        color: black;
        /* Tambahkan jarak di atas (misalnya 100 piksel) */
        margin-top: 200px;
    ">
        © 2025 **SMP Negeri 2 Banguntapan** | Dibuat oleh <b>Mersi</b>
    </div>
    """,
    unsafe_allow_html=True
)

