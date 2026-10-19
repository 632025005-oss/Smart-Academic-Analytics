# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import hashlib
import io
import zipfile
import base64
from openpyxl import Workbook
from streamlit_option_menu import option_menu
import pytz

from db_storage import (
    load_users as _db_load_users, save_users as _db_save_users,
    load_database, insert_siswa, hapus_semua_siswa,
    load_config, save_config, load_log, save_log, tambah_log,
    update_kuesioner_siswa, load_all_siswa, hapus_siswa_by_id,
    hapus_semua_siswa_admin, cek_status_kuesioner, update_user_password
)

SUPER_ADMIN = {"admin", "regina"}

# ================================================================
# HELPER & AUTH
# ================================================================
def get_wib_time():
    tz = pytz.timezone('Asia/Jakarta')
    return datetime.now(tz)

def get_secret(section, key=None, default=None):
    try:
        if key is None: return st.secrets[section]
        return st.secrets[section][key]
    except (KeyError, FileNotFoundError):
        return default

def hash_password(p): return hashlib.sha256(p.encode()).hexdigest()
def verify_password(p, hashed): return hash_password(p) == hashed

def load_users():
    users = _db_load_users()
    if not users:
        super_admins = get_secret("super_admin", None, {})
        users = {}
        for username, info in super_admins.items():
            users[username] = {
                "password": hash_password(info["password"]),
                "nama":     info.get("nama", username.capitalize()),
                "role":     "admin", "kelas_ampu": None, "email": info.get("email", ""),
                "status":   "approved", "created":  get_wib_time().strftime("%Y-%m-%d %H:%M"),
            }
        _db_save_users(users)
    return users

def save_users(u): _db_save_users(u)

# ================================================================
# PAGE CONFIG & CSS
# ================================================================
st.set_page_config(page_title="SAA Analytics", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&family=Caveat:wght@600;700&display=swap');

:root {
    --teal: #0D9488; --teal-dark: #0F766E; --gold: #F59E0B;
    --bg: #F4F7F6; --text-main: #1E293B; --text-muted: #64748B; --card-bg: #FFFFFF;
}

html, body, [class*="css"] { font-family: 'Nunito', sans-serif; color: var(--text-main); }
.stApp { background-color: var(--bg); }
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1400px; }

#MainMenu, footer { visibility: hidden; }

/* =========================================
   DESAIN SIDEBAR & MENU
   ========================================= */
[data-testid="stSidebar"] {
    background-color: #006B62 !important;
    background-image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none"><path d="M0,60 Q30,90 60,60 T100,60 L100,100 L0,100 Z" fill="%23004d45"/></svg>') !important;
    background-repeat: no-repeat !important;
    background-position: bottom !important;
    background-size: 100% 300px !important;
    border-right: none;
}
[data-testid="stSidebar"] * { color: white !important; }

.sb-logo { display: flex; align-items: center; gap: 15px; padding: 20px 20px 10px 20px; }
.sb-logo-icon { display: flex; align-items: center; justify-content: center; }
.sb-logo-text h2 { margin: 0; font-size: 1.6rem; font-weight: 900; letter-spacing: -0.5px; }

.sb-bottom { position: relative; padding: 20px; margin-top: 30px; text-align: center; }
.sb-quote { font-family: 'Caveat', cursive; font-size: 1.6rem; line-height: 1.2; color: #FDE68A; transform: rotate(-5deg); margin-bottom: 20px; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);}
.sb-img { width: 100%; max-width: 200px; margin: 0 auto; display: block; }

/* =========================================
   DESAIN KONTEN UTAMA
   ========================================= */
.top-header {
    display: flex; justify-content: space-between; align-items: center;
    background: white; padding: 1.5rem 2rem; border-radius: 16px;
    margin-bottom: 1.5rem; box-shadow: 0 4px 20px rgba(13, 148, 136, 0.05);
    border: 1px solid #E2E8F0;
}
.top-header h1 { margin: 0; font-size: 1.6rem; font-weight: 900; color: var(--teal-dark); }
.top-header p { margin: 0; color: var(--text-muted); font-size: 0.9rem; font-weight: 600;}

.kpi-card {
    background: white; border-radius: 16px; padding: 1.2rem;
    box-shadow: 0 4px 15px rgba(0,0,0,0.03); display: flex; align-items: center;
    gap: 1rem; border: 1px solid #E2E8F0; margin-bottom: 1rem;
}
.kpi-icon {
    width: 50px; height: 50px; border-radius: 12px; display: flex; align-items: center; justify-content: center; font-size: 1.5rem; flex-shrink: 0;
}
.kpi-icon.blue { background: #E0F2FE; color: #0284C7; }
.kpi-icon.green { background: #DCFCE7; color: #16A34A; }
.kpi-icon.orange { background: #FEF3C7; color: #D97706; }
.kpi-icon.purple { background: #F3E8FF; color: #9333EA; }

.kpi-info h4 { margin: 0; font-size: 0.75rem; color: var(--text-muted); font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;}
.kpi-info h2 { margin: 0; font-size: 1.8rem; color: var(--gold); font-weight: 900; line-height: 1.2;}
.kpi-info p { margin: 0; font-size: 0.7rem; color: var(--text-muted); font-weight: 600;}

.module-card {
    background: white; border-radius: 12px; padding: 1.2rem; border: 1px solid #E2E8F0;
    transition: all 0.3s ease; min-height: 110px; margin-bottom: 0.5rem;
    display: flex; flex-direction: column; justify-content: flex-start; box-shadow: 0 2px 5px rgba(0,0,0,0.02);
}
.module-card:hover { transform: translateY(-3px); box-shadow: 0 8px 20px rgba(13, 148, 136, 0.1); border-color: var(--teal); }
.module-icon { font-size: 1.8rem; margin-bottom: 0.5rem;}
.module-title { font-weight: 800; font-size: 1rem; color: #1E293B; margin-bottom: 0.2rem;}
.module-desc { font-size: 0.75rem; color: #64748B; line-height: 1.4;}

.stButton > button { background: var(--teal); color: white; border-radius: 8px; font-weight: 800; width: 100%; border: none; padding: 0.5rem;}
.stButton > button:hover { background: var(--teal-dark); color: white; border: none;}
.stDataFrame { border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.03); }

/* =========================================
   MEDIA QUERIES (RESPONSIVE DESIGN)
   ========================================= */
@media (min-width: 769px) {
    header { visibility: hidden; }
    [data-testid="collapsedControl"] { display: none; } 
    [data-testid="stSidebar"] {
        min-width: 280px !important; max-width: 280px !important;
        transform: translateX(0) !important; visibility: visible !important;
    }
}

@media (max-width: 768px) {
    header { visibility: visible !important; background-color: transparent !important; }
    [data-testid="collapsedControl"] { display: flex !important; visibility: visible !important; color: #0D9488 !important; }
    [data-testid="stSidebar"] { min-width: 250px !important; max-width: 250px !important; }
    .block-container { padding-left: 1rem !important; padding-right: 1rem !important; }
    
    .top-header { flex-direction: column; align-items: flex-start; gap: 10px; padding: 1.2rem; }
    .top-header h1 { font-size: 1.3rem; }
    .top-header h2 { font-size: 1.2rem; }

    .kpi-card { padding: 1rem; gap: 0.8rem; }
    .kpi-icon { width: 40px; height: 40px; font-size: 1.2rem; }
    .kpi-info h2 { font-size: 1.5rem; }
    .module-card { padding: 1rem; min-height: auto; }
}

/* =========================================
   FIX TOTAL: BACKGROUND MENU JADI HIJAU GELAP
   ========================================= */
.stCustomComponentV1 {
    background-color: transparent !important;
}
[data-testid="stSidebar"] ul.nav-pills {
    background-color: #006B62 !important;
}
[data-testid="stSidebar"] ul.nav-pills a.nav-link {
    background-color: #006B62 !important;
    color: white !important;
}
[data-testid="stSidebar"] ul.nav-pills a.nav-link.active {
    background-color: #004d45 !important;
    color: white !important;
    box-shadow: 0 4px 10px rgba(0,0,0,0.2);
}
</style>
""", unsafe_allow_html=True)

# ================================================================
# STATE MANAGEMENT
# ================================================================
for key, val in {
    "logged_in": False, "user_nama": None, "user_role": None, 
    "database_siswa": [], "config": None, "log": [], "users": {},
    "current_page": "home", "last_saved": None
}.items():
    if key not in st.session_state: st.session_state[key] = val

def reload_data():
    if st.session_state.user_nama:
        st.session_state.database_siswa = load_database(st.session_state.user_nama)
        st.session_state.log = load_log(user=st.session_state.user_nama) 
    else:
        st.session_state.log = []
    st.session_state.config = load_config()
    st.session_state.users = load_users()

if st.session_state.config is None:
    reload_data()

def add_log(aksi, detail=""):
    tambah_log({"waktu": get_wib_time().strftime("%Y-%m-%d %H:%M:%S"), "user": st.session_state.user_nama or "-", "role": st.session_state.user_role or "-", "aksi": aksi, "detail": detail})

def goto_page(page):
    st.session_state.current_page = page
    if "sidebar_menu" in st.session_state: 
        del st.session_state["sidebar_menu"]
    st.rerun()

# ================================================================
# KONSTANTA & FUNGSI ANALITIK
# ================================================================
BASELINE_ASPEK = {"Self-Efficacy Akademik": 4.634, "Keterlibatan Orang Tua": 4.348, "Harapan Orang Tua": 3.450, "Dukungan Sekolah": 4.600, "Motivasi Belajar": 2.517, "Kecemasan Akademik": 2.893, "Kemalasan Belajar": 1.487, "Fasilitas Sekolah": 4.880}
PENGARUH_DATA = {"Self-Efficacy Akademik": 2.1995, "Keterlibatan Orang Tua": 1.7559, "Harapan Orang Tua": -0.6987, "Dukungan Sekolah": -1.0351, "Motivasi Belajar": -1.9093, "Kecemasan Akademik": 0.7197, "Kemalasan Belajar": 0.7320, "Fasilitas Sekolah": 0.6182}
KEPENTINGAN_DATA = {"Self-Efficacy Akademik": 0.8197, "Keterlibatan Orang Tua": 0.6717, "Harapan Orang Tua": 0.5766, "Dukungan Sekolah": 0.4325, "Motivasi Belajar": 0.3601, "Kecemasan Akademik": 0.1553, "Fasilitas Sekolah": 0.0654, "Kemalasan Belajar": 0.0846}
BOBOT_PENGARUH = {k: PENGARUH_DATA[k] * KEPENTINGAN_DATA[k] for k in PENGARUH_DATA}
RATA_RATA_NILAI = 83.78

KELAS_LIST = [f"VII-{x}" for x in "ABCDEFGH"] + [f"VIII-{x}" for x in "ABCDEFGH"] + [f"IX-{x}" for x in "ABCDEFGH"]

SKALA_PILIHAN = {
    "Self-Efficacy Akademik": {"q": "Seberapa yakin siswa terhadap kemampuan akademiknya?", "o": [("1","Hampir tidak pernah"),("2","Jarang"),("3","Kadang-kadang"),("4","Sering"),("5","Hampir selalu")]},
    "Keterlibatan Orang Tua": {"q": "Seberapa aktif orang tua siswa mendampingi belajarnya?", "o": [("1","Hampir tidak pernah"),("2","Jarang"),("3","Kadang-kadang"),("4","Sering"),("5","Hampir selalu")]},
    "Harapan Orang Tua": {"q": "Seberapa tinggi tuntutan orang tua terhadap nilai siswa?", "o": [("1","Sangat rendah"),("2","Rendah"),("3","Sedang"),("4","Tinggi"),("5","Sangat tinggi")]},
    "Dukungan Sekolah": {"q": "Seberapa besar dukungan sekolah kepada siswa?", "o": [("1","Sangat kurang"),("2","Kurang"),("3","Cukup"),("4","Baik"),("5","Sangat baik")]},
    "Motivasi Belajar": {"q": "Seberapa besar motivasi belajar siswa?", "o": [("1","Tidak ada"),("2","Kurang"),("3","Cukup"),("4","Baik"),("5","Sangat baik")]},
    "Kecemasan Akademik": {"q": "Seberapa sering siswa cemas saat belajar/ujian?", "o": [("1","Hampir tidak pernah"),("2","Jarang"),("3","Kadang-kadang"),("4","Sering"),("5","Hampir selalu")]},
    "Fasilitas Sekolah": {"q": "Seberapa memadai fasilitas belajar siswa?", "o": [("1","Sangat kurang"),("2","Kurang"),("3","Cukup"),("4","Memadai"),("5","Sangat memadai")]},
    "Kemalasan Belajar": {"q": "Seberapa sering siswa malas belajar?", "o": [("1","Tidak pernah"),("2","Jarang"),("3","Kadang-kadang"),("4","Sering"),("5","Hampir selalu")]}
}

def analisis_kausal(nilai_akademik, profil_siswa):
    selisih_nilai = nilai_akademik - RATA_RATA_NILAI
    kontribusi = {}
    for aspek, nilai_input in profil_siswa.items():
        selisih_aspek = nilai_input - BASELINE_ASPEK[aspek]
        kontribusi[aspek] = selisih_aspek * BOBOT_PENGARUH[aspek]
    total = sum(kontribusi.values())
    if abs(total) > 0.01:
        skala = selisih_nilai / total
        kontribusi = {k: v * skala for k, v in kontribusi.items()}
    return selisih_nilai, kontribusi

def kategori_nilai(n):
    if n >= 88: return "Sangat Baik", "#10B981"
    if n >= 84: return "Baik", "#2563EB"
    if n >= 80: return "Cukup", "#F0B900"
    if n == 0: return "Belum Dinilai", "#94A3B8"
    return "Perlu Perhatian", "#EF5B67"

def input_kategori(key, fname, is_student=False):
    cfg = SKALA_PILIHAN[fname]
    pertanyaan = cfg['q'] if not is_student else cfg['q'].replace("siswa", "kamu").replace("nya", "mu").replace("belajarnya", "belajarmu").replace("nilainya", "nilaimu")
    
    st.markdown(f"**{fname}** — {pertanyaan}")
    labels = [o[1] for o in cfg["o"]]
    vals = [float(o[0]) for o in cfg["o"]]
    default_idx = min(range(len(vals)), key=lambda i: abs(vals[i] - BASELINE_ASPEK[fname]))
    p = st.radio(fname, labels, index=default_idx, key=key, label_visibility="collapsed")
    return vals[labels.index(p)]

# ================================================================
# HALAMAN LOGIN
# ================================================================
def halaman_login():
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br><h1 style='text-align:center; color:#0D9488; font-weight:900;'>🎓 SAA Analytics</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align:center; color:#64748B; font-weight:600;'>Silakan masuk untuk mengakses dashboard</p>", unsafe_allow_html=True)
        
        tab_login, tab_daftar, tab_siswa = st.tabs(["🔐 Masuk", "📝 Daftar Akun", "🎒 Kuesioner Siswa"])

        with tab_login:
            with st.form("login_form"):
                username = st.text_input("Username")
                password = st.text_input("Password", type="password")
                submit = st.form_submit_button("Masuk", use_container_width=True)
                if submit:
                    users = load_users()
                    if username in users and verify_password(password, users[username]["password"]):
                        if users[username].get("status") == "pending":
                            st.warning("Akun belum disetujui Admin.")
                        else:
                            st.session_state.logged_in = True
                            st.session_state.user_nama = users[username]["nama"]
                            st.session_state.user_role = users[username]["role"]
                            
                            st.session_state.current_page = "home"
                            if "sidebar_menu" in st.session_state: del st.session_state["sidebar_menu"]
                            
                            reload_data()
                            add_log("Login", username)
                            st.rerun()
                    else:
                        st.error("Username atau password salah.")

        with tab_daftar:
            with st.form("register_form"):
                reg_nama = st.text_input("Nama Lengkap")
                reg_user = st.text_input("Username Baru")
                reg_pw = st.text_input("Password Baru", type="password")
                reg_submit = st.form_submit_button("Daftar")
                if reg_submit:
                    users = load_users()
                    if reg_user in users: st.error("Username sudah dipakai.")
                    elif len(reg_pw) < 6: st.error("Password minimal 6 karakter.")
                    else:
                        users[reg_user] = {"password": hash_password(reg_pw), "nama": reg_nama, "role": "guru", "status": "pending", "created": get_wib_time().strftime("%Y-%m-%d %H:%M")}
                        save_users(users)
                        st.success("Pendaftaran berhasil! Menunggu persetujuan Admin.")

        with tab_siswa:
            st.info("Isi kuesioner dengan jujur. Data akan masuk ke sistem guru Anda.")
            with st.expander("🔍 Cek Status Pengisian"):
                cek_nama = st.text_input("Nama Lengkap", key="cek_nama")
                cek_kelas = st.selectbox("Kelas", KELAS_LIST, key="cek_kelas")
                cek_absen = st.number_input("Absen", 1, 50, 1, key="cek_absen")
                if st.button("Cek Status"):
                    sudah, pesan = cek_status_kuesioner(cek_nama, cek_kelas, cek_absen)
                    if sudah: st.success(f"✅ Anda sudah mengisi kuesioner pada {pesan}.")
                    else: st.warning(f"❌ {pesan}")
            st.markdown("---")
            
            with st.form("form_kuesioner"):
                k_nama = st.text_input("Nama Lengkap")
                k_kelas = st.selectbox("Kelas", KELAS_LIST)
                k_absen = st.number_input("Absen", 1, 50, 1)
                st.markdown("---")
                f1 = input_kategori("k_s1", "Self-Efficacy Akademik", True)
                f2 = input_kategori("k_s2", "Motivasi Belajar", True)
                f3 = input_kategori("k_s3", "Kecemasan Akademik", True)
                f4 = input_kategori("k_s4", "Kemalasan Belajar", True)
                f5 = input_kategori("k_s5", "Keterlibatan Orang Tua", True)
                f6 = input_kategori("k_s6", "Harapan Orang Tua", True)
                f7 = input_kategori("k_s7", "Dukungan Sekolah", True)
                f8 = input_kategori("k_s8", "Fasilitas Sekolah", True)
                
                sub_siswa = st.form_submit_button("Kirim Kuesioner", type="primary", use_container_width=True)
                if sub_siswa:
                    if not k_nama.strip(): st.error("Nama wajib diisi!")
                    else:
                        profil = {"Self-Efficacy": f1, "Keterlibatan Ortu": f5, "Harapan Ortu": f6, "Dukungan Sekolah": f7, "Motivasi": f2, "Kecemasan": f3, "Fasilitas": f8, "Kemalasan": f4}
                        ok, msg = update_kuesioner_siswa(k_nama, k_kelas, k_absen, profil)
                        if ok: st.success("🎉 Terima kasih! Data kamu sudah tersimpan. Terus semangat belajar!"); st.balloons()
                        else: st.error(msg)

# ================================================================
# ROUTING & SIDEBAR
# ================================================================
if not st.session_state.logged_in:
    halaman_login()
    st.stop()

with st.sidebar:
    def get_base64_image(image_path):
        try:
            with open(image_path, "rb") as img_file:
                return base64.b64encode(img_file.read()).decode()
        except FileNotFoundError:
            return ""

    logo_b64 = get_base64_image("logo_sekolah.png")
    
    if logo_b64:
        logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="width: 45px; height: 45px; object-fit: contain; background-color: white; border-radius: 8px; padding: 2px;">'
    else:
        logo_html = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="white" width="45px" height="45px"><path d="M12 3L1 9l11 6 9-4.91V17h2V9L12 3zm0 10.91L5 13.18l7-3.82 7 3.82-7 3.73zM5 13.18v4L12 21l7-3.82v-4L12 17l-7-3.82z"/></svg>'

    st.markdown(f"""
    <div class="sb-logo">
        <div class="sb-logo-icon">
            {logo_html}
        </div>
        <div class="sb-logo-text">
            <h2>SAA.Analytics</h2>
        </div>
    </div>
    <hr style="border-color: rgba(255,255,255,0.1); margin: 10px 20px;">
    """, unsafe_allow_html=True)

    menu_options = ["Beranda", "Input & Upload", "Data Siswa", "Pengaruh Faktor", "Laporan", "Prediksi", "Tentang"]
    menu_icons = ["house-door", "pencil-square", "people", "search", "file-earmark-text", "graph-up-arrow", "info-circle"]
    
    if st.session_state.user_role == "admin":
        menu_options.append("Pengaturan Admin")
        menu_icons.append("gear")
        menu_options.append("Logout")
        menu_icons.append("box-arrow-right")
    else:
        menu_options.append("Pengaturan")
        menu_icons.append("gear")
        menu_options.append("Logout")
        menu_icons.append("box-arrow-right")

    mapping_menu = {
        "Beranda": "home", "Input & Upload": "analisis", "Data Siswa": "monitoring", 
        "Pengaruh Faktor": "kausal", "Laporan": "rekomendasi", "Prediksi": "prediksi", 
        "Tentang": "tentang", "Pengaturan": "pengaturan", "Pengaturan Admin": "admin_users"
    }
    
    current_page_name = "Beranda"
    for k, v in mapping_menu.items():
        if v == st.session_state.current_page:
            current_page_name = k
            break
            
    default_idx = menu_options.index(current_page_name) if current_page_name in menu_options else 0

    menu_choice = option_menu(
        menu_title=None,
        options=menu_options,
        icons=menu_icons,
        default_index=default_idx,
        key="sidebar_menu",
        styles={
            "container": {"padding": "0!important", "background-color": "#006B62 !important"},
            "icon": {"color": "white", "font-size": "18px"}, 
            "nav-link": {
                "font-size": "15px", 
                "text-align": "left", 
                "margin": "0px", 
                "--hover-color": "#005f57", 
                "color": "white !important", 
                "font-weight": "600", 
                "background-color": "#006B62 !important"
            },
            "nav-link-selected": {
                "background-color": "#004d45 !important", 
                "color": "white !important", 
                "font-weight": "700"
            },
        }
    )

    if menu_choice == "Logout":
        add_log("Logout")
        st.session_state.logged_in = False
        st.session_state.user_nama = None
        
        st.session_state.user_role = None
        st.session_state.current_page = "home"
        if "sidebar_menu" in st.session_state: del st.session_state["sidebar_menu"]
        
        st.rerun()
    else:
        new_page = mapping_menu.get(menu_choice)
        if new_page and st.session_state.current_page != new_page:
            st.session_state.current_page = new_page
            st.rerun()

    st.markdown("""
    <div class="sb-bottom">
        <div class="sb-quote">Belajar<br>Hari Ini<br>Untuk Masa Depan<br>yang Lebih Baik ❤️</div>
        <img src="https://cdn-icons-png.flaticon.com/512/3145/3145765.png" class="sb-img" style="filter: drop-shadow(0px 10px 10px rgba(0,0,0,0.5));">
    </div>
    """, unsafe_allow_html=True)


# ================================================================
# HALAMAN BERANDA
# ================================================================
if st.session_state.current_page == "home":
    st.markdown(f"""
    <div class="top-header">
        <div>
            <h1>👋 Halo, Selamat Datang {st.session_state.user_nama}!</h1>
            <p>Daftar lengkap informasi dan data analisis prestasi akademik Anda.</p>
        </div>
        <div><h2 style="color:#0D9488; margin:0; font-weight:900;">SMPN 6 Salatiga 🏫</h2></div>
    </div>
    """, unsafe_allow_html=True)
    
    df = pd.DataFrame(st.session_state.database_siswa)
    
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-icon green">👥</div><div class="kpi-info"><h4>Jumlah Siswa</h4><h2>{len(df)}</h2><p>Siswa yang sudah didata</p></div></div>', unsafe_allow_html=True)
    with c2:
        avg = df['Nilai Akademik'].mean() if not df.empty else 0
        st.markdown(f'<div class="kpi-card"><div class="kpi-icon blue">📈</div><div class="kpi-info"><h4>Rata-rata Nilai</h4><h2>{avg:.2f}</h2><p>Nilai rata-rata keseluruhan</p></div></div>', unsafe_allow_html=True)
    with c3:
        max_val = df['Nilai Akademik'].max() if not df.empty else 0
        st.markdown(f'<div class="kpi-card"><div class="kpi-icon orange">🏆</div><div class="kpi-info"><h4>Nilai Tertinggi</h4><h2>{max_val:.2f}</h2><p>Skor tertinggi yang dicapai</p></div></div>', unsafe_allow_html=True)
    with c4:
        min_val = df['Nilai Akademik'].min() if not df.empty else 0
        st.markdown(f'<div class="kpi-card"><div class="kpi-icon purple">📉</div><div class="kpi-info"><h4>Nilai Terendah</h4><h2>{min_val:.2f}</h2><p>Skor terendah yang dicapai</p></div></div>', unsafe_allow_html=True)

    if st.session_state.user_role != "admin" and not df.empty:
        total_siswa = len(df)
        sudah_isi = len(df[df["Self-Efficacy"] > 0])
        persentase = sudah_isi / total_siswa if total_siswa > 0 else 0
        st.markdown("### 📊 Progres Pengisian Kuesioner")
        st.progress(persentase)
        st.write(f"**{sudah_isi}** dari **{total_siswa}** siswa telah mengisi kuesioner.")

    st.markdown("<br><hr>", unsafe_allow_html=True)
    
    st.markdown("### Modul Navigasi")
    mods = [
        ("📝", "Input & Upload", "Input manual atau upload Excel nilai.", "analisis"),
        ("📊", "Data Siswa", "Lihat & kelola daftar siswa Anda.", "monitoring"),
        ("🔍", "Pengaruh Faktor", "Lihat dampak & prioritas per siswa.", "kausal"),
        ("💡", "Laporan", "Saran tindak lanjut & laporan.", "rekomendasi"),
        ("⚙️", "Pengaturan", "Ubah password & profil.", "pengaturan"),
        ("ℹ️", "Tentang", "Info metode Hybrid Causal-Explainable.", "tentang")
    ]
    
    if st.session_state.user_role == "admin":
        mods.append(("👥", "Manajemen Pengguna", "Setujui akun & hapus data.", "admin_users"))
        
    cols = st.columns(3)
    for i, (icon, title, desc, page) in enumerate(mods):
        with cols[i % 3]:
            st.markdown(f'''
            <div class="module-card">
                <div class="module-icon">{icon}</div>
                <div class="module-title">{title}</div>
                <div class="module-desc">{desc}</div>
            </div>
            ''', unsafe_allow_html=True)
            
            if st.button(f"Buka {title}", key=f"home_btn_{page}", use_container_width=True):
                goto_page(page)

# ================================================================
# MODUL 03: PENGARUH FAKTOR (GABUNGAN KAUSAL & SHAP)
# ================================================================
elif st.session_state.current_page == "kausal":
    st.title("🔍 Pengaruh Faktor terhadap Nilai")
    st.write("Halaman ini menjelaskan faktor apa saja yang mempengaruhi nilai siswa, baik dari segi **sebab-akibat** (faktor yang jika diperbaiki bisa menaikkan nilai) maupun dari segi **prediksi** (faktor yang paling sering digunakan sistem untuk menebak nilai).")

    tab_sekolah, tab_siswa = st.tabs(["🏫 Gambaran Umum Sekolah", "👤 Analisis Per Siswa"])

    with tab_sekolah:
        st.markdown("### 1. Faktor yang Jika Diperbaiki Bisa Menaikkan Nilai (Dampak Sebab-Akibat)")
        st.info("Grafik ini menunjukkan besarnya dampak (dalam poin nilai) jika suatu faktor ditingkatkan. Batang hijau (ke kanan) berarti menaikkan nilai, batang merah (ke kiri) berarti menurunkan nilai.")
        
        df_ate = pd.DataFrame({"Faktor": list(PENGARUH_DATA.keys()), "Dampak (Poin)": list(PENGARUH_DATA.values())}).sort_values("Dampak (Poin)")
        
        fig1 = px.bar(df_ate, x="Dampak (Poin)", y="Faktor", orientation='h', 
                      color="Dampak (Poin)", color_continuous_scale="RdBu",
                      labels={"Dampak (Poin)": "Dampak terhadap Nilai"})
        fig1.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig1, use_container_width=True)

        st.markdown("### 2. Faktor yang Paling Sering Dipakai untuk Memprediksi Nilai (Tingkat Kepentingan)")
        st.info("Grafik ini menunjukkan seberapa sering sistem komputer menggunakan faktor tersebut untuk memprediksi nilai siswa. Semakin panjang batangnya, semakin penting faktor tersebut dalam prediksi.")
        
        df_shap = pd.DataFrame({"Faktor": list(KEPENTINGAN_DATA.keys()), "Tingkat Kepentingan": list(KEPENTINGAN_DATA.values())}).sort_values("Tingkat Kepentingan")
        
        fig2 = px.bar(df_shap, x="Tingkat Kepentingan", y="Faktor", orientation='h', 
                      color_discrete_sequence=['#0D9488'],
                      labels={"Tingkat Kepentingan": "Skor Kepentingan"})
        fig2.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig2, use_container_width=True)

        st.markdown("### 💡 Kesimpulan untuk Guru")
        st.success("Berdasarkan analisis di atas, **Self-Efficacy Akademik** dan **Keterlibatan Orang Tua** adalah dua faktor paling kuat yang bisa menaikkan nilai siswa sekaligus paling penting dalam memprediksi nilai. Fokuskan intervensi sekolah pada dua aspek ini.")

    with tab_siswa:
        st.markdown("### Analisis Profil Siswa Spesifik")
        df_all = pd.DataFrame(st.session_state.database_siswa)
        if df_all.empty:
            st.warning("Belum ada data siswa. Silakan input data di menu Input & Upload.")
            st.stop()
            
        df_siswa = df_all[df_all["Self-Efficacy"] > 0]
        if df_siswa.empty:
            st.warning("⚠️ Belum ada siswa yang mengisi kuesioner.")
            st.info("Analisis hanya bisa dilakukan untuk siswa yang sudah mengisi kuesioner mandiri di halaman login.")
            st.stop()

        opsi_siswa = df_siswa.apply(lambda x: f"{x['Nama']} - {x['Kelas']} (Absen {x['Absen']})", axis=1).tolist()
        siswa_terpilih = st.selectbox("Pilih Siswa untuk Dianalisis", opsi_siswa)
        idx = opsi_siswa.index(siswa_terpilih)
        b = df_siswa.iloc[idx]

        nama = b["Nama"]; kelas = b["Kelas"]; absen = int(b["Absen"]); nilai = float(b["Nilai Akademik"])
        profil = {
            "Self-Efficacy Akademik": float(b["Self-Efficacy"]), "Keterlibatan Orang Tua": float(b["Keterlibatan Ortu"]),
            "Harapan Orang Tua": float(b["Harapan Ortu"]), "Dukungan Sekolah": float(b["Dukungan Sekolah"]),
            "Motivasi Belajar": float(b["Motivasi"]), "Kecemasan Akademik": float(b["Kecemasan"]),
            "Fasilitas Sekolah": float(b["Fasilitas"]), "Kemalasan Belajar": float(b["Kemalasan"]),
        }

        selisih, kontribusi = analisis_kausal(nilai, profil)

        st.markdown(f"#### Profil Analisis: **{nama}**")
        st.info(f"Nilai Rata-rata: **{nilai:.2f}** | Posisi vs Rata-rata Sekolah: **{selisih:+.2f} poin**")

        df_k = pd.DataFrame([
            {"Faktor": k, "Nilai Siswa": profil[k], "Rata-rata Sekolah": BASELINE_ASPEK[k],
             "Selisih": profil[k] - BASELINE_ASPEK[k], "Dampak (Poin)": v}
            for k, v in kontribusi.items()
        ]).sort_values("Dampak (Poin)")

        st.markdown("**Radar Chart: Perbandingan Profil Siswa vs Rata-rata Sekolah**")
        df_radar = pd.DataFrame({"Faktor": list(profil.keys()), "Nilai Siswa": list(profil.values()), "Rata-rata Sekolah": [BASELINE_ASPEK[k] for k in profil.keys()]})
        
        fig_radar = go.Figure()
        fig_radar.add_trace(go.Scatterpolar(r=df_radar["Rata-rata Sekolah"], theta=df_radar["Faktor"], fill='toself', name='Rata-rata Sekolah', line_color='#94A3B8'))
        fig_radar.add_trace(go.Scatterpolar(r=df_radar["Nilai Siswa"], theta=df_radar["Faktor"], fill='toself', name=f'Profil {nama}', line_color='#0D9488'))
        fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 5])), showlegend=True, margin=dict(l=40, r=40, t=40, b=40), height=400)
        st.plotly_chart(fig_radar, use_container_width=True)

        st.markdown("**Grafik Dampak Faktor (Poin)**")
        fig_bar = px.bar(df_k, x="Dampak (Poin)", y="Faktor", orientation='h', color="Dampak (Poin)", color_continuous_scale="RdBu")
        fig_bar.update_layout(height=350, margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fig_bar, use_container_width=True)

        st.markdown("**Tabel Detail**")
        st.dataframe(df_k, use_container_width=True, hide_index=True)

# ================================================================
# MODUL PREDIKSI
# ================================================================
elif st.session_state.current_page == "prediksi":
    st.title("📈 Prediksi Nilai Siswa")
    st.write("Halaman ini mensimulasikan prediksi nilai siswa berdasarkan profil 8 faktor. (Catatan: Model Machine Learning asli dapat diintegrasikan di sini).")
    
    with st.container():
        st.markdown("### Masukkan Profil Siswa")
        col1, col2 = st.columns(2)
        with col1:
            p_se = st.slider("Self-Efficacy Akademik", 1.0, 5.0, 4.0, 0.1)
            p_mot = st.slider("Motivasi Belajar", 1.0, 5.0, 3.5, 0.1)
            p_cem = st.slider("Kecemasan Akademik", 1.0, 5.0, 2.5, 0.1)
            p_mal = st.slider("Kemalasan Belajar", 1.0, 5.0, 2.0, 0.1)
        with col2:
            p_ortu = st.slider("Keterlibatan Orang Tua", 1.0, 5.0, 4.0, 0.1)
            p_harap = st.slider("Harapan Orang Tua", 1.0, 5.0, 3.5, 0.1)
            p_sekolah = st.slider("Dukungan Sekolah", 1.0, 5.0, 4.0, 0.1)
            p_fas = st.slider("Fasilitas Sekolah", 1.0, 5.0, 4.5, 0.1)
            
        if st.button("🔮 Hitung Prediksi Nilai", type="primary", use_container_width=True):
            base_score = 75.0
            score = base_score + (p_se * 1.5) + (p_ortu * 1.2) + (p_mot * 1.0) - (p_cem * 0.8) - (p_mal * 1.5) + (p_sekolah * 0.5)
            score = min(max(score, 60.0), 100.0)
            
            kat, warna = kategori_nilai(score)
            
            st.markdown(f"""
            <div style="background:white; border-left:5px solid {warna}; padding:2rem; border-radius:12px; box-shadow:0 4px 15px rgba(0,0,0,0.05); text-align:center; margin-top:20px;">
                <h3 style="margin:0; color:#64748B;">Hasil Prediksi Nilai Akademik</h3>
                <h1 style="font-size:4rem; color:{warna}; margin:10px 0;">{score:.2f}</h1>
                <h4 style="margin:0; color:{warna};">Kategori: {kat}</h4>
            </div>
            """, unsafe_allow_html=True)
            st.info("💡 *Angka di atas adalah simulasi (Mockup) menggunakan perhitungan linear sederhana. Kamu bisa mengganti bagian logika ini dengan model Random Forest (.pkl) milikmu.*")

# ================================================================
# MODUL ADMIN: MANAJEMEN PENGGUNA
# ================================================================
elif st.session_state.current_page == "admin_users":
    st.title("👥 Manajemen Pengguna")
    users = load_users()
    pending_users = {u: d for u, d in users.items() if d.get("status") == "pending"}
    st.markdown(f"### Antrean Persetujuan ({len(pending_users)})")
    
    if not pending_users:
        st.success("Tidak ada akun yang menunggu persetujuan.")
    else:
        for u, d in pending_users.items():
            with st.container():
                col1, col2, col3 = st.columns([3, 1, 1])
                col1.write(f"**{d['nama']}** (@{u}) - Mendaftar: {d['created']}")
                if col2.button("✅ Setujui", key=f"app_{u}", use_container_width=True):
                    users[u]["status"] = "approved"
                    save_users(users)
                    st.success(f"Akun {u} disetujui.")
                    st.rerun()
                if col3.button("❌ Tolak", key=f"rej_{u}", use_container_width=True):
                    users[u]["status"] = "rejected"
                    save_users(users)
                    st.error(f"Akun {u} ditolak.")
                    st.rerun()
            st.divider()

    st.markdown("### Daftar Semua Akun")
    for u, d in users.items():
        with st.expander(f"{d['nama']} (@{u}) - Status: {d['status'].upper()}"):
            c1, c2 = st.columns([1, 3])
            with c1: st.write(f"**Role:** {d['role']}")
            with c2:
                if u not in SUPER_ADMIN:
                    new_pw = st.text_input(f"Password Baru untuk {u}", type="password", key=f"pw_{u}")
                    col_reset, col_hapus = st.columns(2)
                    with col_reset:
                        if st.button(f"Reset Password", key=f"reset_{u}", use_container_width=True):
                            if len(new_pw) >= 6:
                                users[u]["password"] = hash_password(new_pw)
                                save_users(users)
                                st.success(f"Password {u} berhasil diubah.")
                                st.rerun()
                            else: st.error("Password minimal 6 karakter.")
                    with col_hapus:
                        if st.button(f"🗑️ Hapus Akun", key=f"del_{u}", use_container_width=True, type="primary"):
                            del users[u]
                            save_users(users)
                            st.success(f"Akun @{u} berhasil dihapus permanen.")
                            st.rerun()
                else: st.info("Akun Super Admin (Terkunci).")

    st.markdown("---")
    st.markdown("### 📋 Pantau Semua Log Aktivitas (Global)")
    if st.button("Muat Semua Log Sistem"):
        all_logs = load_log(user=None) 
        if all_logs:
            df_log = pd.DataFrame(all_logs)[["waktu", "user", "role", "aksi"]]
            st.dataframe(df_log, use_container_width=True, hide_index=True)
        else: st.info("Belum ada log sama sekali.")

    st.markdown("---")
    st.markdown("### 🗑️ Hapus Data Siswa (Kuesioner Salah)")
    all_siswa = load_all_siswa()
    if not all_siswa:
        st.info("Belum ada data siswa di sistem.")
    else:
        df_all = pd.DataFrame(all_siswa)
        st.dataframe(df_all[["Nama", "Kelas", "Absen", "Nilai Akademik", "Dicatat Oleh"]], use_container_width=True, hide_index=True)
        opsi_hapus = {s["_id"]: f"{s['Nama']} - {s['Kelas']} (Absen {s['Absen']}) - Input oleh: {s.get('Dicatat Oleh', 'Sistem')}" for s in all_siswa}
        siswa_id_terpilih = st.selectbox("Pilih siswa yang akan dihapus:", options=list(opsi_hapus.keys()), format_func=lambda x: opsi_hapus[x])
        
        konfirmasi_hapus = st.checkbox("Saya yakin ingin menghapus data siswa ini.")
        if st.button("Hapus Data Siswa Ini", type="primary", disabled=not konfirmasi_hapus):
            hapus_siswa_by_id(siswa_id_terpilih)
            st.success("Data siswa berhasil dihapus permanen dari database.")
            st.rerun()

        st.markdown("---")
        with st.expander("🚨 ZONA BERBAHAYA: Hapus SEMUA Data Siswa"):
            st.warning("Tindakan ini akan menghapus seluruh data siswa dari SEMUA guru secara permanen. Pastikan Anda sudah melakukan backup!")
            konfirmasi_semua = st.checkbox("Saya mengerti dan ingin menghapus SEMUA data siswa.")
            if st.button("Hapus SEMUA Data Siswa", type="primary", disabled=not konfirmasi_semua):
                jumlah = hapus_semua_siswa_admin()
                st.success(f"Berhasil menghapus {jumlah} data siswa dari database.")
                st.rerun()

# ================================================================
# MODUL 01: INPUT & UPLOAD
# ================================================================
elif st.session_state.current_page == "analisis":
    st.title("📝 Input & Upload Data Nilai")
    tab_manual, tab_excel = st.tabs(["✍️ Input Manual (Satu Siswa)", "📤 Upload Excel (Banyak Siswa)"])
    
    with tab_excel:
        st.write("Upload data nilai siswa sekaligus. (Kuesioner diisi mandiri oleh siswa).")
        wb = Workbook(); ws = wb.active; ws.title = "Nilai"
        ws.append(["Nama Siswa", "Kelas", "Absen", "Nilai Rapor"])
        ws.append(["Contoh 1", "IX-A", 1, 85.5])
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 18
        out = io.BytesIO(); wb.save(out)
        
        st.download_button("📥 Download Template Excel", data=out.getvalue(), file_name="Template_Nilai.xlsx")
        
        uploaded = st.file_uploader("Upload file Excel", type=["xlsx"])
        if uploaded:
            df = pd.read_excel(uploaded)
            st.dataframe(df.head())
            if st.button("Simpan Semua Nilai ke Database", type="primary"):
                for _, row in df.iterrows():
                    n = float(row["Nilai Rapor"])
                    kat, _ = kategori_nilai(n)
                    insert_siswa({
                        "Timestamp": get_wib_time().strftime("%Y-%m-%d %H:%M"),
                        "Nama": str(row["Nama Siswa"]).strip(), "Kelas": str(row["Kelas"]).strip(),
                        "Absen": int(row["Absen"]), "Nilai Akademik": n, "Kategori": kat,
                        "Self-Efficacy": 0.0, "Keterlibatan Ortu": 0.0, "Harapan Ortu": 0.0,
                        "Dukungan Sekolah": 0.0, "Motivasi": 0.0, "Kecemasan": 0.0,
                        "Fasilitas": 0.0, "Kemalasan": 0.0,
                        "Dicatat Oleh": st.session_state.user_nama
                    })
                add_log("Upload Excel", f"{len(df)} siswa")
                reload_data()
                st.success("Data nilai berhasil disimpan! Kuesioner akan menyusul jika siswa mengisi.")
                st.rerun()
                
    with tab_manual:
        st.write("Input satu siswa beserta kuesionernya sekaligus.")
        with st.form("manual_form"):
            c1, c2, c3 = st.columns([2, 1, 1])
            m_nama = c1.text_input("Nama Siswa")
            m_kelas = c2.selectbox("Kelas", KELAS_LIST)
            m_absen = c3.number_input("Absen", 1, 50, 1)
            m_nilai = st.number_input("Nilai Rapor", 60.0, 100.0, 83.78)
            
            st.markdown("---")
            f1 = input_kategori("m_s1", "Self-Efficacy Akademik", False)
            f2 = input_kategori("m_s2", "Motivasi Belajar", False)
            f3 = input_kategori("m_s3", "Kecemasan Akademik", False)
            f4 = input_kategori("m_s4", "Kemalasan Belajar", False)
            f5 = input_kategori("m_s5", "Keterlibatan Orang Tua", False)
            f6 = input_kategori("m_s6", "Harapan Orang Tua", False)
            f7 = input_kategori("m_s7", "Dukungan Sekolah", False)
            f8 = input_kategori("m_s8", "Fasilitas Sekolah", False)
            
            if st.form_submit_button("Simpan Siswa Ini"):
                kat, _ = kategori_nilai(m_nilai)
                insert_siswa({
                    "Timestamp": get_wib_time().strftime("%Y-%m-%d %H:%M"),
                    "Nama": m_nama, "Kelas": m_kelas, "Absen": m_absen,
                    "Nilai Akademik": m_nilai, "Kategori": kat,
                    "Self-Efficacy": f1, "Keterlibatan Ortu": f5, "Harapan Ortu": f6,
                    "Dukungan Sekolah": f7, "Motivasi": f2, "Kecemasan": f3,
                    "Fasilitas": f8, "Kemalasan": f4, "Dicatat Oleh": st.session_state.user_nama
                })
                add_log("Input Manual", m_nama)
                reload_data()
                st.success("Data berhasil disimpan.")
                st.rerun()

# ================================================================
# MODUL 02: DATA SISWA
# ================================================================
elif st.session_state.current_page == "monitoring":
    st.title("📊 Data Siswa Anda")
    df = pd.DataFrame(st.session_state.database_siswa)
    if df.empty:
        st.warning("Belum ada data siswa. Silakan input atau upload di menu Input & Upload.")
    else:
        col1, col2 = st.columns(2)
        with col1: cari = st.text_input("🔍 Cari Nama Siswa", placeholder="Ketik nama...")
        with col2: filter_kelas = st.selectbox("Filter Kelas", ["Semua"] + sorted(df["Kelas"].unique().tolist()))
            
        if cari: df = df[df["Nama"].str.contains(cari, case=False)]
        if filter_kelas != "Semua": df = df[df["Kelas"] == filter_kelas]
            
        st.dataframe(df[["Nama", "Kelas", "Absen", "Nilai Akademik", "Kategori", "Dicatat Oleh"]], use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("### 📥 Download Data (Excel)")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Data Siswa')
        
        st.download_button("⬇️ Download Data Siswa (Excel)", data=buffer.getvalue(), file_name=f"Data_Siswa_{st.session_state.user_nama}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        with st.expander("🗑️ Zona Berbahaya"):
            if st.button("Hapus SEMUA Siswa Saya", type="primary"):
                hapus_semua_siswa(st.session_state.user_nama)
                reload_data()
                st.success("Data Anda berhasil dihapus.")
                st.rerun()

# ================================================================
# MODUL 04: REKOMENDASI & LAPORAN
# ================================================================
elif st.session_state.current_page == "rekomendasi":
    st.title("💡 Laporan & Rekomendasi")
    df_all = pd.DataFrame(st.session_state.database_siswa)
    if df_all.empty:
        st.warning("Belum ada data siswa. Silakan input data di menu Input & Upload.")
        st.stop()
        
    df_siswa = df_all[df_all["Self-Efficacy"] > 0]
    if df_siswa.empty:
        st.warning("⚠️ Belum ada siswa yang mengisi kuesioner.")
        st.stop()

    opsi_siswa = df_siswa.apply(lambda x: f"{x['Nama']} - {x['Kelas']} (Absen {x['Absen']})", axis=1).tolist()
    siswa_terpilih = st.selectbox("Pilih Siswa untuk Rekomendasi", opsi_siswa)
    idx = opsi_siswa.index(siswa_terpilih)
    b = df_siswa.iloc[idx]

    nama = b["Nama"]; kelas = b["Kelas"]; absen = int(b["Absen"]); nilai = float(b["Nilai Akademik"])
    profil = {
        "Self-Efficacy Akademik": float(b["Self-Efficacy"]), "Keterlibatan Orang Tua": float(b["Keterlibatan Ortu"]),
        "Harapan Orang Tua": float(b["Harapan Ortu"]), "Dukungan Sekolah": float(b["Dukungan Sekolah"]),
        "Motivasi Belajar": float(b["Motivasi"]), "Kecemasan Akademik": float(b["Kecemasan"]),
        "Fasilitas Sekolah": float(b["Fasilitas"]), "Kemalasan Belajar": float(b["Kemalasan"]),
    }
    selisih, kontribusi = analisis_kausal(nilai, profil)
    kategori, warna_kategori = kategori_nilai(nilai)

    df_k = pd.DataFrame([
        {"Aspek": k, "Kontribusi": v, "Nilai_Siswa": profil[k], "Baseline": BASELINE_ASPEK[k], "Selisih": profil[k] - BASELINE_ASPEK[k]}
        for k, v in kontribusi.items()
    ]).sort_values("Kontribusi")
    neg = df_k[df_k["Kontribusi"] < -0.1].sort_values("Kontribusi")
    pos = df_k[df_k["Kontribusi"] > 0.1].sort_values("Kontribusi", ascending=False)

    posisi = "di atas" if selisih >= 0 else "di bawah"
    fn = neg.iloc[0] if len(neg) > 0 else None
    fp = pos.iloc[0] if len(pos) > 0 else None

    narasi_neg = f"Faktor yang paling menekan nilai <b>{nama}</b> adalah <b style='color:#EF5B67;'>{fn['Aspek']}</b> (kontribusi {fn['Kontribusi']:.2f} poin)." if fn is not None else "Tidak ada faktor yang signifikan menekan nilai."
    narasi_pos = f"Kekuatan utama terletak pada <b style='color:#10B981;'>{fp['Aspek']}</b> (kontribusi +{fp['Kontribusi']:.2f} poin)." if fp is not None else "Belum ada faktor kekuatan dominan."

    st.markdown(f"""
    <div style="background:white; border-left:5px solid {warna_kategori}; padding:1.5rem; border-radius:12px; box-shadow:0 4px 15px rgba(0,0,0,0.03); margin-bottom:1rem;">
        <div style="display:flex; align-items:center; gap:1rem; margin-bottom:1rem;">
            <div style="width:56px;height:56px;border-radius:15px;background:{warna_kategori};color:white;display:flex;align-items:center;justify-content:center;font-size:1.2rem;font-weight:800;">{"".join(x[0] for x in nama.split()[:2]).upper()}</div>
            <div><div style="font-weight:800;font-size:1.3rem; color:#1E293B;">{nama}</div><div style="color:#64748B;font-size:.8rem;">{kelas} - Absen {absen:02d} - Nilai {nilai:.2f}</div></div>
        </div>
        <div style="font-size:.95rem;line-height:1.75;color:#1E293B;">
            Nilai saat ini <b>{abs(selisih):.2f} poin {posisi} rata-rata sekolah</b> ({RATA_RATA_NILAI:.2f}).<br><br>{narasi_neg}<br><br>{narasi_pos}<br><br>
            <b>Kategori:</b> <span style="color:{warna_kategori};font-weight:800;">{kategori}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Prioritas Perbaikan (3 Hal Utama)")
    prioritas = neg.head(3)
    if len(prioritas) == 0:
        st.success("Tidak ada prioritas perbaikan mendesak.")
    else:
        cols = st.columns(len(prioritas))
        for i, (_, row) in enumerate(prioritas.iterrows()):
            with cols[i]:
                gap_val = row["Selisih"]
                level = "Segera" if gap_val < -1 else "Perhatian" if gap_val < -0.5 else "Pantau"
                st.markdown(f"""
                <div style="background:white; border-top:4px solid #EF5B67; padding:1.5rem; border-radius:12px; box-shadow:0 4px 15px rgba(0,0,0,0.03); min-height:200px;">
                    <div style="font-size:.65rem;font-weight:800;letter-spacing:1.5px;color:#EF5B67;">PRIORITAS {i + 1}</div>
                    <div style="font-weight:800;font-size:1rem;margin:.6rem 0 .3rem; color:#1E293B;">{row["Aspek"]}</div>
                    <div style="font-size:.7rem;color:#64748B;margin-bottom:.5rem;">{level}</div>
                    <div style="font-size:.75rem;line-height:1.6;color:#475467;">
                        Nilai siswa: <b>{row["Nilai_Siswa"]:.2f}</b><br>Rata-rata: <b>{row["Baseline"]:.2f}</b><br>
                        <span style="color:#EF5B67;font-weight:700;">Selisih: {row["Selisih"]:+.2f}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)

    st.markdown("### Tindak Lanjut (Checklist Guru & Siswa)")
    TINDAK = {
        "Self-Efficacy Akademik": {"guru": ["Berikan tugas bertahap", "Pujian spesifik atas usaha", "Ajak refleksi mingguan", "Pasangkan dengan peer-mentor"], "siswa": ["Jurnal harian 1 hal yang berhasil", "Tetapkan target kecil mingguan"]},
        "Keterlibatan Orang Tua": {"guru": ["Kirim kabar positif ke orang tua", "Ajak orang tua ikut sesi belajar", "Panduan mendampingi belajar 15 menit/hari", "Komunikasi rutin 2 minggu sekali"], "siswa": ["Ceritakan 1 hal yang dipelajari", "Minta orang tua periksa PR"]},
        "Harapan Orang Tua": {"guru": ["Pertemuan ekspektasi realistis", "Bantu pahami tahap perkembangan anak", "Sarankan fokus pada usaha", "Contoh memotivasi tanpa menekan"], "siswa": ["Belajar menyampaikan perasaan", "Fokus pada usaha yang bisa dikontrol"]},
        "Dukungan Sekolah": {"guru": ["Refleksi bantuan berlebih", "Kurangi bantuan yang bisa dilakukan sendiri", "Berikan kesempatan mencoba", "Fokus membimbing, bukan menggantikan"], "siswa": ["Coba selesaikan tugas 10 menit sebelum bertanya", "Catat apa yang sudah dicoba"]},
        "Motivasi Belajar": {"guru": ["Kaitkan materi dengan kehidupan nyata", "Berikan pilihan tugas", "Apresiasi proses", "Ciptakan suasana kelas menyenangkan"], "siswa": ["Cari 1 hal menarik dari tiap pelajaran", "Belajar bersama teman"]},
        "Kecemasan Akademik": {"guru": ["Ajarkan teknik relaksasi", "Ubah suasana ujian lebih santai", "Ujian formatif yang tidak menakutkan", "Normalisasi cemas itu wajar"], "siswa": ["Latihan pernapasan 4-7-8", "Persiapan lebih awal"]},
        "Fasilitas Sekolah": {"guru": ["Informasikan fasilitas yang tersedia", "Bantu akses perpustakaan atau lab", "Cek hambatan akses"], "siswa": ["Manfaatkan perpustakaan", "Tanyakan ke guru jika butuh bantuan"]},
        "Kemalasan Belajar": {"guru": ["Cari akar kemalasan", "Beri tugas lebih menantang jika bosan", "Pecah tugas besar jadi langkah kecil", "Buat sistem reward sederhana"], "siswa": ["Mulai dari tugas 5 menit", "Gunakan teknik Pomodoro"]},
    }

    if len(prioritas) == 0:
        st.info("Tidak ada tindak lanjut khusus.")
    else:
        for i, (_, row) in enumerate(prioritas.iterrows()):
            aspek = row["Aspek"]
            tugas = TINDAK.get(aspek, {"guru": [], "siswa": []})
            with st.expander(f"Prioritas {i + 1}: {aspek}", expanded=(i == 0)):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**Yang bisa dilakukan GURU:**")
                    for j, t in enumerate(tugas["guru"], 1): st.markdown(f"{j}. {t}")
                with c2:
                    st.markdown("**Yang bisa dilakukan SISWA:**")
                    for j, s in enumerate(tugas["siswa"], 1): st.markdown(f"{j}. {s}")

    st.markdown("---")
    st.markdown("### 📄 Unduh Laporan Konsultasi")
    st.write("Laporan ini dirancang dengan bahasa yang mudah dipahami dan siap di-print.")

    def buat_dokumen_word(siswa_data):
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        doc = Document()
        header = doc.add_heading('LAPORAN PERKEMBANGAN SISWA', 0); header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subheader = doc.add_paragraph('Smart Academic Analytics (SAA) - SMPN 6 Salatiga'); subheader.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
        table = doc.add_table(rows=4, cols=2); table.style = 'Table Grid'
        table.rows[0].cells[0].text = 'Nama Siswa'; table.rows[0].cells[1].text = siswa_data['nama']
        table.rows[1].cells[0].text = 'Kelas / Absen'; table.rows[1].cells[1].text = f"{siswa_data['kelas']} / {siswa_data['absen']}"
        table.rows[2].cells[0].text = 'Nilai Rata-rata'; table.rows[2].cells[1].text = f"{siswa_data['nilai']:.2f}"
        table.rows[3].cells[0].text = 'Kategori'; table.rows[3].cells[1].text = siswa_data['kategori']
        doc.add_paragraph()
        doc.add_heading('Ringkasan Kondisi', level=1)
        posisi_teks = "di atas" if siswa_data['selisih'] >= 0 else "di bawah"
        doc.add_paragraph(f"Saat ini, {siswa_data['nama']} memiliki nilai rata-rata {siswa_data['nilai']:.2f}, yang berada {abs(siswa_data['selisih']):.2f} poin {posisi_teks} rata-rata sekolah. Berdasarkan analisis, {siswa_data['nama']} masuk dalam kategori {siswa_data['kategori']}.")
        doc.add_heading('Catatan untuk Orang Tua & Guru', level=1)
        if siswa_data['fp'] is not None:
            doc.add_paragraph(f"KEKUATAN UTAMA: {siswa_data['nama']} menunjukkan sikap yang sangat baik pada aspek {siswa_data['fp']['Aspek']}.")
        if siswa_data['fn'] is not None:
            doc.add_paragraph(f"AREA YANG PERLU DIDAMPINGI: Untuk meningkatkan hasil belajar, kita bisa fokus membantu {siswa_data['nama']} pada aspek {siswa_data['fn']['Aspek']}. Berikut saran tindak lanjutnya:")
            saran_guru = siswa_data['tindak'].get(siswa_data['fn']['Aspek'], {}).get("guru", ["Berikan pendampingan ekstra."])
            saran_siswa = siswa_data['tindak'].get(siswa_data['fn']['Aspek'], {}).get("siswa", ["Belajar lebih giat lagi."])
            doc.add_paragraph("Saran untuk Guru/Sekolah:", style='List Bullet')
            for s in saran_guru[:3]: doc.add_paragraph(s, style='List Bullet 2')
            doc.add_paragraph("Saran untuk Siswa/Orang Tua di Rumah:", style='List Bullet')
            for s in saran_siswa[:2]: doc.add_paragraph(s, style='List Bullet 2')
        doc.add_paragraph()
        ttd_table = doc.add_table(rows=2, cols=2); ttd_table.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        ttd_table.rows[0].cells[0].text = 'Guru / Konselor,'; ttd_table.rows[0].cells[1].text = f"Salatiga, {get_wib_time().strftime('%d %B %Y')}"
        ttd_table.rows[1].cells[0].text = '\n\n\n\n(_______________________)'; ttd_table.rows[1].cells[1].text = f'\n\n\n\n( {st.session_state.user_nama} )'
        buffer = io.BytesIO(); doc.save(buffer); buffer.seek(0)
        return buffer

    def buat_dokumen_pdf(siswa_data):
        from fpdf import FPDF
        pdf = FPDF(); pdf.add_page()
        pdf.set_font("Arial", 'B', 16); pdf.cell(0, 10, 'LAPORAN PERKEMBANGAN SISWA', ln=True, align='C')
        pdf.set_font("Arial", 'I', 10); pdf.cell(0, 10, 'Smart Academic Analytics (SAA) - SMPN 6 Salatiga', ln=True, align='C'); pdf.ln(10)
        
        # Tabel Info
        pdf.set_font("Arial", 'B', 11); pdf.cell(40, 8, 'Nama Siswa', border=1); pdf.set_font("Arial", '', 11); pdf.cell(0, 8, siswa_data['nama'], border=1, ln=True)
        pdf.set_font("Arial", 'B', 11); pdf.cell(40, 8, 'Kelas / Absen', border=1); pdf.set_font("Arial", '', 11); pdf.cell(0, 8, f"{siswa_data['kelas']} / {siswa_data['absen']}", border=1, ln=True)
        pdf.set_font("Arial", 'B', 11); pdf.cell(40, 8, 'Nilai Rata-rata', border=1); pdf.set_font("Arial", '', 11); pdf.cell(0, 8, f"{siswa_data['nilai']:.2f}", border=1, ln=True)
        pdf.set_font("Arial", 'B', 11); pdf.cell(40, 8, 'Kategori', border=1); pdf.set_font("Arial", '', 11); pdf.cell(0, 8, siswa_data['kategori'], border=1, ln=True); pdf.ln(10)
        
        # Ringkasan
        pdf.set_font("Arial", 'B', 12); pdf.cell(0, 8, 'Ringkasan Kondisi', ln=True)
        pdf.set_font("Arial", '', 11); posisi_teks = "di atas" if siswa_data['selisih'] >= 0 else "di bawah"
        
        teks_ringkasan = f"Saat ini, {siswa_data['nama']} memiliki nilai rata-rata {siswa_data['nilai']:.2f}, yang berada {abs(siswa_data['selisih']):.2f} poin {posisi_teks} rata-rata sekolah."
        pdf.multi_cell(190, 6, teks_ringkasan); pdf.ln(5)
        
        # Catatan
        pdf.set_font("Arial", 'B', 12); pdf.cell(0, 8, 'Catatan untuk Orang Tua & Guru', ln=True); pdf.set_font("Arial", '', 11)
        
        if siswa_data['fp'] is not None:
            teks_kekuatan = f"KEKUATAN UTAMA: {siswa_data['nama']} menunjukkan sikap yang sangat baik pada aspek {siswa_data['fp']['Aspek']}."
            pdf.multi_cell(190, 6, teks_kekuatan); pdf.ln(2)
            
        if siswa_data['fn'] is not None:
            teks_area = f"AREA YANG PERLU DIDAMPINGI: Fokus pada aspek {siswa_data['fn']['Aspek']}."
            pdf.multi_cell(190, 6, teks_area); pdf.ln(2)
            
            saran_guru = siswa_data['tindak'].get(siswa_data['fn']['Aspek'], {}).get("guru", ["Berikan pendampingan ekstra."])
            saran_siswa = siswa_data['tindak'].get(siswa_data['fn']['Aspek'], {}).get("siswa", ["Belajar lebih giat lagi."])
            
            pdf.set_font("Arial", 'B', 11); pdf.cell(0, 6, 'Saran untuk Guru/Sekolah:', ln=True); pdf.set_font("Arial", '', 11)
            for s in saran_guru[:3]: 
                pdf.multi_cell(190, 6, f"- {s}")
                
            pdf.set_font("Arial", 'B', 11); pdf.cell(0, 6, 'Saran untuk Siswa/Orang Tua di Rumah:', ln=True); pdf.set_font("Arial", '', 11)
            for s in saran_siswa[:2]: 
                pdf.multi_cell(190, 6, f"- {s}")
            
        # Tanda Tangan
        pdf.ln(15); pdf.cell(100, 8, 'Guru / Konselor,', ln=False); pdf.cell(0, 8, f"Salatiga, {get_wib_time().strftime('%d %B %Y')}", ln=True)
        pdf.ln(15); pdf.cell(100, 8, '(_______________________)', ln=False); pdf.cell(0, 8, f"( {st.session_state.user_nama} )", ln=True)
        
        # Output PDF
        pdf_output = pdf.output(dest='S')
        return io.BytesIO(pdf_output)

    data_laporan = {'nama': nama, 'kelas': kelas, 'absen': absen, 'nilai': nilai, 'kategori': kategori, 'selisih': selisih, 'fp': fp, 'fn': fn, 'tindak': TINDAK}

    col_word, col_pdf = st.columns(2)
    with col_word:
        if st.button("📥 Download Laporan (Word)", type="primary", use_container_width=True):
            try:
                file_word = buat_dokumen_word(data_laporan)
                st.download_button(label="⬇️ Simpan Word", data=file_word, file_name=f"Laporan_{nama.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as e: st.error(f"Gagal membuat Word: {e}")
    with col_pdf:
        if st.button("📥 Download Laporan (PDF)", type="primary", use_container_width=True):
            try:
                file_pdf = buat_dokumen_pdf(data_laporan)
                st.download_button(label="⬇️ Simpan PDF", data=file_pdf, file_name=f"Laporan_{nama.replace(' ', '_')}.pdf", mime="application/pdf")
            except Exception as e: st.error(f"Gagal membuat PDF: {e}")

    st.markdown("---")
    st.markdown("### 🗂️ Unduh Semua Laporan (Massal)")
    if st.button("📦 Buat File ZIP (Semua Siswa)", use_container_width=True):
        with st.spinner("Sedang menyiapkan file ZIP..."):
            try:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for _, row in df_siswa.iterrows():
                        s_nama = row["Nama"]
                        s_profil = {"Self-Efficacy Akademik": float(row["Self-Efficacy"]), "Keterlibatan Orang Tua": float(row["Keterlibatan Ortu"]), "Harapan Orang Tua": float(row["Harapan Ortu"]), "Dukungan Sekolah": float(row["Dukungan Sekolah"]), "Motivasi Belajar": float(row["Motivasi"]), "Kecemasan Akademik": float(row["Kecemasan"]), "Fasilitas Sekolah": float(row["Fasilitas"]), "Kemalasan Belajar": float(row["Kemalasan"])}
                        s_selisih, s_kontribusi = analisis_kausal(float(row["Nilai Akademik"]), s_profil)
                        s_kat, _ = kategori_nilai(float(row["Nilai Akademik"]))
                        s_df_k = pd.DataFrame([{"Aspek": k, "Kontribusi": v} for k, v in s_kontribusi.items()]).sort_values("Kontribusi")
                        s_neg = s_df_k[s_df_k["Kontribusi"] < -0.1]; s_pos = s_df_k[s_df_k["Kontribusi"] > 0.1].sort_values("Kontribusi", ascending=False)
                        s_data = {'nama': s_nama, 'kelas': row["Kelas"], 'absen': int(row["Absen"]), 'nilai': float(row["Nilai Akademik"]), 'kategori': s_kat, 'selisih': s_selisih, 'fp': s_pos.iloc[0] if len(s_pos) > 0 else None, 'fn': s_neg.iloc[0] if len(s_neg) > 0 else None, 'tindak': TINDAK}
                        file_buffer = buat_dokumen_word(s_data)
                        zip_file.writestr(f"Laporan_{s_nama.replace(' ', '_')}.docx", file_buffer.getvalue())
                zip_buffer.seek(0)
                st.success(f"Berhasil membuat ZIP berisi {len(df_siswa)} laporan Word!")
                st.download_button(label="⬇️ Download Semua Laporan (.zip)", data=zip_buffer, file_name="Semua_Laporan_Siswa.zip", mime="application/zip")
            except Exception as e: st.error(f"Gagal membuat ZIP: {e}")

# ================================================================
# MODUL 05: PENGATURAN
# ================================================================
elif st.session_state.current_page == "pengaturan":
    st.title("⚙️ Pengaturan & Profil")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Profil Anda")
        st.info(f"**Nama:** {st.session_state.user_nama}\n\n**Role:** {st.session_state.user_role.capitalize()}")
        with st.expander("🔒 Ubah Password"):
            old_pw = st.text_input("Password Lama", type="password")
            new_pw = st.text_input("Password Baru", type="password")
            if st.button("Simpan Password Baru"):
                users = load_users()
                curr_user = next((u for u, d in users.items() if d["nama"] == st.session_state.user_nama), None)
                if curr_user and verify_password(old_pw, users[curr_user]["password"]):
                    if len(new_pw) >= 6:
                        users[curr_user]["password"] = hash_password(new_pw)
                        save_users(users)
                        st.success("Password berhasil diubah!")
                    else: st.error("Password baru minimal 6 karakter.")
                else: st.error("Password lama salah.")
    with c2:
        st.markdown("### Log Aktivitas Terakhir (Milik Anda)")
        if st.session_state.log:
            df_log = pd.DataFrame(st.session_state.log[:5])[["waktu", "aksi"]]
            st.dataframe(df_log, use_container_width=True, hide_index=True)
        else: st.write("Belum ada aktivitas.")

# ================================================================
# MODUL 06: TENTANG SISTEM
# ================================================================
elif st.session_state.current_page == "tentang":
    st.title("ℹ️ Tentang Sistem")
    st.info("""
    **Smart Academic Analytics (SAA)** adalah sistem pendukung keputusan berbasis data yang dirancang khusus untuk SMP Negeri 6 Salatiga.
    Sistem ini mengimplementasikan pendekatan **Hybrid Causal-Explainable Machine Learning** untuk menganalisis determinan (faktor penentu) prestasi akademik siswa.
    ### Komponen Utama:
    1. **Pembentukan Konstruk (NLP)**: Mengelompokkan item kuesioner menjadi 8 konstruk utama menggunakan pendekatan semantik (Sentence-BERT).
    2. **Causal Inference (DoWhy)**: Mengestimasi besarnya pengaruh sebab-akibat (Average Treatment Effect/ATE) dari setiap konstruk terhadap nilai siswa.
    3. **Machine Learning (Random Forest)**: Memprediksi nilai siswa dan mengekstraksi tingkat kepentingan fitur (Feature Importance).
    4. **Explainable AI (SHAP)**: Menjelaskan seberapa besar kontribusi setiap faktor terhadap prediksi model.
    5. **Composite Effect Index (CEI)**: Menggabungkan hasil analisis kausal (ATE) dan prediktif (SHAP) untuk menghasilkan rekomendasi prioritas intervensi yang komprehensif bagi guru dan sekolah.
    *Sistem ini dikembangkan sebagai luaran terapan dari Tesis Program Studi Magister Sains Data, Universitas Kristen Satya Wacana.*
    """)

# ================================================================
# END OF FILE
# ================================================================