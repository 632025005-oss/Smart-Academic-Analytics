# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import hashlib
import json
import io
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from db_storage import (
    load_users as _db_load_users, save_users as _db_save_users,
    load_database, insert_siswa, hapus_semua_siswa,
    load_config, save_config, load_log, save_log, tambah_log,
    update_kuesioner_siswa, load_all_siswa, hapus_siswa_by_id,
    hapus_semua_siswa_admin
)

SUPER_ADMIN = {"admin", "regina"}

# ================================================================
# HELPERS
# ================================================================
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
                "status":   "approved", "created":  datetime.now().strftime("%Y-%m-%d %H:%M"),
            }
        _db_save_users(users)
    return users

def save_users(u): _db_save_users(u)

# ================================================================
# PAGE CONFIG & CSS (DESAIN KACA)
# ================================================================
st.set_page_config(page_title="SAA Analytics", page_icon="🎓", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap');
:root {
    --primary: #6366f1; --primary-hover: #4f46e5;
    --bg: #f8fafc; --surface: #ffffff; --text: #0f172a; --text-muted: #64748b;
}
html, body, [class*="css"] { font-family: 'Outfit', sans-serif; color: var(--text); }
.stApp {
    background: radial-gradient(circle at 0% 0%, #e0e7ff 0%, transparent 50%),
                radial-gradient(circle at 100% 100%, #fce7f3 0%, transparent 50%), var(--bg);
}
.block-container { max-width: 1200px; padding-top: 2rem; }
#MainMenu, footer, header { visibility: hidden; }

/* Glassmorphism Cards */
.glass-card {
    background: rgba(255, 255, 255, 0.7);
    backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px);
    border: 1px solid rgba(255, 255, 255, 0.5);
    border-radius: 20px; padding: 1.5rem;
    box-shadow: 0 8px 32px rgba(31, 38, 135, 0.05);
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}
.glass-card:hover { transform: translateY(-5px); box-shadow: 0 12px 40px rgba(31, 38, 135, 0.1); }

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, var(--primary), var(--primary-hover));
    color: white; border: none; border-radius: 12px; padding: 0.6rem 1.2rem;
    font-weight: 600; transition: all 0.3s ease; box-shadow: 0 4px 15px rgba(99, 102, 241, 0.3);
}
.stButton > button:hover { transform: scale(1.02); box-shadow: 0 6px 20px rgba(99, 102, 241, 0.5); color: white;}

/* Tabs */
.stTabs [data-baseweb="tab-list"] { gap: 1rem; }
.stTabs [data-baseweb="tab"] { border-radius: 10px; padding: 0.5rem 1rem; font-weight: 600; }

/* Metric */
.metric-box { text-align: center; }
.metric-val { font-size: 2.5rem; font-weight: 800; color: var(--primary); }
.metric-lbl { font-size: 0.9rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 1px; }
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
    tambah_log({
        "waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "user": st.session_state.user_nama or "-", 
        "role": st.session_state.user_role or "-",
        "aksi": aksi, 
        "detail": detail
    })

def goto_page(page):
    st.session_state.current_page = page
    st.rerun()

# ================================================================
# KONSTANTA & FUNGSI ANALITIK
# ================================================================
BASELINE_ASPEK = {"Self-Efficacy Akademik": 4.634, "Keterlibatan Orang Tua": 4.348, "Harapan Orang Tua": 3.450, "Dukungan Sekolah": 4.600, "Motivasi Belajar": 2.517, "Kecemasan Akademik": 2.893, "Kemalasan Belajar": 1.487, "Fasilitas Sekolah": 4.880}
PENGARUH_DATA = {"Self-Efficacy Akademik": 2.1995, "Keterlibatan Orang Tua": 1.7559, "Harapan Orang Tua": -0.6987, "Dukungan Sekolah": -1.0351, "Motivasi Belajar": -1.9093, "Kecemasan Akademik": 0.7197, "Kemalasan Belajar": 0.7320, "Fasilitas Sekolah": 0.6182}
KEPENTINGAN_DATA = {"Self-Efficacy Akademik": 0.8197, "Keterlibatan Orang Tua": 0.6717, "Harapan Orang Tua": 0.5766, "Dukungan Sekolah": 0.4325, "Motivasi Belajar": 0.3601, "Kecemasan Akademik": 0.1553, "Fasilitas Sekolah": 0.0654, "Kemalasan Belajar": 0.0846}
BOBOT_PENGARUH = {k: PENGARUH_DATA[k] * KEPENTINGAN_DATA[k] for k in PENGARUH_DATA}
RATA_RATA_NILAI = 83.78

KELAS_LIST = (
    [f"VII-{x}" for x in "ABCDEFGH"] +
    [f"VIII-{x}" for x in "ABCDEFGH"] +
    [f"IX-{x}" for x in "ABCDEFGH"]
)

SKALA_PILIHAN = {
    "Self-Efficacy Akademik": {"q": "Seberapa yakin kamu terhadap kemampuan akademikmu?", "o": [("1","Hampir tidak pernah"),("2","Jarang"),("3","Kadang-kadang"),("4","Sering"),("5","Hampir selalu")]},
    "Keterlibatan Orang Tua": {"q": "Seberapa aktif orang tua mendampingi kamu belajar?", "o": [("1","Hampir tidak pernah"),("2","Jarang"),("3","Kadang-kadang"),("4","Sering"),("5","Hampir selalu")]},
    "Harapan Orang Tua": {"q": "Seberapa tinggi tuntutan orang tua terhadap nilaimu?", "o": [("1","Sangat rendah"),("2","Rendah"),("3","Sedang"),("4","Tinggi"),("5","Sangat tinggi")]},
    "Dukungan Sekolah": {"q": "Seberapa besar dukungan sekolah kepadamu?", "o": [("1","Sangat kurang"),("2","Kurang"),("3","Cukup"),("4","Baik"),("5","Sangat baik")]},
    "Motivasi Belajar": {"q": "Seberapa besar motivasi belajarmu?", "o": [("1","Tidak ada"),("2","Kurang"),("3","Cukup"),("4","Baik"),("5","Sangat baik")]},
    "Kecemasan Akademik": {"q": "Seberapa sering kamu cemas saat belajar/ujian?", "o": [("1","Hampir tidak pernah"),("2","Jarang"),("3","Kadang-kadang"),("4","Sering"),("5","Hampir selalu")]},
    "Fasilitas Sekolah": {"q": "Seberapa memadai fasilitas belajarmu?", "o": [("1","Sangat kurang"),("2","Kurang"),("3","Cukup"),("4","Memadai"),("5","Sangat memadai")]},
    "Kemalasan Belajar": {"q": "Seberapa sering kamu malas belajar?", "o": [("1","Tidak pernah"),("2","Jarang"),("3","Kadang-kadang"),("4","Sering"),("5","Hampir selalu")]}
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
    subjek = "kamu" if is_student else "siswa"
    st.markdown(f"**{fname}** — {cfg['q'].replace('kamu', subjek)}")
    labels = [o[1] for o in cfg["o"]]
    vals = [float(o[0]) for o in cfg["o"]]
    default_idx = min(range(len(vals)), key=lambda i: abs(vals[i] - BASELINE_ASPEK[fname]))
    p = st.radio(fname, labels, index=default_idx, key=key, label_visibility="collapsed")
    return vals[labels.index(p)]

# ================================================================
# HALAMAN LOGIN & KUESIONER SISWA
# ================================================================
def halaman_login():
    st.markdown("""
        <div style="text-align: center; margin-top: 2rem; margin-bottom: 2rem;">
            <h1 style="color: #312e81; font-weight: 800; font-size: 3rem;">SAA.<span style="color: #6366f1;">Analytics</span></h1>
            <p style="color: #64748b; font-size: 1.1rem;">Smart Academic Analytics - SMPN 6 Salatiga</p>
        </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        tab_login, tab_daftar, tab_siswa = st.tabs(["🔐 Masuk Guru", "📝 Daftar Akun", "🎒 Kuesioner Siswa"])

        with tab_login:
            with st.form("login_form"):
                username = st.text_input("Username", placeholder="Masukkan username")
                password = st.text_input("Password", type="password", placeholder="Masukkan password")
                submit = st.form_submit_button("Masuk", use_container_width=True)
                
                if submit:
                    users = load_users()
                    if username in users and verify_password(password, users[username]["password"]):
                        if users[username].get("status") == "pending":
                            st.warning("⏳ Akun Anda belum disetujui. Silakan hubungi Admin.")
                        elif users[username].get("status") == "rejected":
                            st.error("❌ Akun Anda ditolak oleh Admin.")
                        else:
                            st.session_state.logged_in = True
                            st.session_state.user_nama = users[username]["nama"]
                            st.session_state.user_role = users[username]["role"]
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
                reg_role = st.selectbox("Peran", ["guru", "wali_kelas"])
                reg_submit = st.form_submit_button("Daftar", use_container_width=True)
                
                if reg_submit:
                    users = load_users()
                    if reg_user in users:
                        st.error("Username sudah dipakai.")
                    elif len(reg_pw) < 6:
                        st.error("Password minimal 6 karakter.")
                    else:
                        users[reg_user] = {
                            "password": hash_password(reg_pw), "nama": reg_nama, "role": reg_role,
                            "status": "pending", 
                            "created": datetime.now().strftime("%Y-%m-%d %H:%M")
                        }
                        save_users(users)
                        st.success("Pendaftaran berhasil! Akun Anda menunggu persetujuan Admin.")

        with tab_siswa:
            st.info("Isi kuesioner ini dengan jujur. Data akan masuk ke sistem guru Anda.")
            with st.form("form_kuesioner"):
                c1, c2, c3 = st.columns([2, 1, 1])
                k_nama = c1.text_input("Nama Lengkap")
                k_kelas = c2.selectbox("Kelas", KELAS_LIST)
                k_absen = c3.number_input("Absen", 1, 50, 1)
                
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
                        profil = {
                            "Self-Efficacy": f1, "Keterlibatan Ortu": f5, "Harapan Ortu": f6,
                            "Dukungan Sekolah": f7, "Motivasi": f2, "Kecemasan": f3,
                            "Fasilitas": f8, "Kemalasan": f4
                        }
                        ok, msg = update_kuesioner_siswa(k_nama, k_kelas, k_absen, profil)
                        if ok: st.success(msg); st.balloons()
                        else: st.error(msg)

# ================================================================
# ROUTING
# ================================================================
if not st.session_state.logged_in:
    halaman_login()
    st.stop()

# ================================================================
# HOME (KAMAR GURU)
# ================================================================
if st.session_state.current_page == "home":
    c1, c2 = st.columns([4, 1])
    with c1:
        st.markdown(f"<h2>Halo, <b>{st.session_state.user_nama}</b>! 👋</h2>", unsafe_allow_html=True)
        st.write(f"Selamat datang. Anda login sebagai **{st.session_state.user_role.capitalize()}**.")
    with c2:
        if st.button("Logout", type="primary"):
            add_log("Logout")
            st.session_state.logged_in = False
            st.session_state.user_nama = None
            st.rerun()

    st.markdown("---")
    
    df = pd.DataFrame(st.session_state.database_siswa)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f'<div class="glass-card metric-box"><div class="metric-val">{len(df)}</div><div class="metric-lbl">Siswa Anda</div></div>', unsafe_allow_html=True)
    with col2:
        avg = df["Nilai Akademik"].mean() if not df.empty else 0
        st.markdown(f'<div class="glass-card metric-box"><div class="metric-val">{avg:.2f}</div><div class="metric-lbl">Rata-rata Kelas</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="glass-card metric-box"><div class="metric-val">{len(st.session_state.log)}</div><div class="metric-lbl">Aktivitas Anda</div></div>', unsafe_allow_html=True)

    st.markdown("### Modul Anda")
    mods = [
        ("📝", "Input & Upload", "Input manual atau upload Excel nilai.", "analisis"),
        ("📊", "Data Siswa", "Lihat & kelola daftar siswa Anda.", "monitoring"),
        ("🔬", "Analisis Faktor", "Lihat dampak & prioritas per siswa.", "kausal"),
        ("💡", "Rekomendasi", "Saran tindak lanjut personal.", "rekomendasi"),
        ("⚙️", "Pengaturan", "Pengaturan sistem.", "pengaturan")
    ]
    
    if st.session_state.user_role == "admin":
        mods.append(("👥", "Manajemen Pengguna", "Setujui akun, reset password, & hapus akun.", "admin_users"))
        
    cols = st.columns(3)
    for i, (icon, title, desc, page) in enumerate(mods):
        with cols[i % 3]:
            st.markdown(f'<div class="glass-card" style="min-height:150px;"><h3>{icon} {title}</h3><p style="color:grey;font-size:0.9rem;">{desc}</p></div>', unsafe_allow_html=True)
            if st.button(f"Buka {title}", key=f"btn_{page}", use_container_width=True):
                goto_page(page)

# ================================================================
# MODUL ADMIN: MANAJEMEN PENGGUNA
# ================================================================
elif st.session_state.current_page == "admin_users":
    st.button("← Kembali", on_click=goto_page, args=("home",))
    st.markdown("## 👥 Manajemen Pengguna")
    
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
            with c1:
                st.write(f"**Role:** {d['role']}")
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
                            else:
                                st.error("Password minimal 6 karakter.")
                    with col_hapus:
                        if st.button(f"🗑️ Hapus Akun", key=f"del_{u}", use_container_width=True, type="primary"):
                            del users[u]
                            save_users(users)
                            st.success(f"Akun @{u} berhasil dihapus permanen.")
                            st.rerun()
                else:
                    st.info("Akun Super Admin (Terkunci).")

    st.markdown("---")
    st.markdown("### 📋 Pantau Semua Log Aktivitas (Global)")
    if st.button("Muat Semua Log Sistem"):
        all_logs = load_log(user=None) 
        if all_logs:
            df_log = pd.DataFrame(all_logs)[["waktu", "user", "role", "aksi"]]
            st.dataframe(df_log, use_container_width=True, hide_index=True)
        else:
            st.info("Belum ada log sama sekali.")

        # 4. Hapus Data Siswa (Kuesioner Salah)
    st.markdown("---")
    st.markdown("### 🗑️ Hapus Data Siswa (Kuesioner Salah)")
    st.write("Gunakan fitur ini jika ada siswa yang salah mengisi kuesioner.")
    
    all_siswa = load_all_siswa()
    if not all_siswa:
        st.info("Belum ada data siswa di sistem.")
    else:
        df_all = pd.DataFrame(all_siswa)
        st.dataframe(df_all[["Nama", "Kelas", "Absen", "Nilai Akademik", "Dicatat Oleh"]], use_container_width=True, hide_index=True)
        
        opsi_hapus = {s["_id"]: f"{s['Nama']} - {s['Kelas']} (Absen {s['Absen']}) - Input oleh: {s.get('Dicatat Oleh', 'Sistem')}" for s in all_siswa}
        
        siswa_id_terpilih = st.selectbox("Pilih siswa yang akan dihapus:", options=list(opsi_hapus.keys()), format_func=lambda x: opsi_hapus[x])
        
        if st.button("Hapus Data Siswa Ini", type="primary"):
            hapus_siswa_by_id(siswa_id_terpilih)
            st.success("Data siswa berhasil dihapus permanen dari database.")
            st.rerun()

                    # ZONA BERBAHAYA (HAPUS SEMUA)
        st.markdown("---")
        with st.expander("🚨 ZONA BERBAHAYA: Hapus SEMUA Data Siswa"):
            st.warning("Tindakan ini akan menghapus seluruh data siswa dari SEMUA guru secara permanen. Pastikan Anda sudah melakukan backup!")
            konfirmasi = st.checkbox("Saya mengerti dan ingin menghapus SEMUA data siswa.")
            
            if st.button("Hapus SEMUA Data Siswa", type="primary", disabled=not konfirmasi):
                jumlah = hapus_semua_siswa_admin()
                st.success(f"Berhasil menghapus {jumlah} data siswa dari database.")
                st.rerun()

# ================================================================
# MODUL 01: INPUT & UPLOAD (DIGABUNG)
# ================================================================
elif st.session_state.current_page == "analisis":
    st.button("← Kembali", on_click=goto_page, args=("home",))
    st.markdown("## 📝 Modul 01: Input Data Siswa")
    
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
                        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
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
            f1 = input_kategori("m_s1", "Self-Efficacy Akademik")
            f2 = input_kategori("m_s2", "Motivasi Belajar")
            f3 = input_kategori("m_s3", "Kecemasan Akademik")
            f4 = input_kategori("m_s4", "Kemalasan Belajar")
            f5 = input_kategori("m_s5", "Keterlibatan Orang Tua")
            f6 = input_kategori("m_s6", "Harapan Orang Tua")
            f7 = input_kategori("m_s7", "Dukungan Sekolah")
            f8 = input_kategori("m_s8", "Fasilitas Sekolah")
            
            if st.form_submit_button("Simpan Siswa Ini"):
                kat, _ = kategori_nilai(m_nilai)
                insert_siswa({
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
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
# MODUL 02: DATA SISWA (MONITORING)
# ================================================================
elif st.session_state.current_page == "monitoring":
    st.button("← Kembali", on_click=goto_page, args=("home",))
    st.markdown("## 📊 Data Siswa Anda")
    st.info("Data ini bersifat privat. Hanya Anda yang bisa melihat tabel ini.")
    
    df = pd.DataFrame(st.session_state.database_siswa)
    if df.empty:
        st.warning("Belum ada data siswa. Silakan input atau upload di Modul 01.")
    else:
        st.dataframe(df[["Nama", "Kelas", "Absen", "Nilai Akademik", "Kategori"]], use_container_width=True)
        
        with st.expander("🗑️ Zona Berbahaya"):
            if st.button("Hapus SEMUA Siswa Saya", type="primary"):
                hapus_semua_siswa(st.session_state.user_nama)
                reload_data()
                st.success("Data Anda berhasil dihapus.")
                st.rerun()

# ================================================================
# MODUL 03: FAKTOR PENGARUH (ANALISIS PER SISWA)
# ================================================================
elif st.session_state.current_page == "kausal":
    st.button("← Kembali", on_click=goto_page, args=("home",))
    st.markdown("## 🔬 Modul 03: Analisis Faktor Siswa")
    st.write("Lihat faktor apa saja yang paling mempengaruhi nilai siswa secara spesifik.")

    df_siswa = pd.DataFrame(st.session_state.database_siswa)
    if df_siswa.empty:
        st.warning("Belum ada data siswa. Silakan input data di Modul 01.")
        st.stop()

    # Pilih Siswa
    opsi_siswa = df_siswa.apply(lambda x: f"{x['Nama']} - {x['Kelas']} (Absen {x['Absen']})", axis=1).tolist()
    siswa_terpilih = st.selectbox("Pilih Siswa untuk Dianalisis", opsi_siswa)
    idx = opsi_siswa.index(siswa_terpilih)
    b = df_siswa.iloc[idx]

    nama = b["Nama"]; kelas = b["Kelas"]; absen = int(b["Absen"])
    nilai = float(b["Nilai Akademik"])
    profil = {
        "Self-Efficacy Akademik": float(b["Self-Efficacy"]),
        "Keterlibatan Orang Tua": float(b["Keterlibatan Ortu"]),
        "Harapan Orang Tua": float(b["Harapan Ortu"]),
        "Dukungan Sekolah": float(b["Dukungan Sekolah"]),
        "Motivasi Belajar": float(b["Motivasi"]),
        "Kecemasan Akademik": float(b["Kecemasan"]),
        "Fasilitas Sekolah": float(b["Fasilitas"]),
        "Kemalasan Belajar": float(b["Kemalasan"]),
    }

    selisih, kontribusi = analisis_kausal(nilai, profil)

    st.markdown(f"### Profil Analisis: **{nama}**")
    c1, c2 = st.columns(2)
    with c1:
        st.info(f"Nilai Rata-rata: **{nilai:.2f}**")
    with c2:
        st.info(f"Posisi vs Rata-rata Sekolah: **{selisih:+.2f} poin**")

    df_k = pd.DataFrame([
        {"Faktor": k, "Nilai Siswa": profil[k], "Rata-rata Sekolah": BASELINE_ASPEK[k],
         "Selisih": profil[k] - BASELINE_ASPEK[k], "Dampak (Poin)": v}
        for k, v in kontribusi.items()
    ]).sort_values("Dampak (Poin)")

    # Grafik batang horizontal
    st.markdown("**Grafik Dampak Faktor (Poin)**")
    st.bar_chart(df_k.set_index("Faktor")["Dampak (Poin)"], color="#6366f1")
    st.caption("Batang positif (ke kanan) menaikkan nilai, batang negatif (ke kiri) menurunkan nilai.")

    st.markdown("**Tabel Detail**")
    st.dataframe(df_k, use_container_width=True, hide_index=True)

# ================================================================
# MODUL 04: REKOMENDASI (CATATAN GURU & DOWNLOAD WORD)
# ================================================================
elif st.session_state.current_page == "rekomendasi":
    st.button("← Kembali", on_click=goto_page, args=("home",))
    st.markdown("## 💡 Modul 04: Rekomendasi & Catatan Guru")
    st.write("Saran tindak lanjut personal dan unduh laporan resmi dalam bentuk Word.")

    df_siswa = pd.DataFrame(st.session_state.database_siswa)
    if df_siswa.empty:
        st.warning("Belum ada data siswa. Silakan input data di Modul 01.")
        st.stop()

    opsi_siswa = df_siswa.apply(lambda x: f"{x['Nama']} - {x['Kelas']} (Absen {x['Absen']})", axis=1).tolist()
    siswa_terpilih = st.selectbox("Pilih Siswa untuk Rekomendasi", opsi_siswa)
    idx = opsi_siswa.index(siswa_terpilih)
    b = df_siswa.iloc[idx]

    nama = b["Nama"]; kelas = b["Kelas"]; absen = int(b["Absen"]); nilai = float(b["Nilai Akademik"])
    profil = {
        "Self-Efficacy Akademik": float(b["Self-Efficacy"]),
        "Keterlibatan Orang Tua": float(b["Keterlibatan Ortu"]),
        "Harapan Orang Tua": float(b["Harapan Ortu"]),
        "Dukungan Sekolah": float(b["Dukungan Sekolah"]),
        "Motivasi Belajar": float(b["Motivasi"]),
        "Kecemasan Akademik": float(b["Kecemasan"]),
        "Fasilitas Sekolah": float(b["Fasilitas"]),
        "Kemalasan Belajar": float(b["Kemalasan"]),
    }
    selisih, kontribusi = analisis_kausal(nilai, profil)
    kategori, warna_kategori = kategori_nilai(nilai)

    df_k = pd.DataFrame([
        {"Aspek": k, "Kontribusi": v, "Nilai_Siswa": profil[k],
         "Baseline": BASELINE_ASPEK[k],
         "Selisih": profil[k] - BASELINE_ASPEK[k]}
        for k, v in kontribusi.items()
    ]).sort_values("Kontribusi")
    neg = df_k[df_k["Kontribusi"] < -0.1].sort_values("Kontribusi")
    pos = df_k[df_k["Kontribusi"] > 0.1].sort_values("Kontribusi", ascending=False)

    posisi = "di atas" if selisih >= 0 else "di bawah"
    fn = neg.iloc[0] if len(neg) > 0 else None
    fp = pos.iloc[0] if len(pos) > 0 else None

    narasi_neg = (
        f"Faktor yang paling menekan nilai <b>{nama}</b> adalah "
        f"<b style='color:#EF5B67;'>{fn['Aspek']}</b> "
        f"(kontribusi {fn['Kontribusi']:.2f} poin)."
    ) if fn is not None else "Tidak ada faktor yang signifikan menekan nilai."

    narasi_pos = (
        f"Kekuatan utama terletak pada <b style='color:#10B981;'>{fp['Aspek']}</b> "
        f"(kontribusi +{fp['Kontribusi']:.2f} poin)."
    ) if fp is not None else "Belum ada faktor kekuatan dominan."

    st.markdown(
        f'<div class="glass-card" style="border-left:5px solid {warna_kategori};margin-top:1rem;">'
        f'<div style="display:flex;align-items:center;gap:1rem;margin-bottom:1rem;">'
        f'<div style="width:56px;height:56px;border-radius:15px;background:{warna_kategori};color:white;display:flex;align-items:center;justify-content:center;font-size:1.2rem;font-weight:800;">'
        f'{"".join(x[0] for x in nama.split()[:2]).upper()}</div>'
        f'<div><div style="font-weight:800;font-size:1.3rem;">{nama}</div>'
        f'<div style="color:#64748B;font-size:.8rem;">{kelas} - Absen {absen:02d} - Nilai {nilai:.2f}</div>'
        f'</div></div>'
        f'<div style="font-size:.95rem;line-height:1.75;color:#1E293B;">'
        f'Nilai saat ini <b>{abs(selisih):.2f} poin {posisi} rata-rata sekolah</b> ({RATA_RATA_NILAI:.2f}).'
        f'<br><br>{narasi_neg}<br><br>{narasi_pos}<br><br>'
        f'<b>Kategori:</b> <span style="color:{warna_kategori};font-weight:800;">{kategori}</span>'
        f'</div></div>',
        unsafe_allow_html=True,
    )

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
                st.markdown(
                    f'<div class="glass-card" style="border-top:4px solid #EF5B67;min-height:200px;">'
                    f'<div style="font-size:.65rem;font-weight:800;letter-spacing:1.5px;color:#EF5B67;">PRIORITAS {i + 1}</div>'
                    f'<div style="font-weight:800;font-size:1rem;margin:.6rem 0 .3rem;">{row["Aspek"]}</div>'
                    f'<div style="font-size:.7rem;color:#64748B;margin-bottom:.5rem;">{level}</div>'
                    f'<div style="font-size:.75rem;line-height:1.6;color:#475467;">'
                    f'Nilai siswa: <b>{row["Nilai_Siswa"]:.2f}</b><br>'
                    f'Rata-rata: <b>{row["Baseline"]:.2f}</b><br>'
                    f'<span style="color:#EF5B67;font-weight:700;">Selisih: {row["Selisih"]:+.2f}</span>'
                    f'</div></div>',
                    unsafe_allow_html=True,
                )

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
                    for j, t in enumerate(tugas["guru"], 1):
                        st.markdown(f"{j}. {t}")
                with c2:
                    st.markdown("**Yang bisa dilakukan SISWA:**")
                    for j, s in enumerate(tugas["siswa"], 1):
                        st.markdown(f"{j}. {s}")

         # ==========================================
    # FITUR DOWNLOAD LAPORAN (WORD, PDF, & ZIP)
    # ==========================================
    st.markdown("---")
    st.markdown("### 📄 Unduh Laporan Konsultasi")
    st.write("Laporan ini dirancang dengan bahasa yang mudah dipahami dan siap di-print.")

    # --- FUNGSI PEMBUAT DOKUMEN (Dibuat terpisah agar bisa dipakai berulang) ---
    def buat_dokumen_word(siswa_data):
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        doc = Document()
        
        header = doc.add_heading('LAPORAN PERKEMBANGAN SISWA', 0)
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subheader = doc.add_paragraph('Smart Academic Analytics (SAA) - SMPN 6 Salatiga')
        subheader.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()

        table = doc.add_table(rows=4, cols=2)
        table.style = 'Table Grid'
        table.rows[0].cells[0].text = 'Nama Siswa'
        table.rows[0].cells[1].text = siswa_data['nama']
        table.rows[1].cells[0].text = 'Kelas / Absen'
        table.rows[1].cells[1].text = f"{siswa_data['kelas']} / {siswa_data['absen']}"
        table.rows[2].cells[0].text = 'Nilai Rata-rata'
        table.rows[2].cells[1].text = f"{siswa_data['nilai']:.2f}"
        table.rows[3].cells[0].text = 'Kategori'
        table.rows[3].cells[1].text = siswa_data['kategori']

        doc.add_paragraph()
        doc.add_heading('Ringkasan Kondisi', level=1)
        posisi_teks = "di atas" if siswa_data['selisih'] >= 0 else "di bawah"
        doc.add_paragraph(f"Saat ini, {siswa_data['nama']} memiliki nilai rata-rata {siswa_data['nilai']:.2f}, yang berada {abs(siswa_data['selisih']):.2f} poin {posisi_teks} rata-rata sekolah. Berdasarkan analisis, {siswa_data['nama']} masuk dalam kategori {siswa_data['kategori']}.")

        doc.add_heading('Catatan untuk Orang Tua & Guru', level=1)
        if siswa_data['fp'] is not None:
            doc.add_paragraph(f"KEKUATAN UTAMA: {siswa_data['nama']} menunjukkan sikap yang sangat baik pada aspek {siswa_data['fp']['Aspek']}. Hal ini menjadi modal penting untuk terus berkembang.")
        
        if siswa_data['fn'] is not None:
            doc.add_paragraph(f"AREA YANG PERLU DIDAMPINGI: Untuk meningkatkan hasil belajar, kita bisa fokus membantu {siswa_data['nama']} pada aspek {siswa_data['fn']['Aspek']}. Berikut saran tindak lanjutnya:")
            
            saran_guru = siswa_data['tindak'].get(siswa_data['fn']['Aspek'], {}).get("guru", ["Berikan pendampingan ekstra."])
            saran_siswa = siswa_data['tindak'].get(siswa_data['fn']['Aspek'], {}).get("siswa", ["Belajar lebih giat lagi."])
            
            doc.add_paragraph("Saran untuk Guru/Sekolah:", style='List Bullet')
            for s in saran_guru[:3]:
                doc.add_paragraph(s, style='List Bullet 2')
            doc.add_paragraph("Saran untuk Siswa/Orang Tua di Rumah:", style='List Bullet')
            for s in saran_siswa[:2]:
                doc.add_paragraph(s, style='List Bullet 2')

        doc.add_paragraph()
        ttd_table = doc.add_table(rows=2, cols=2)
        ttd_table.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        ttd_table.rows[0].cells[0].text = 'Guru / Konselor,'
        ttd_table.rows[0].cells[1].text = f"Salatiga, {datetime.now().strftime('%d %B %Y')}"
        ttd_table.rows[1].cells[0].text = '\n\n\n\n(_______________________)'
        ttd_table.rows[1].cells[1].text = f'\n\n\n\n( {st.session_state.user_nama} )'

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    def buat_dokumen_pdf(siswa_data):
        from fpdf import FPDF
        pdf = FPDF()
        pdf.add_page()
        
        # Header
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, 'LAPORAN PERKEMBANGAN SISWA', ln=True, align='C')
        pdf.set_font("Arial", 'I', 10)
        pdf.cell(0, 10, 'Smart Academic Analytics (SAA) - SMPN 6 Salatiga', ln=True, align='C')
        pdf.ln(10)

        # Info Siswa
        pdf.set_font("Arial", 'B', 11)
        pdf.cell(40, 8, 'Nama Siswa', border=1)
        pdf.set_font("Arial", '', 11)
        pdf.cell(0, 8, siswa_data['nama'], border=1, ln=True)
        
        pdf.set_font("Arial", 'B', 11)
        pdf.cell(40, 8, 'Kelas / Absen', border=1)
        pdf.set_font("Arial", '', 11)
        pdf.cell(0, 8, f"{siswa_data['kelas']} / {siswa_data['absen']}", border=1, ln=True)
        
        pdf.set_font("Arial", 'B', 11)
        pdf.cell(40, 8, 'Nilai Rata-rata', border=1)
        pdf.set_font("Arial", '', 11)
        pdf.cell(0, 8, f"{siswa_data['nilai']:.2f}", border=1, ln=True)
        
        pdf.set_font("Arial", 'B', 11)
        pdf.cell(40, 8, 'Kategori', border=1)
        pdf.set_font("Arial", '', 11)
        pdf.cell(0, 8, siswa_data['kategori'], border=1, ln=True)
        pdf.ln(10)

        # Ringkasan
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, 'Ringkasan Kondisi', ln=True)
        pdf.set_font("Arial", '', 11)
        posisi_teks = "di atas" if siswa_data['selisih'] >= 0 else "di bawah"
        teks_ringkasan = f"Saat ini, {siswa_data['nama']} memiliki nilai rata-rata {siswa_data['nilai']:.2f}, yang berada {abs(siswa_data['selisih']):.2f} poin {posisi_teks} rata-rata sekolah. Berdasarkan analisis, {siswa_data['nama']} masuk dalam kategori {siswa_data['kategori']}."
        pdf.multi_cell(0, 6, teks_ringkasan)
        pdf.ln(5)

        # Catatan
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, 'Catatan untuk Orang Tua & Guru', ln=True)
        
        pdf.set_font("Arial", '', 11)
        if siswa_data['fp'] is not None:
            teks_kekuatan = f"KEKUATAN UTAMA: {siswa_data['nama']} menunjukkan sikap yang sangat baik pada aspek {siswa_data['fp']['Aspek']}."
            pdf.multi_cell(0, 6, teks_kekuatan)
            pdf.ln(2)
        
        if siswa_data['fn'] is not None:
            teks_area = f"AREA YANG PERLU DIDAMPINGI: Fokus pada aspek {siswa_data['fn']['Aspek']}."
            pdf.multi_cell(0, 6, teks_area)
            pdf.ln(2)
            
            saran_guru = siswa_data['tindak'].get(siswa_data['fn']['Aspek'], {}).get("guru", ["Berikan pendampingan ekstra."])
            saran_siswa = siswa_data['tindak'].get(siswa_data['fn']['Aspek'], {}).get("siswa", ["Belajar lebih giat lagi."])
            
            pdf.set_font("Arial", 'B', 11)
            pdf.cell(0, 6, 'Saran untuk Guru/Sekolah:', ln=True)
            pdf.set_font("Arial", '', 11)
            for s in saran_guru[:3]:
                pdf.multi_cell(0, 6, f"- {s}")
                
            pdf.set_font("Arial", 'B', 11)
            pdf.cell(0, 6, 'Saran untuk Siswa/Orang Tua di Rumah:', ln=True)
            pdf.set_font("Arial", '', 11)
            for s in saran_siswa[:2]:
                pdf.multi_cell(0, 6, f"- {s}")

        # Tanda Tangan
        pdf.ln(15)
        pdf.cell(100, 8, 'Guru / Konselor,', ln=False)
        pdf.cell(0, 8, f"Salatiga, {datetime.now().strftime('%d %B %Y')}", ln=True)
        pdf.ln(15)
        pdf.cell(100, 8, '(_______________________)', ln=False)
        pdf.cell(0, 8, f"( {st.session_state.user_nama} )", ln=True)

        # Output PDF
        pdf_output = pdf.output(dest='S').encode('latin-1') # fpdf2 butuh format bytes
        return io.BytesIO(pdf_output)

    # Siapkan Data Siswa
    data_laporan = {
        'nama': nama, 'kelas': kelas, 'absen': absen, 'nilai': nilai,
        'kategori': kategori, 'selisih': selisih, 'fp': fp, 'fn': fn, 'tindak': TINDAK
    }

    # --- TOMBOL DOWNLOAD INDIVIDU ---
    col_word, col_pdf = st.columns(2)
    
    with col_word:
        if st.button("📥 Download Laporan (Word)", type="primary", use_container_width=True):
            try:
                file_word = buat_dokumen_word(data_laporan)
                st.download_button(
                    label="⬇️ Klik di Sini untuk Simpan Word",
                    data=file_word,
                    file_name=f"Laporan_{nama.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Gagal membuat Word: {e}")
                
    with col_pdf:
        if st.button("📥 Download Laporan (PDF)", type="primary", use_container_width=True):
            try:
                file_pdf = buat_dokumen_pdf(data_laporan)
                st.download_button(
                    label="⬇️ Klik di Sini untuk Simpan PDF",
                    data=file_pdf,
                    file_name=f"Laporan_{nama.replace(' ', '_')}.pdf",
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"Gagal membuat PDF: {e}")

    # --- TOMBOL DOWNLOAD SEMUA (ZIP) ---
    st.markdown("---")
    st.markdown("### 🗂️ Unduh Semua Laporan (Massal)")
    st.write("Unduh laporan seluruh siswa Anda dalam satu file ZIP.")
    
    if st.button("📦 Buat File ZIP (Semua Siswa)", use_container_width=True):
        import zipfile
        try:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for _, row in df_siswa.iterrows():
                    # Ambil data per siswa
                    s_nama = row["Nama"]
                    s_profil = {
                        "Self-Efficacy Akademik": float(row["Self-Efficacy"]),
                        "Keterlibatan Orang Tua": float(row["Keterlibatan Ortu"]),
                        "Harapan Orang Tua": float(row["Harapan Ortu"]),
                        "Dukungan Sekolah": float(row["Dukungan Sekolah"]),
                        "Motivasi Belajar": float(row["Motivasi"]),
                        "Kecemasan Akademik": float(row["Kecemasan"]),
                        "Fasilitas Sekolah": float(row["Fasilitas"]),
                        "Kemalasan Belajar": float(row["Kemalasan"]),
                    }
                    s_selisih, s_kontribusi = analisis_kausal(float(row["Nilai Akademik"]), s_profil)
                    s_kat, _ = kategori_nilai(float(row["Nilai Akademik"]))
                    
                    s_df_k = pd.DataFrame([
                        {"Aspek": k, "Kontribusi": v} for k, v in s_kontribusi.items()
                    ]).sort_values("Kontribusi")
                    s_neg = s_df_k[s_df_k["Kontribusi"] < -0.1]
                    s_pos = s_df_k[s_df_k["Kontribusi"] > 0.1].sort_values("Kontribusi", ascending=False)
                    
                    s_data = {
                        'nama': s_nama, 'kelas': row["Kelas"], 'absen': int(row["Absen"]),
                        'nilai': float(row["Nilai Akademik"]), 'kategori': s_kat,
                        'selisih': s_selisih, 
                        'fp': s_pos.iloc[0] if len(s_pos) > 0 else None,
                        'fn': s_neg.iloc[0] if len(s_neg) > 0 else None, 
                        'tindak': TINDAK
                    }
                    
                    # Buat dokumen Word (bisa diganti PDF jika mau)
                    file_buffer = buat_dokumen_word(s_data)
                    
                    # Masukkan ke dalam ZIP
                    nama_file = f"Laporan_{s_nama.replace(' ', '_')}.docx"
                    zip_file.writestr(nama_file, file_buffer.getvalue())
            
            zip_buffer.seek(0)
            st.success(f"Berhasil membuat ZIP berisi {len(df_siswa)} laporan Word!")
            st.download_button(
                label="⬇️ Download Semua Laporan (.zip)",
                data=zip_buffer,
                file_name="Semua_Laporan_Siswa.zip",
                mime="application/zip"
            )
        except Exception as e:
            st.error(f"Gagal membuat ZIP: {e}")

# ================================================================
# MODUL 05: PENGATURAN
# ================================================================
elif st.session_state.current_page == "pengaturan":
    st.button("← Kembali", on_click=goto_page, args=("home",))
    st.markdown("## ⚙️ Pengaturan & Profil")
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Profil Anda")
        st.info(f"**Nama:** {st.session_state.user_nama}\n\n**Role:** {st.session_state.user_role.capitalize()}")
        if st.button("Logout dari Perangkat Ini"):
            add_log("Logout")
            st.session_state.logged_in = False
            st.session_state.user_nama = None
            st.rerun()
            
    with c2:
        st.markdown("### Log Aktivitas Terakhir (Milik Anda)")
        if st.session_state.log:
            df_log = pd.DataFrame(st.session_state.log[:5])[["waktu", "aksi"]]
            st.dataframe(df_log, use_container_width=True, hide_index=True)
        else:
            st.write("Belum ada aktivitas.")

# ================================================================
# END OF FILE
# ================================================================