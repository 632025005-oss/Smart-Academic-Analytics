# -*- coding: utf-8 -*-
# ================================================================
# SMART ACADEMIC ANALYTICS (SAA) - SMP Negeri 6 Salatiga
# Versi: Guru Upload Nilai (Excel) + Siswa Isi Kuesioner (Form)
# ================================================================

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import hashlib
import json
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# ================================================================
# MONGODB STORAGE
# ================================================================
from db_storage import (
    load_users as _db_load_users,
    save_users as _db_save_users,
    load_database, save_database,
    load_config, save_config,
    load_log, save_log, tambah_log,
    ambil_semua_hasil, hapus_hasil,
    insert_siswa, hapus_semua_siswa,
    update_kuesioner_siswa  # Fungsi baru
)

SUPER_ADMIN = {"admin", "regina"}

# ================================================================
# HELPER: SECRETS & PASSWORD
# ================================================================
def get_secret(section, key=None, default=None):
    try:
        if key is None:
            return st.secrets[section]
        return st.secrets[section][key]
    except (KeyError, FileNotFoundError):
        return default

def hash_password(p):
    return hashlib.sha256(p.encode()).hexdigest()

def verify_password(p, hashed):
    return hash_password(p) == hashed

def load_users():
    users = _db_load_users()
    if not users:
        super_admins = get_secret("super_admin", None, {})
        if not super_admins:
            raise RuntimeError("❌ Kredensial super admin tidak ditemukan di secrets.")
        users = {}
        for username, info in super_admins.items():
            users[username] = {
                "password": hash_password(info["password"]),
                "nama":     info.get("nama", username.capitalize()),
                "role":     "admin",
                "kelas_ampu": None,
                "email":    info.get("email", ""),
                "status":   "approved",
                "created":  datetime.now().strftime("%Y-%m-%d %H:%M"),
            }
        _db_save_users(users)
    return users

def save_users(u):
    _db_save_users(u)

# ================================================================
# PAGE CONFIG
# ================================================================
st.set_page_config(
    page_title="SAA - Smart Academic Analytics",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ================================================================
# CSS (Tampilan UI)
# ================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=Manrope:wght@600;700;800&family=Plus+Jakarta+Sans:wght@700;800&display=swap');
:root{
    --ink:#0F1B33; --muted:#5C6B85; --line:#E4EAF2;
    --blue:#2563EB; --blue-soft:#EAF1FF;
    --violet:#7C5CFC; --violet-soft:#F1EDFF;
    --yellow:#F6C945; --yellow-soft:#FFF7D6;
    --mint:#10B981; --mint-soft:#E8F8F1;
    --coral:#EF5B67; --coral-soft:#FFF0F2;
    --pink:#EC4899; --pink-soft:#FCE7F3;
    --cyan:#06B6D4; --cyan-soft:#CFFAFE;
}
html, body, [class*="css"]{font-family:'DM Sans',sans-serif;color:#0F1B33;}
.stApp{
    background:
        radial-gradient(circle at 100% 0%, rgba(37,99,235,.07), transparent 32rem),
        radial-gradient(circle at 0% 100%, rgba(124,92,252,.06), transparent 32rem),
        #FFFFFF;
}
.main{background:transparent;}
.block-container{max-width:1480px;padding:1.6rem 2.7rem 4rem 2.7rem;}
#MainMenu, footer, header{visibility:hidden;}
h1,h2,h3,h4{font-family:'Manrope',sans-serif;}

@keyframes riseIn{from{opacity:0;transform:translateY(14px);}to{opacity:1;transform:translateY(0);}}
@keyframes floatDot{0%,100%{transform:translateY(0);}50%{transform:translateY(-8px);}}
@keyframes sparkle{0%,100%{opacity:.3;transform:scale(1);}50%{opacity:1;transform:scale(1.35);}}
.motion{animation:riseIn .5s ease both;}

.login-page-header{display:flex;justify-content:space-between;align-items:center;padding:1rem 2rem;
    background:linear-gradient(135deg,#FFFFFF 0%,#F3F6FF 100%);border-bottom:1px solid #DCE4F2;
    margin:-1.6rem -2.7rem 0 -2.7rem;box-shadow:0 4px 20px rgba(37,99,235,.06);}
.login-logo-area{display:flex;align-items:center;gap:.8rem;justify-content:flex-end;width:100%;}
.login-logo-icon{font-size:2.2rem;}
.login-logo-text{text-align:right;line-height:1.15;}
.login-logo-title{font-family:'Plus Jakarta Sans',sans-serif;font-size:1.6rem;font-weight:800;color:#0F1B33;letter-spacing:-.8px;}
.login-logo-title .blue-part{background:linear-gradient(135deg,#2563EB 0%,#7C5CFC 100%);
    -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;}
.login-logo-sub{font-size:.72rem;color:#5C6B85;margin-top:.15rem;}
.login-content{max-width:1050px;margin:2rem auto;padding:0 2rem;}
.login-date-logout{display:flex;justify-content:space-between;align-items:center;padding:.9rem 1.1rem;
    background:linear-gradient(135deg,#FFFFFF 0%,#F5F8FF 100%);border:1px solid #DCE4F2;
    border-left:5px solid #2563EB;border-radius:12px;margin-bottom:2rem;}
.login-date{font-family:'Manrope',sans-serif;font-weight:800;font-size:.95rem;color:#0F1B33;}
.siasat-label{font-weight:700;font-size:.9rem;color:#0F1B33;padding-top:.65rem;}
.siasat-info-box{background:linear-gradient(135deg,#FFF9E6 0%,#FFF4CC 100%);
    border:1px solid #F3DF8D;border-left:5px solid #F6C945;border-radius:12px;padding:1.2rem 1.4rem;margin-top:2.5rem;}
.siasat-info-header{display:flex;align-items:center;gap:.6rem;margin-bottom:.7rem;}
.siasat-info-title{font-family:'Manrope',sans-serif;font-weight:800;font-size:.9rem;color:#0F1B33;}
.siasat-info-list{font-size:.8rem;color:#475467;line-height:1.85;}
.siasat-info-list div{display:flex;gap:.5rem;}
.siasat-info-list .num{color:#2563EB;font-weight:800;min-width:18px;}
.siasat-footer{text-align:center;padding:2rem 0;margin-top:3rem;border-top:1px solid #DCE4F2;
    font-size:.72rem;color:#5C6B85;line-height:1.8;}
.siasat-footer strong{color:#2563EB;font-weight:800;}

.dashboard-hero{position:relative;overflow:hidden;
    background:linear-gradient(135deg,#FFFFFF 0%,#F5F8FF 50%,#EEF3FF 100%);
    border:1px solid #DCE4F2;border-radius:28px;padding:2.3rem 2.5rem;
    box-shadow:0 20px 55px rgba(30,50,90,.10);animation:riseIn .5s ease both;}

.menu-hero{position:relative;overflow:hidden;
    background:linear-gradient(135deg,#0B1730 0%,#16294A 30%,#1E3A8A 60%,#7C5CFC 100%);
    border-radius:32px;padding:3.5rem 3rem;color:#fff;margin-bottom:2.5rem;
    box-shadow:0 30px 80px rgba(11,23,48,.40);animation:riseIn .5s ease both;}
.menu-hero-kicker{display:inline-flex;align-items:center;gap:8px;font-size:.7rem;font-weight:800;
    letter-spacing:2px;color:#F6C945;text-transform:uppercase;margin-bottom:1rem;}
.menu-hero-title{font-family:'Plus Jakarta Sans',sans-serif;
    font-size:clamp(2.5rem,5vw,4.5rem);line-height:.98;letter-spacing:-3px;font-weight:800;margin:0;max-width:900px;}
.menu-hero-title em{background:linear-gradient(135deg,#F6C945 0%,#FFA94D 50%,#F6C945 100%);
    -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;font-style:italic;}
.menu-hero-sub{color:#C6D1E5;font-size:1.05rem;line-height:1.6;max-width:680px;margin:1.2rem 0 0;}
.menu-hero-meta{display:flex;gap:.7rem;flex-wrap:wrap;margin-top:1.8rem;}
.menu-hero-meta span{padding:.55rem .9rem;border-radius:999px;
    background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);
    color:#F0F5FF;font-size:.68rem;font-weight:800;letter-spacing:.7px;}

.menu-card{position:relative;overflow:hidden;
    background:linear-gradient(135deg,#FFFFFF 0%,#FBFDFF 100%);
    border:1px solid #DCE4F2;border-radius:24px;padding:1.75rem 1.6rem 1.6rem;
    transition:transform .3s ease, box-shadow .3s ease;
    box-shadow:0 14px 40px rgba(30,50,90,.08);min-height:230px;
    display:flex;flex-direction:column;justify-content:space-between;animation:riseIn .55s ease both;}
.menu-card:hover{transform:translateY(-8px);box-shadow:0 28px 65px rgba(30,50,90,.18);}
.menu-card.mc-blue{background:linear-gradient(135deg,#FFFFFF 0%,#EAF1FF 100%);border-color:#D4E0FF;}
.menu-card.mc-yellow{background:linear-gradient(135deg,#FFFFFF 0%,#FFF7D6 100%);border-color:#F3DF8D;}
.menu-card.mc-mint{background:linear-gradient(135deg,#FFFFFF 0%,#E8F8F1 100%);border-color:#BCE8D7;}
.menu-card.mc-purple{background:linear-gradient(135deg,#FFFFFF 0%,#F1EDFF 100%);border-color:#DDD5FF;}
.menu-card.mc-pink{background:linear-gradient(135deg,#FFFFFF 0%,#FCE7F3 100%);border-color:#F8C7DF;}
.menu-card.mc-teal{background:linear-gradient(135deg,#FFFFFF 0%,#CFFAFE 100%);border-color:#A5E8EF;}
.menu-card-num{font-family:'Plus Jakarta Sans',sans-serif;font-size:.68rem;font-weight:800;
    letter-spacing:1.6px;color:#5C6B85;text-transform:uppercase;}
.menu-card-icon{width:56px;height:56px;border-radius:16px;display:flex;align-items:center;
    justify-content:center;font-size:1.7rem;margin:.9rem 0;}
.menu-card.mc-blue .menu-card-icon{background:linear-gradient(135deg,#DBE7FF,#B8CEFF);}
.menu-card.mc-yellow .menu-card-icon{background:linear-gradient(135deg,#FFF1B8,#FFE37A);}
.menu-card.mc-mint .menu-card-icon{background:linear-gradient(135deg,#C9F2E0,#9BE5C8);}
.menu-card.mc-purple .menu-card-icon{background:linear-gradient(135deg,#E2DBFF,#C9BEFF);}
.menu-card.mc-pink .menu-card-icon{background:linear-gradient(135deg,#FBD5E8,#F8AFD0);}
.menu-card.mc-teal .menu-card-icon{background:linear-gradient(135deg,#B6EFF6,#7FE0EC);}
.menu-card-title{font-family:'Manrope',sans-serif;font-size:1.15rem;font-weight:800;
    letter-spacing:-.4px;line-height:1.2;margin:0 0 .35rem;}
.menu-card-desc{color:#5C6B85;font-size:.78rem;line-height:1.55;margin:0;}
.menu-card-cta{display:inline-flex;align-items:center;gap:.45rem;
    font-size:.72rem;font-weight:800;letter-spacing:.5px;margin-top:1rem;text-transform:uppercase;color:#2563EB;}

.section-wrap{margin-top:2rem;margin-bottom:1rem;}
.section-number{display:inline-flex;width:42px;height:42px;align-items:center;justify-content:center;
    border-radius:14px;background:linear-gradient(135deg,#2563EB 0%,#7C5CFC 100%);
    color:#fff;font-size:.8rem;font-weight:800;box-shadow:0 10px 24px rgba(37,99,235,.35);}
.section-title{font-size:1.45rem;letter-spacing:-.7px;margin:0;}
.section-sub{color:#5C6B85;font-size:.72rem;text-transform:uppercase;
    letter-spacing:1.1px;font-weight:800;margin-top:.2rem;}
.section-line{height:1px;background:linear-gradient(90deg,#DCE4F2 0%,transparent 100%);margin-top:.9rem;}

.card{background:linear-gradient(135deg,#FFFFFF 0%,#FBFDFF 100%);
    border:1px solid #E4EAF2;border-radius:20px;padding:1.25rem 1.35rem;
    box-shadow:0 10px 30px rgba(30,50,90,.06);animation:riseIn .5s ease both;}
.card-blue{background:linear-gradient(135deg,#FFFFFF 0%,#EAF1FF 100%);border-color:#D4E0FF;}
.card-yellow{background:linear-gradient(135deg,#FFFFFF 0%,#FFF7D6 100%);border-color:#F3DF8D;}
.card-mint{background:linear-gradient(135deg,#FFFFFF 0%,#E8F8F1 100%);border-color:#BCE8D7;}
.card-dark{background:linear-gradient(135deg,#0B1730 0%,#16294A 100%);
    border-color:#0B1730;color:#fff;box-shadow:0 20px 50px rgba(11,23,48,.35);}
.card-label{font-size:.65rem;font-weight:800;text-transform:uppercase;letter-spacing:1px;color:#5C6B85;}
.card-dark .card-label{color:#9DBBFF;}
.card-value{font-family:'Manrope',sans-serif;font-size:2rem;font-weight:800;
    letter-spacing:-1.2px;line-height:1;margin-top:.5rem;}
.card-note{color:#5C6B85;font-size:.74rem;margin-top:.45rem;}
.card-dark .card-note{color:#9DBBFF;}

.stat-card{min-height:122px;position:relative;overflow:hidden;}
.stat-card .topline{width:40px;height:5px;border-radius:99px;margin-bottom:.85rem;}
.topline.blue{background:linear-gradient(90deg,#2563EB,#7C5CFC);}
.topline.yellow{background:linear-gradient(90deg,#F6C945,#FFA94D);}
.topline.mint{background:linear-gradient(90deg,#10B981,#06B6D4);}
.topline.coral{background:linear-gradient(90deg,#EF5B67,#EC4899);}
.topline.purple{background:linear-gradient(90deg,#7C5CFC,#EC4899);}

.eyebrow{display:inline-flex;align-items:center;gap:8px;font-size:.68rem;font-weight:800;
    letter-spacing:1.5px;text-transform:uppercase;color:#2563EB;margin-bottom:.7rem;}
.eyebrow-dot{width:8px;height:8px;background:#F6C945;border-radius:50%;display:inline-block;}

.hero-title{font-family:'Plus Jakarta Sans',sans-serif;
    font-size:clamp(2.15rem,4vw,3.8rem);line-height:1.02;letter-spacing:-2.4px;margin:0;max-width:850px;}
.hero-title .accent{background:linear-gradient(135deg,#2563EB 0%,#7C5CFC 50%,#EC4899 100%);
    -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;}
.hero-sub{color:#5C6B85;font-size:.98rem;line-height:1.65;max-width:790px;margin:.9rem 0 0;}
.hero-meta{display:flex;flex-wrap:wrap;gap:8px;margin-top:1.25rem;}
.pill{display:inline-flex;align-items:center;gap:7px;padding:.42rem .7rem;border-radius:999px;
    font-size:.65rem;font-weight:800;letter-spacing:.5px;border:1px solid #E4EAF2;background:#fff;}
.pill.blue{background:#EAF1FF;color:#2563EB;border-color:#D4E0FF;}
.pill.yellow{background:#FFF7D6;color:#936F00;border-color:#F3DF8D;}
.pill.mint{background:#E8F8F1;color:#087453;border-color:#BCE8D7;}
.pill.purple{background:#F1EDFF;color:#5B3FCC;border-color:#DDD5FF;}
.pill.coral{background:#FFF0F2;color:#B91C1C;border-color:#F4CDD3;}

.profile-card{display:flex;align-items:center;justify-content:space-between;gap:1rem;
    padding:1.15rem 1.3rem;background:linear-gradient(135deg,#FFFFFF 0%,#F5F8FF 100%);
    border:1px solid #DCE4F2;border-radius:18px;box-shadow:0 8px 24px rgba(30,50,90,.05);}
.profile-avatar{width:48px;height:48px;border-radius:15px;display:flex;align-items:center;
    justify-content:center;background:linear-gradient(135deg,#2563EB 0%,#7C5CFC 50%,#EC4899 100%);
    color:#fff;font-weight:800;font-size:1rem;}
.profile-name{font-family:'Manrope',sans-serif;font-weight:800;font-size:1.05rem;}
.profile-meta{color:#5C6B85;font-size:.72rem;margin-top:.2rem;}

.factor-card{border:1px solid #E4EAF2;border-radius:18px;
    background:linear-gradient(135deg,#FFFFFF 0%,#FBFDFF 100%);
    padding:1rem 1.05rem;margin:.65rem 0;}
.factor-head{display:flex;justify-content:space-between;align-items:center;gap:1rem;}
.factor-name{font-weight:700;font-size:.86rem;}
.factor-value{font-family:'Manrope',sans-serif;font-weight:800;font-size:1rem;}
.factor-bar{height:7px;background:#EEF2F7;border-radius:99px;overflow:hidden;margin-top:.75rem;}
.factor-fill{height:100%;border-radius:99px;background:linear-gradient(90deg,#7C5CFC,#EC4899);}

.rank-row{display:grid;grid-template-columns:34px 1fr auto;
    align-items:center;gap:.8rem;padding:.75rem 0;border-bottom:1px solid #E4EAF2;}
.rank-row:last-child{border-bottom:0;}
.rank-num{width:30px;height:30px;border-radius:10px;
    background:linear-gradient(135deg,#EAF1FF 0%,#F1EDFF 100%);
    color:#2563EB;display:flex;align-items:center;justify-content:center;
    font-size:.68rem;font-weight:800;}
.rank-name{font-weight:700;font-size:.8rem;}
.rank-desc{font-size:.67rem;color:#5C6B85;margin-top:.15rem;}
.rank-value{font-family:'Manrope',sans-serif;font-size:.86rem;font-weight:800;}
.positive{color:#10B981;}
.negative{color:#EF5B67;}

.info-box{border-radius:18px;padding:1.1rem 1.2rem;border:1px solid #E4EAF2;
    background:linear-gradient(135deg,#FFFFFF 0%,#FBFDFF 100%);}
.info-box.blue{background:linear-gradient(135deg,#EAF1FF 0%,#DBE7FF 100%);border-color:#D5E1FF;}
.info-box.yellow{background:linear-gradient(135deg,#FFF7D6 0%,#FFF1B8 100%);border-color:#F1DF96;}
.info-box.mint{background:linear-gradient(135deg,#E8F8F1 0%,#C9F2E0 100%);border-color:#C5EBDD;}
.info-box.coral{background:linear-gradient(135deg,#FFF0F2 0%,#FFD9DF 100%);border-color:#F4CDD3;}
.info-title{font-weight:800;font-size:.9rem;}
.info-text{font-size:.78rem;line-height:1.6;margin-top:.35rem;color:#475467;}

.notif-success{background:linear-gradient(135deg,#ECFDF5 0%,#D1FAE5 100%);
    border:2px solid #10B981;border-left:6px solid #10B981;
    border-radius:14px;padding:1.2rem 1.4rem;margin:1rem 0;}
.notif-success-title{font-family:'Manrope',sans-serif;font-weight:800;font-size:1rem;color:#065F46;}
.notif-success-body{font-size:.85rem;color:#064E3B;line-height:1.6;margin-top:.4rem;}

.role-badge{display:inline-flex;align-items:center;gap:6px;padding:.35rem .75rem;
    border-radius:999px;font-size:.65rem;font-weight:800;letter-spacing:.5px;
    background:linear-gradient(135deg,#EAF1FF,#F1EDFF);color:#2563EB;border:1px solid #D4E0FF;}
.role-badge.admin{background:linear-gradient(135deg,#FCE7F3,#FBCFE8);color:#BE185D;border-color:#F8C7DF;}
.role-badge.kepala_sekolah{background:linear-gradient(135deg,#FEF3C7,#FDE68A);color:#92400E;border-color:#FCD34D;}
.role-badge.wali_kelas{background:linear-gradient(135deg,#D1FAE5,#A7F3D0);color:#065F46;border-color:#6EE7B7;}
.role-badge.guru{background:linear-gradient(135deg,#DBEAFE,#BFDBFE);color:#1E40AF;border-color:#93C5FD;}

.status-badge{display:inline-flex;align-items:center;gap:5px;padding:.3rem .6rem;
    border-radius:999px;font-size:.65rem;font-weight:800;}
.status-badge.approved{background:#D1FAE5;color:#065F46;}
.status-badge.pending{background:#FEF3C7;color:#92400E;}
.status-badge.rejected{background:#FEE2E2;color:#991B1B;}

@media(max-width:850px){
    .block-container{padding:1rem 1rem 3rem;}
    .menu-hero{padding:2.5rem 1.7rem;}
    .login-logo-title{font-size:1.2rem;}
}
</style>
""", unsafe_allow_html=True)

# ================================================================
# ROLE & SESSION
# ================================================================
ROLE_LABEL = {
    "admin": "Administrator",
    "kepala_sekolah": "Kepala Sekolah",
    "wali_kelas": "Wali Kelas",
    "guru": "Guru",
}

def can_access_pengaturan(role): return role == "admin"

for key, val in {
    "logged_in": False, "user_nama": None, "user_role": None, "user_kelas": None,
    "database_siswa": load_database(),
    "config": load_config(),
    "log": load_log(),
    "users": load_users(),
    "current_page": "home",
    "last_saved": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

def add_log(aksi, detail=""):
    entry = {
        "waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "user": st.session_state.user_nama or "-",
        "role": st.session_state.user_role or "-",
        "aksi": aksi, "detail": detail,
    }
    st.session_state.log.insert(0, entry)
    st.session_state.log = st.session_state.log[:200]
    save_log(st.session_state.log)

def reload_users():
    st.session_state.users = load_users()

# ================================================================
# DATA 8 FAKTOR
# ================================================================
BASELINE_ASPEK = {
    "Self-Efficacy Akademik": 4.634,
    "Keterlibatan Orang Tua": 4.348,
    "Harapan Orang Tua": 3.450,
    "Dukungan Sekolah": 4.600,
    "Motivasi Belajar": 2.517,
    "Kecemasan Akademik": 2.893,
    "Kemalasan Belajar": 1.487,
    "Fasilitas Sekolah": 4.880,
}
PENGARUH_DATA = {
    "Self-Efficacy Akademik": 2.1995,
    "Keterlibatan Orang Tua": 1.7559,
    "Harapan Orang Tua": -0.6987,
    "Dukungan Sekolah": -1.0351,
    "Motivasi Belajar": -1.9093,
    "Kecemasan Akademik": 0.7197,
    "Kemalasan Belajar": 0.7320,
    "Fasilitas Sekolah": 0.6182,
}
KEPENTINGAN_DATA = {
    "Self-Efficacy Akademik": 0.8197,
    "Keterlibatan Orang Tua": 0.6717,
    "Harapan Orang Tua": 0.5766,
    "Dukungan Sekolah": 0.4325,
    "Motivasi Belajar": 0.3601,
    "Kecemasan Akademik": 0.1553,
    "Fasilitas Sekolah": 0.0654,
    "Kemalasan Belajar": 0.0846,
}
BOBOT_PENGARUH = {k: PENGARUH_DATA[k] * KEPENTINGAN_DATA[k] for k in PENGARUH_DATA}
RATA_RATA_NILAI = 83.78
STD_NILAI = 3.64

KELAS_LIST = (
    [f"VII-{x}" for x in "ABCDEFGH"]
    + [f"VIII-{x}" for x in "ABCDEFGH"]
    + [f"IX-{x}" for x in "ABCDEFGH"]
)

SKALA_PILIHAN = {
    "Self-Efficacy Akademik": {"question": "Seberapa yakin kamu terhadap kemampuan akademikmu?", "options": [
        ("1","Hampir tidak pernah percaya diri"),("2","Jarang percaya diri"),
        ("3","Kadang-kadang percaya diri"),("4","Sering percaya diri"),("5","Hampir selalu percaya diri")]},
    "Keterlibatan Orang Tua": {"question": "Seberapa aktif orang tua mendampingi kamu belajar?", "options": [
        ("1","Hampir tidak pernah mendampingi"),("2","Jarang mendampingi"),
        ("3","Kadang-kadang mendampingi"),("4","Sering mendampingi"),("5","Hampir selalu mendampingi")]},
    "Harapan Orang Tua": {"question": "Seberapa tinggi tuntutan orang tua terhadap nilaimu?", "options": [
        ("1","Sangat rendah"),("2","Rendah"),("3","Sedang / wajar"),
        ("4","Tinggi"),("5","Sangat tinggi / menekan")]},
    "Dukungan Sekolah": {"question": "Seberapa besar dukungan sekolah kepadamu?", "options": [
        ("1","Sangat kurang"),("2","Kurang"),("3","Cukup"),("4","Baik"),("5","Sangat baik")]},
    "Motivasi Belajar": {"question": "Seberapa besar motivasi dan semangat belajarmu?", "options": [
        ("1","Tidak ada motivasi"),("2","Kurang termotivasi"),("3","Cukup termotivasi"),
        ("4","Termotivasi"),("5","Sangat termotivasi")]},
    "Kecemasan Akademik": {"question": "Seberapa sering kamu merasa cemas saat belajar atau ujian?", "options": [
        ("1","Hampir tidak pernah cemas"),("2","Jarang cemas"),("3","Kadang-kadang cemas"),
        ("4","Sering cemas"),("5","Hampir selalu cemas")]},
    "Fasilitas Sekolah": {"question": "Seberapa memadai fasilitas belajar di sekolahmu?", "options": [
        ("1","Sangat kurang memadai"),("2","Kurang memadai"),("3","Cukup memadai"),
        ("4","Memadai"),("5","Sangat memadai")]},
    "Kemalasan Belajar": {"question": "Seberapa sering kamu menunjukkan sikap malas belajar?", "options": [
        ("1","Tidak pernah malas"),("2","Jarang malas"),("3","Kadang-kadang malas"),
        ("4","Sering malas"),("5","Hampir selalu malas")]},
}

# ================================================================
# FUNGSI ANALITIK
# ================================================================
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

def prediksi_nilai(profil_siswa):
    pred = RATA_RATA_NILAI
    for aspek, nilai in profil_siswa.items():
        pred += (nilai - BASELINE_ASPEK[aspek]) * PENGARUH_DATA[aspek] * 0.35
    return max(60, min(100, pred))

def potensi_maksimal(profil_siswa):
    optimum = dict(profil_siswa)
    for aspek in profil_siswa:
        if PENGARUH_DATA[aspek] > 0:
            optimum[aspek] = 5.0
        else:
            optimum[aspek] = 1.0
    return prediksi_nilai(optimum)

def kategori_nilai(nilai):
    if nilai >= 88: return "Sangat Baik", "#10B981"
    if nilai >= 84: return "Baik", "#2563EB"
    if nilai >= 80: return "Cukup", "#F0B900"
    return "Perlu Perhatian", "#EF5B67"

def label_pengaruh(nilai):
    if nilai > 2: return "Sangat kuat menaikkan nilai"
    if nilai > 1: return "Kuat menaikkan nilai"
    if nilai > 0.5: return "Sedikit menaikkan nilai"
    if nilai > -0.5: return "Hampir tidak berpengaruh"
    if nilai > -1: return "Sedikit menurunkan nilai"
    if nilai > -2: return "Sedang menurunkan nilai"
    return "Kuat menurunkan nilai"

def label_kepentingan(nilai):
    if nilai > 0.7: return "Paling menentukan prediksi"
    if nilai > 0.4: return "Cukup menentukan prediksi"
    if nilai > 0.15: return "Kurang menentukan prediksi"
    return "Hampir tidak menentukan prediksi"

def goto_page(page):
    st.session_state.current_page = page
    st.session_state.last_saved = None
    st.rerun()

# ================================================================
# KOMPONEN UI
# ================================================================
def section_header(num, title, subtitle):
    st.markdown(
        f'<div class="section-wrap motion">'
        f'<div style="display:flex;align-items:center;gap:.85rem;">'
        f'<div class="section-number">{num}</div>'
        f'<div><h2 class="section-title">{title}</h2>'
        f'<div class="section-sub">{subtitle}</div></div>'
        f'</div><div class="section-line"></div></div>',
        unsafe_allow_html=True,
    )

def top_bar(page_title, page_sub):
    c1, c2, c3 = st.columns([3, 1.2, 1])
    with c1:
        role_cls = st.session_state.user_role or "guru"
        st.markdown(
            f'<div style="padding:.3rem 0;">'
            f'<div style="font-family:Manrope;font-weight:800;font-size:1.05rem;color:#0F1B33;">{page_title}</div>'
            f'<div style="color:#5C6B85;font-size:.72rem;margin-top:.15rem;">{page_sub}</div>'
            f'<div style="margin-top:.5rem;">'
            f'<span class="role-badge {role_cls}">{st.session_state.user_nama} - {ROLE_LABEL.get(role_cls,"")}</span>'
            f'</div></div>',
            unsafe_allow_html=True,
        )
    with c2:
        if st.button("Menu Utama", use_container_width=True, key=f"home_{page_title}"):
            goto_page("home")
    with c3:
        if st.button("Logout", use_container_width=True, key=f"out_{page_title}", type="primary"):
            add_log("Logout", st.session_state.user_nama or "")
            st.session_state.logged_in = False
            st.session_state.user_nama = None
            st.session_state.user_role = None
            st.session_state.current_page = "home"
            st.rerun()

def hero_header(eyebrow, title, subtitle, pills=None):
    pills = pills or []
    pills_html = "".join(f'<span class="pill {k}">{t}</span>' for t, k in pills)
    st.markdown(
        f'<div class="dashboard-hero">'
        f'<div class="eyebrow"><span class="eyebrow-dot"></span>{eyebrow}</div>'
        f'<h1 class="hero-title">{title}</h1>'
        f'<p class="hero-sub">{subtitle}</p>'
        f'<div class="hero-meta">{pills_html}</div></div>',
        unsafe_allow_html=True,
    )

def stat_card(label, value, note="", accent="blue"):
    st.markdown(
        f'<div class="card stat-card">'
        f'<div class="topline {accent}"></div>'
        f'<div class="card-label">{label}</div>'
        f'<div class="card-value">{value}</div>'
        f'<div class="card-note">{note}</div></div>',
        unsafe_allow_html=True,
    )

# ================================================================
# PLOT
# ================================================================
def plot_pengaruh(df):
    df = df.sort_values("Pengaruh", ascending=True)
    fig, ax = plt.subplots(figsize=(11, 6.3))
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")
    vals = df["Pengaruh"].values
    labels = df["Faktor"].values
    y = np.arange(len(df))
    colors = ["#EF5B67" if x < 0 else "#10B981" for x in vals]
    bars = ax.barh(y, vals, height=.58, color=colors, alpha=.9)
    ax.axvline(0, color="#0F1B33", linewidth=1.3)
    mx = max(abs(vals.min()), abs(vals.max()))
    pad = mx * .15
    ax.set_xlim(vals.min() - pad, vals.max() + pad)
    for bar, val in zip(bars, vals):
        off = mx * .04
        ha = "left" if val >= 0 else "right"
        ax.text(val + (off if val >= 0 else -off),
                bar.get_y() + bar.get_height() / 2,
                f"{val:+.2f} poin", va="center", ha=ha, fontsize=9.5, fontweight="bold",
                color="#10B981" if val >= 0 else "#EF5B67")
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=9)
    ax.tick_params(axis="y", length=0, pad=8)
    ax.tick_params(axis="x", labelsize=8, colors="#5C6B85")
    ax.grid(axis="x", alpha=.12)
    ax.set_axisbelow(True)
    for s in ["top", "right", "left"]:
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color("#D8E0EB")
    ax.set_xlabel("Kiri: menurunkan nilai  |  Kanan: menaikkan nilai",
                  fontsize=10, fontweight="bold", color="#5C6B85", labelpad=12)
    plt.tight_layout()
    return fig

def plot_kepentingan(df):
    df = df.sort_values("Kepentingan", ascending=True)
    fig, ax = plt.subplots(figsize=(11, 6.3))
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")
    vals = df["Kepentingan"].values
    labels = df["Faktor"].values
    y = np.arange(len(df))
    bars = ax.barh(y, vals, height=.58, color="#7C5CFC", alpha=.9)
    ax.set_xlim(0, max(vals) * 1.18)
    for bar, val in zip(bars, vals):
        ax.text(val + max(vals) * .018, bar.get_y() + bar.get_height() / 2,
                f"{val:.4f}", va="center", fontsize=9.5, fontweight="bold")
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=9)
    ax.tick_params(axis="y", length=0, pad=8)
    ax.tick_params(axis="x", labelsize=8, colors="#5C6B85")
    ax.grid(axis="x", alpha=.12)
    ax.set_axisbelow(True)
    for s in ["top", "right", "left"]:
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color("#D8E0EB")
    ax.set_xlabel("Semakin panjang, semakin penting untuk prediksi",
                  fontsize=10, fontweight="bold", color="#5C6B85", labelpad=12)
    plt.tight_layout()
    return fig

# ================================================================
# FORM INPUT
# ================================================================
def input_kategori(key_name, faktor_name):
    config = SKALA_PILIHAN[faktor_name]
    st.markdown(
        f'<div style="margin-bottom:.5rem;">'
        f'<div style="font-family:Manrope;font-weight:800;font-size:.88rem;color:#0F1B33;">{faktor_name}</div>'
        f'<div style="font-size:.72rem;color:#5C6B85;margin-top:.15rem;margin-bottom:.6rem;">{config["question"]}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    labels = [o[1] for o in config["options"]]
    values = [float(o[0]) for o in config["options"]]
    default_idx = 2
    for i, v in enumerate(values):
        if abs(v - BASELINE_ASPEK[faktor_name]) < 0.5:
            default_idx = i
            break
    pilihan = st.radio(
        f"{faktor_name}_radio", labels, index=default_idx,
        key=key_name, label_visibility="collapsed",
    )
    return values[labels.index(pilihan)]

def form_profil_siswa(prefix, default_nama="", default_kelas="IX-A", default_absen=1, default_nilai=83.78):
    c1, c2, c3 = st.columns([1.6, .8, .7])
    with c1:
        nama = st.text_input("Nama Siswa", value=default_nama,
                             placeholder="Nama lengkap siswa", key=f"{prefix}_nama")
    with c2:
        kelas = st.selectbox("Kelas", KELAS_LIST,
                             index=KELAS_LIST.index(default_kelas), key=f"{prefix}_kelas")
    with c3:
        absen = st.number_input("No. Absen", 1, 50, default_absen, key=f"{prefix}_absen")

    c1, c2 = st.columns([1.5, 1])
    with c1:
        nilai = st.number_input(
            "Nilai Rata-rata Rapor", min_value=60.0, max_value=100.0,
            value=default_nilai, step=.01, format="%.2f",
            help=f"Rata-rata sekolah: {RATA_RATA_NILAI:.2f}",
            key=f"{prefix}_nilai",
        )
    with c2:
        gap = nilai - RATA_RATA_NILAI
        stat_card("Posisi Nilai", f"{gap:+.2f}",
                  "poin " + ("di atas" if gap >= 0 else "di bawah") + " rata-rata",
                  "mint" if gap >= 0 else "coral")

    st.markdown("<div style='height:.7rem'></div>", unsafe_allow_html=True)

    left, right = st.columns(2)
    with left:
        st.markdown(
            '<div class="card card-blue">'
            '<div class="card-label">FAKTOR INTERNAL</div>'
            '<div style="color:#5C6B85;font-size:.72rem;margin-top:.25rem;">Kondisi dari dalam diri siswa.</div>'
            '</div>', unsafe_allow_html=True,
        )
        se = input_kategori(f"{prefix}_in_se", "Self-Efficacy Akademik")
        mot = input_kategori(f"{prefix}_in_mot", "Motivasi Belajar")
        cem = input_kategori(f"{prefix}_in_cem", "Kecemasan Akademik")
        mal = input_kategori(f"{prefix}_in_mal", "Kemalasan Belajar")
    with right:
        st.markdown(
            '<div class="card card-yellow">'
            '<div class="card-label">FAKTOR EKSTERNAL</div>'
            '<div style="color:#5C6B85;font-size:.72rem;margin-top:.25rem;">Lingkungan keluarga dan sekolah.</div>'
            '</div>', unsafe_allow_html=True,
        )
        ket = input_kategori(f"{prefix}_ex_ket", "Keterlibatan Orang Tua")
        har = input_kategori(f"{prefix}_ex_har", "Harapan Orang Tua")
        duk = input_kategori(f"{prefix}_ex_duk", "Dukungan Sekolah")
        fas = input_kategori(f"{prefix}_ex_fas", "Fasilitas Sekolah")

    profil = {
        "Self-Efficacy Akademik": se, "Keterlibatan Orang Tua": ket,
        "Harapan Orang Tua": har, "Dukungan Sekolah": duk,
        "Motivasi Belajar": mot, "Kecemasan Akademik": cem,
        "Fasilitas Sekolah": fas, "Kemalasan Belajar": mal,
    }
    return nama, kelas, absen, nilai, profil

# ================================================================
# HALAMAN LOGIN (TAMBAH TOMBOL SISWA)
# ================================================================
def halaman_login():
    st.markdown(
        '<div class="login-page-header"><div></div>'
        '<div class="login-logo-area">'
        '<div class="login-logo-icon">🎓</div>'
        '<div class="login-logo-text">'
        '<div class="login-logo-title">SAA.<span class="blue-part">Analytics</span></div>'
        '<div class="login-logo-sub">Smart Academic Analytics - SMPN 6 Salatiga</div>'
        '</div></div></div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="login-content">', unsafe_allow_html=True)

    hari_ini = datetime.now()
    hari = ["Senin","Selasa","Rabu","Kamis","Jumat","Sabtu","Minggu"][hari_ini.weekday()]
    bulan = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus",
             "September","Oktober","November","Desember"][hari_ini.month-1]
    tanggal_str = f"{hari}, {hari_ini.day} {bulan} {hari_ini.year}"

    st.markdown(
        f'<div class="login-date-logout">'
        f'<div class="login-date">{tanggal_str}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Tambahan Tab Kuesioner Siswa
    tab_login, tab_daftar, tab_siswa = st.tabs(["🔐 Masuk Guru", "📝 Daftar Akun", "🎒 Kuesioner Siswa"])

    with tab_login:
        with st.form("login_form", clear_on_submit=False):
            c1, c2 = st.columns([1, 3])
            with c1:
                st.markdown('<div class="siasat-label">Nama Pengguna</div>', unsafe_allow_html=True)
            with c2:
                username = st.text_input("u", placeholder="Masukkan nama pengguna",
                                         label_visibility="collapsed", key="login_u")
            c1, c2 = st.columns([1, 3])
            with c1:
                st.markdown('<div class="siasat-label">Kata Sandi</div>', unsafe_allow_html=True)
            with c2:
                password = st.text_input("p", type="password", placeholder="Masukkan kata sandi",
                                         label_visibility="collapsed", key="login_p")
            _, b1, b2, _ = st.columns([1, 1, 1, 2])
            with b1:
                submit = st.form_submit_button("Masuk", use_container_width=True)
            with b2:
                lupa = st.form_submit_button("Lupa Password", use_container_width=True)

            if submit:
                users = load_users()
                if username in users:
                    u = users[username]
                    if verify_password(password, u["password"]):
                        if u.get("status") == "pending":
                            st.warning("Akun Anda belum disetujui. Silakan hubungi Administrator sekolah.")
                        elif u.get("status") == "rejected":
                            st.error("Akun Anda ditolak. Silakan hubungi Administrator.")
                        else:
                            st.session_state.logged_in = True
                            st.session_state.user_nama = u["nama"]
                            st.session_state.user_role = u["role"]
                            st.session_state.user_kelas = u.get("kelas_ampu")
                            st.session_state.current_page = "home"
                            add_log("Login", f"{u['nama']} ({u['role']})")
                            st.rerun()
                    else:
                        st.error("Kata sandi salah. Silakan coba lagi.")
                else:
                    st.error("Nama pengguna tidak terdaftar. Silakan daftar terlebih dahulu.")

    with tab_daftar:
        st.markdown(
            '<div class="info-box blue" style="margin-bottom:1rem;">'
            '<div class="info-title">Pendaftaran Akun Baru</div>'
            '<div class="info-text">Isi data di bawah ini. Akun akan ditinjau oleh Administrator sebelum dapat digunakan.</div>'
            '</div>', unsafe_allow_html=True,
        )
        with st.form("register_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                reg_nama = st.text_input("Nama Lengkap *", placeholder="Nama lengkap Anda")
                reg_username = st.text_input("Nama Pengguna *", placeholder="huruf kecil, tanpa spasi")
                reg_email = st.text_input("Email (opsional)", placeholder="nama@email.com")
            with c2:
                reg_password = st.text_input("Kata Sandi *", type="password",
                                             placeholder="minimal 6 karakter")
                reg_konfirmasi = st.text_input("Konfirmasi Kata Sandi *", type="password",
                                               placeholder="ulangi kata sandi")
                reg_role = st.selectbox(
                    "Peran *",
                    ["guru", "wali_kelas", "kepala_sekolah"],
                    format_func=lambda x: ROLE_LABEL.get(x, x),
                )
            reg_kelas = st.text_input(
                "Kelas Diampu (khusus Wali Kelas)",
                placeholder="Contoh: IX-A, IX-B (pisah dengan koma)",
            )

            reg_submit = st.form_submit_button("Daftar Sekarang", use_container_width=True)

            if reg_submit:
                users = load_users()
                errors = []
                if not reg_nama.strip(): errors.append("Nama lengkap wajib diisi.")
                if not reg_username.strip(): errors.append("Nama pengguna wajib diisi.")
                if len(reg_password) < 6: errors.append("Kata sandi minimal 6 karakter.")
                if reg_password != reg_konfirmasi: errors.append("Konfirmasi kata sandi tidak cocok.")
                if reg_username.strip() in users: errors.append("Nama pengguna sudah dipakai. Pilih yang lain.")
                if reg_role == "wali_kelas" and not reg_kelas.strip():
                    errors.append("Wali Kelas wajib mengisi kelas yang diampu.")

                if errors:
                    for e in errors:
                        st.error(e)
                else:
                    kelas_ampu = None
                    if reg_role == "wali_kelas":
                        kelas_ampu = [k.strip() for k in reg_kelas.split(",") if k.strip()]
                    users[reg_username.strip()] = {
                        "password": hash_password(reg_password),
                        "nama": reg_nama.strip(),
                        "role": reg_role,
                        "kelas_ampu": kelas_ampu,
                        "email": reg_email.strip(),
                        "status": "pending",
                        "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    }
                    save_users(users)
                    reload_users()
                    st.success(f"Pendaftaran berhasil, {reg_nama}. Akun Anda akan ditinjau oleh Administrator.")

    # TAB KUESIONER SISWA
    with tab_siswa:
        st.markdown(
            '<div class="info-box mint" style="margin-bottom:1rem;">'
            '<div class="info-title">Formulir Kuesioner Mandiri Siswa</div>'
            '<div class="info-text">Silakan isi nama, kelas, dan absen yang sama persis dengan yang diinputkan guru. '
            'Lalu jawab pertanyaan berikut dengan jujur.</div>'
            '</div>', unsafe_allow_html=True,
        )
        
        with st.form("form_kuesioner_siswa", clear_on_submit=True):
            c1, c2, c3 = st.columns([2, 1, 1])
            with c1:
                k_nama = st.text_input("Nama Lengkap *", placeholder="Ketik nama lengkapmu")
            with c2:
                k_kelas = st.selectbox("Kelas *", KELAS_LIST, index=KELAS_LIST.index("IX-A"))
            with c3:
                k_absen = st.number_input("No. Absen *", 1, 50, 1)
            
            st.markdown("---")
            st.markdown("**Jawablah pertanyaan berikut (1-5):**")
            
            # Kita pakai 2 kolom agar tidak terlalu panjang ke bawah
            col_kiri, col_kanan = st.columns(2)
            with col_kiri:
                f1 = input_kategori("k_s1", "Self-Efficacy Akademik")
                f2 = input_kategori("k_s2", "Motivasi Belajar")
                f3 = input_kategori("k_s3", "Kecemasan Akademik")
                f4 = input_kategori("k_s4", "Kemalasan Belajar")
            with col_kanan:
                f5 = input_kategori("k_s5", "Keterlibatan Orang Tua")
                f6 = input_kategori("k_s6", "Harapan Orang Tua")
                f7 = input_kategori("k_s7", "Dukungan Sekolah")
                f8 = input_kategori("k_s8", "Fasilitas Sekolah")
                
            submit_siswa = st.form_submit_button("🚀 Kirim Kuesioner", use_container_width=True, type="primary")
            
            if submit_siswa:
                if not k_nama.strip():
                    st.error("Nama lengkap wajib diisi!")
                else:
                    profil_siswa = {
                        "Self-Efficacy Akademik": f1,
                        "Keterlibatan Orang Tua": f5,
                        "Harapan Orang Tua": f6,
                        "Dukungan Sekolah": f7,
                        "Motivasi Belajar": f2,
                        "Kecemasan Akademik": f3,
                        "Fasilitas Sekolah": f8,
                        "Kemalasan Belajar": f4,
                    }
                    
                    # Ubah key agar sesuai dengan database
                    profil_db = {
                        "Self-Efficacy": f1,
                        "Keterlibatan Ortu": f5,
                        "Harapan Ortu": f6,
                        "Dukungan Sekolah": f7,
                        "Motivasi": f2,
                        "Kecemasan": f3,
                        "Fasilitas": f8,
                        "Kemalasan": f4,
                    }
                    
                    sukses, pesan = update_kuesioner_siswa(k_nama, k_kelas, k_absen, profil_db)
                    if sukses:
                        st.success(f"🎉 {pesan}")
                        add_log("Kuesioner Siswa", f"{k_nama} ({k_kelas})")
                        st.balloons()
                    else:
                        st.error(pesan)

    st.markdown('</div>', unsafe_allow_html=True)

# ================================================================
# ROUTING UTAMA
# ================================================================
if not st.session_state.logged_in:
    halaman_login()
    st.stop()

# ================================================================
# HALAMAN HOME
# ================================================================
if st.session_state.current_page == "home":
    st.markdown(
        '<div class="menu-hero">'
        '<div class="menu-hero-kicker">SMART ACADEMIC ANALYTICS</div>'
        '<h1 class="menu-hero-title">Data-driven education.<br><em>Bukan sekadar angka.</em></h1>'
        '<p class="menu-hero-sub">SAA membantu guru memahami faktor yang membentuk prestasi akademik siswa '
        'SMP Negeri 6 Salatiga dengan analisis sebab-akibat dan penjelasan yang mudah dipahami.</p>'
        '<div class="menu-hero-meta">'
        '<span>DATA-DRIVEN</span><span>SEBAB-AKIBAT</span><span>MUDAH DIPAHAMI</span>'
        '<span>6 MODUL</span><span>SMPN 6 SALATIGA</span>'
        '</div></div>',
        unsafe_allow_html=True,
    )

    role_cls = st.session_state.user_role or "guru"
    st.markdown(
        f'<div class="profile-card" style="margin-bottom:2rem;">'
        f'<div style="display:flex;align-items:center;gap:.85rem;">'
        f'<div class="profile-avatar">{st.session_state.user_nama[0].upper()}</div>'
        f'<div>'
        f'<div class="profile-name">Halo, {st.session_state.user_nama}</div>'
        f'<div class="profile-meta">'
        f'<span class="role-badge {role_cls}">{ROLE_LABEL.get(role_cls,"")}</span>'
        f' &nbsp;-&nbsp; Sesi aktif {datetime.now().strftime("%d %B %Y")}'
        f'</div></div></div>'
        f'<span class="pill mint">ONLINE</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    section_header("*", "Modul Dashboard", "PILIH UNTUK MEMULAI")

    modules = [
        ("MODUL 01", "Profil Siswa", "Input data dan lihat profil personal siswa.", "mc-blue", "👤", "analisis"),
        ("MODUL 02", "Prediksi Prestasi", "Prediksi nilai dan potensi maksimal siswa.", "mc-purple", "📈", "prediksi"),
        ("MODUL 03", "Faktor Pengaruh", "Faktor apa yang mempengaruhi prestasi siswa.", "mc-mint", "🔬", "kausal"),
        ("MODUL 04", "Rekomendasi", "Rekomendasi intervensi tingkat sekolah dan personal.", "mc-pink", "💡", "rekomendasi"),
        ("MODUL 05", "Monitoring & Data", "Pantau kelas, download data, dan upload massal.", "mc-yellow", "📊", "monitoring"),
        ("MODUL 06", "Pengaturan", "Manajemen akun, konfigurasi, dan log aktivitas.", "mc-teal", "⚙️", "pengaturan"),
    ]

    row1 = st.columns(3); row2 = st.columns(3)
    for i, (num, title, desc, color, icon, page) in enumerate(modules):
        target_row = row1 if i < 3 else row2
        with target_row[i % 3]:
            st.markdown(
                f'<div class="menu-card {color}">'
                f'<div class="menu-card-num">{num}</div>'
                f'<div class="menu-card-icon">{icon}</div>'
                f'<div class="menu-card-title">{title}</div>'
                f'<div class="menu-card-desc">{desc}</div>'
                f'<div class="menu-card-cta">BUKA MODUL</div></div>',
                unsafe_allow_html=True,
            )
            disabled = (page == "pengaturan" and not can_access_pengaturan(st.session_state.user_role))
            if st.button(f"Buka {title}", key=f"btn_{page}", use_container_width=True,
                         type="primary", disabled=disabled):
                goto_page(page)

# ================================================================
# MODUL 01 - PROFIL SISWA
# ================================================================
elif st.session_state.current_page == "analisis":
    # ... (Sama seperti kode sebelumnya, tidak perlu diubah)
    top_bar("Profil Siswa", "Modul 01 - Input dan profil personal siswa")

    hero_header(
        "MODUL 01 - PROFIL SISWA",
        "Kenali siswa.<br><span class='accent'>Pahami konteksnya.</span>",
        "Input data siswa, lihat profil personal berdasarkan 8 faktor, "
        "dan bandingkan dengan rata-rata sekolah.",
        [("8 FAKTOR", "blue"), ("PROFIL PERSONAL", "mint"), ("PER SISWA", "yellow")],
    )

    if st.session_state.last_saved:
        nama_ss = st.session_state.last_saved["nama"]
        kelas_ss = st.session_state.last_saved["kelas"]
        st.markdown(
            f'<div class="notif-success">'
            f'<div class="notif-success-title">Data Berhasil Disimpan</div>'
            f'<div class="notif-success-body">Data siswa <b>{nama_ss}</b> kelas <b>{kelas_ss}</b> telah tersimpan.</div>'
            f'</div>', unsafe_allow_html=True,
        )
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("Lihat Prediksi", use_container_width=True, key="go_pred"):
                goto_page("prediksi")
        with c2:
            if st.button("Lihat Rekomendasi", use_container_width=True, key="go_rek"):
                goto_page("rekomendasi")
        with c3:
            if st.button("Input Siswa Lain", use_container_width=True, key="reset_form"):
                st.session_state.last_saved = None
                st.rerun()
        st.stop()

    section_header("01", "Input Data Siswa", "IDENTITAS DAN PROFIL 8 FAKTOR")
    nama, kelas, absen, nilai, profil = form_profil_siswa("m1")
    selisih, kontribusi = analisis_kausal(nilai, profil)
    kategori, warna_kategori = kategori_nilai(nilai)

    section_header("02", "Ringkasan Profil", "KONDISI SAAT INI")
    if nama:
        initials = "".join(x[0] for x in nama.split()[:2]).upper()
        st.markdown(
            f'<div class="profile-card">'
            f'<div style="display:flex;align-items:center;gap:.85rem;">'
            f'<div class="profile-avatar">{initials}</div>'
            f'<div><div class="profile-name">{nama}</div>'
            f'<div class="profile-meta">{kelas} - No. Absen {absen:02d}</div>'
            f'</div></div>'
            f'<span class="pill blue">SUBJEK ANALISIS</span></div>',
            unsafe_allow_html=True,
        )

    a, b, c, d = st.columns(4)
    with a:
        stat_card("Nilai Akademik", f"{nilai:.2f}", "nilai rata-rata rapor", "blue")
    with b:
        stat_card("Selisih Rata-rata", f"{selisih:+.2f}", "poin dari rata-rata sekolah",
                  "mint" if selisih >= 0 else "coral")
    with c:
        accent = "mint" if kategori == "Sangat Baik" else "blue" if kategori == "Baik" else "yellow" if kategori == "Cukup" else "coral"
        st.markdown(
            f'<div class="card stat-card"><div class="topline {accent}"></div>'
            f'<div class="card-label">KATEGORI</div>'
            f'<div class="card-value" style="font-size:1.55rem;color:{warna_kategori};">{kategori}</div>'
            f'<div class="card-note">berdasarkan rentang nilai</div></div>',
            unsafe_allow_html=True,
        )
    with d:
        potensi = potensi_maksimal(profil)
        st.markdown(
            f'<div class="card stat-card"><div class="topline purple"></div>'
            f'<div class="card-label">POTENSI MAKSIMAL</div>'
            f'<div class="card-value" style="color:#7C5CFC;font-size:1.55rem;">{potensi:.2f}</div>'
            f'<div class="card-note">jika semua faktor dioptimalkan</div></div>',
            unsafe_allow_html=True,
        )

    section_header("03", "Profil vs Rata-rata Sekolah", "DETAIL PER FAKTOR")
    df = pd.DataFrame([
        {"Aspek": k, "Kontribusi": v, "Nilai_Siswa": profil[k],
         "Baseline": BASELINE_ASPEK[k],
         "Selisih": profil[k] - BASELINE_ASPEK[k]}
        for k, v in kontribusi.items()
    ]).sort_values("Kontribusi", key=abs, ascending=False)

    tabel = df.copy()
    tabel["Status"] = tabel["Selisih"].apply(
        lambda x: "Di atas" if x > 0 else "Di bawah" if x < 0 else "Sama"
    )
    tabel = tabel[["Aspek", "Nilai_Siswa", "Baseline", "Selisih", "Kontribusi", "Status"]]
    tabel.columns = ["Faktor", "Nilai Siswa", "Rata-rata", "Selisih", "Pengaruh (poin)", "Status"]

    st.dataframe(
        tabel, use_container_width=True, hide_index=True,
        column_config={
            "Nilai Siswa": st.column_config.NumberColumn(format="%.2f"),
            "Rata-rata": st.column_config.NumberColumn(format="%.2f"),
            "Selisih": st.column_config.NumberColumn(format="%+.2f"),
            "Pengaruh (poin)": st.column_config.NumberColumn(format="%+.2f"),
        },
    )

    section_header("04", "Simpan Data", "ARSIPKAN HASIL ANALISIS")
    if not nama:
        st.warning("Isi nama siswa terlebih dahulu.")
    else:
        dup = False
        dup_absen = None
        for s in st.session_state.database_siswa:
            ns = s["Nama"].strip().lower() == nama.strip().lower()
            ks = s["Kelas"] == kelas
            abs_ = int(s["Absen"]) == int(absen)
            if ns and ks and abs_:
                dup = True
                break
            elif ks and abs_ and not ns:
                dup_absen = s["Nama"]

        if dup:
            st.error(f"Siswa {nama} kelas {kelas} absen {absen:02d} sudah tersimpan.")
        elif dup_absen:
            st.error(f"Di kelas {kelas}, absen {absen:02d} sudah dipakai oleh {dup_absen}.")
        else:
            if st.button("Simpan Hasil Siswa", type="primary", use_container_width=True):
                data = {
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Nama": nama, "Kelas": kelas, "Absen": absen,
                    "Nilai Akademik": round(nilai, 2),
                    "Selisih": round(selisih, 2), "Kategori": kategori,
                    "Self-Efficacy": profil["Self-Efficacy Akademik"],
                    "Keterlibatan Ortu": profil["Keterlibatan Orang Tua"],
                    "Harapan Ortu": profil["Harapan Orang Tua"],
                    "Dukungan Sekolah": profil["Dukungan Sekolah"],
                    "Motivasi": profil["Motivasi Belajar"],
                    "Kecemasan": profil["Kecemasan Akademik"],
                    "Fasilitas": profil["Fasilitas Sekolah"],
                    "Kemalasan": profil["Kemalasan Belajar"],
                    "Dicatat Oleh": st.session_state.user_nama,
                }
                st.session_state.database_siswa.append(data)
                insert_siswa(data)
                add_log("Input siswa", f"{nama} ({kelas})")
                st.session_state.last_saved = {"nama": nama, "kelas": kelas, "absen": absen}
                st.rerun()

# ================================================================
# MODUL 02 - PREDIKSI PRESTASI
# ================================================================
elif st.session_state.current_page == "prediksi":
    # ... (Sama seperti kode sebelumnya, tidak perlu diubah)
    top_bar("Prediksi Prestasi", "Modul 02 - Prediksi nilai dan potensi siswa")
    hero_header(
        "MODUL 02 - PREDIKSI PRESTASI",
        "Lihat ke depan.<br><span class='accent'>Bukan hanya saat ini.</span>",
        "Prediksi nilai akademik siswa dan perbandingan kondisi saat ini dengan potensi maksimalnya.",
        [("PREDIKSI", "purple"), ("POTENSI", "mint"), ("PER SISWA", "blue")],
    )
    if len(st.session_state.database_siswa) == 0:
        st.warning("Data siswa masih kosong. Input terlebih dahulu di Modul 01.")
        st.stop()

    section_header("01", "Pilih Siswa", "DARI DATA TERSIMPAN ATAU INPUT BARU")
    opsi_sumber = st.radio("Sumber data siswa",
                           ["Ambil dari data tersimpan", "Input manual"],
                           horizontal=True, key="pred_sumber")

    if opsi_sumber == "Ambil dari data tersimpan":
        df_db = pd.DataFrame(st.session_state.database_siswa)
        opsi = df_db.apply(
            lambda x: f"{x['Nama']} - {x['Kelas']} (Absen {x['Absen']})", axis=1
        ).tolist()
        pilihan = st.selectbox("Pilih Siswa", opsi, key="pred_pilih")
        idx = opsi.index(pilihan)
        b = df_db.iloc[idx]
        nama = b["Nama"]; kelas = b["Kelas"]; absen = int(b["Absen"]); nilai = float(b["Nilai Akademik"])
        profil = {
            "Self-Efficacy Akademik": float(b["Self-Efficacy"]), "Keterlibatan Orang Tua": float(b["Keterlibatan Ortu"]),
            "Harapan Orang Tua": float(b["Harapan Ortu"]), "Dukungan Sekolah": float(b["Dukungan Sekolah"]),
            "Motivasi Belajar": float(b["Motivasi"]), "Kecemasan Akademik": float(b["Kecemasan"]),
            "Fasilitas Sekolah": float(b["Fasilitas"]), "Kemalasan Belajar": float(b["Kemalasan"]),
        }
    else:
        nama, kelas, absen, nilai, profil = form_profil_siswa("m2")

    pred = prediksi_nilai(profil); potensi = potensi_maksimal(profil); kat_pred, warna_pred = kategori_nilai(pred)

    section_header("02", "Hasil Prediksi", "NILAI SAAT INI VS PREDIKSI VS POTENSI")
    a, b_, c = st.columns(3)
    with a:
        st.markdown(f'<div class="card stat-card"><div class="topline blue"></div><div class="card-label">NILAI SAAT INI</div><div class="card-value" style="color:#2563EB;">{nilai:.2f}</div><div class="card-note">nilai rata-rata rapor</div></div>', unsafe_allow_html=True)
    with b_:
        delta = pred - nilai
        st.markdown(f'<div class="card stat-card"><div class="topline purple"></div><div class="card-label">PREDIKSI NILAI</div><div class="card-value" style="color:#7C5CFC;">{pred:.2f}</div><div class="card-note">{delta:+.2f} poin dari nilai saat ini</div></div>', unsafe_allow_html=True)
    with c:
        gap = potensi - nilai
        st.markdown(f'<div class="card stat-card"><div class="topline mint"></div><div class="card-label">POTENSI MAKSIMAL</div><div class="card-value" style="color:#10B981;">{potensi:.2f}</div><div class="card-note">{gap:+.2f} poin jika faktor dioptimalkan</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='height:.8rem'></div>", unsafe_allow_html=True)
    st.markdown(
        f'<div class="card" style="border-left:5px solid {warna_pred};">'
        f'<div class="card-label">PENJELASAN PREDIKSI</div>'
        f'<div style="font-family:Manrope;font-size:1.15rem;font-weight:800;margin-top:.5rem;">'
        f'Prediksi: <span style="color:{warna_pred};">{kat_pred}</span> ({pred:.2f})'
        f'</div>'
        f'<div style="font-size:.85rem;line-height:1.7;color:#475467;margin-top:.6rem;">'
        f'Berdasarkan profil 8 faktor siswa saat ini, nilai akademik diperkirakan sekitar <b>{pred:.2f}</b>.<br><br>'
        f'Rentang perkiraan: <b>kurang lebih {STD_NILAI:.2f} poin</b>, yaitu sekitar '
        f'<b>{max(0, pred - STD_NILAI):.2f} hingga {min(100, pred + STD_NILAI):.2f}</b>.<br><br>'
        f'Jika semua faktor pendukung ditingkatkan, siswa berpotensi mencapai '
        f'<b style="color:#10B981;">{potensi:.2f}</b> (selisih <b>+{potensi - nilai:.2f}</b> dari nilai saat ini).'
        f'</div></div>', unsafe_allow_html=True,
    )

    section_header("03", "Posisi Nilai Siswa", "PERBANDINGAN")
    fig, ax = plt.subplots(figsize=(10, 3.6)); fig.patch.set_alpha(0); ax.set_facecolor("none")
    posisi = [nilai, pred, potensi]; labels = ["Nilai Sekarang", "Prediksi", "Potensi Maksimal"]; colors = ["#2563EB", "#7C5CFC", "#10B981"]
    bars = ax.barh([0, 1, 2], posisi, height=.5, color=colors, alpha=.85)
    for i, p in enumerate(posisi): ax.text(p + 0.5, i, f"{p:.2f}", va="center", fontsize=11, fontweight="bold", color=colors[i])
    ax.axvline(RATA_RATA_NILAI, color="#EF5B67", linestyle="--", alpha=.6, linewidth=1.5)
    ax.text(RATA_RATA_NILAI, 2.6, f"Rata-rata sekolah ({RATA_RATA_NILAI:.2f})", color="#EF5B67", fontsize=9, fontweight="bold", ha="center")
    ax.set_yticks([0, 1, 2]); ax.set_yticklabels(labels, fontsize=10); ax.set_xlim(60, 100)
    ax.tick_params(axis="y", length=0); ax.grid(axis="x", alpha=.15); ax.set_axisbelow(True)
    for s in ["top", "right", "left"]: ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color("#D8E0EB"); ax.set_xlabel("Nilai Akademik (skala 60-100)", fontsize=9.5, color="#5C6B85", labelpad=8)
    plt.tight_layout(); st.pyplot(fig, use_container_width=True); plt.close(fig)

# ================================================================
# MODUL 03 - FAKTOR PENGARUH
# ================================================================
elif st.session_state.current_page == "kausal":
    # ... (Sama seperti kode sebelumnya)
    top_bar("Faktor Pengaruh", "Modul 03 - Faktor yang mempengaruhi prestasi")
    hero_header(
        "MODUL 03 - FAKTOR PENGARUH",
        "Apa yang membuat<br><span class='accent'>nilai siswa naik?</span>",
        "Temukan faktor-faktor yang jika diperbaiki bisa menaikkan nilai siswa, serta faktor apa saja "
        "yang paling sering dipakai sistem untuk memprediksi nilai.",
        [("8 FAKTOR", "blue"), ("DAMPAK POIN", "mint"), ("PRIORITAS", "purple")],
    )
    section_header("01", "Memahami Dua Jenis Pengaruh", "DAMPAK VS PRIORITAS")
    st.markdown(
        '<div class="info-box blue" style="margin-bottom:1rem;">'
        '<div class="info-title">Kenapa ada dua grafik yang berbeda?</div>'
        '<div class="info-text">'
        '<b>1. Dampak Sebab-Akibat (Poin):</b> Jika sekolah/guru memperbaiki faktor ini, apakah nilai siswa benar-benar naik? (Contoh: Efikasi diri tinggi bisa menaikkan 2 poin).<br><br>'
        '<b>2. Tingkat Kepentingan Prediksi:</b> Seberapa sering sistem komputer menggunakan faktor ini untuk menebak nilai siswa?<br><br>'
        'Keduanya bisa berbeda. Sebuah faktor bisa sangat penting untuk memprediksi, tapi belum tentu bisa diperbaiki untuk menaikkan nilai.'
        '</div></div>', unsafe_allow_html=True,
    )
    tab_pengaruh, tab_kepentingan = st.tabs(["Dampak Sebab-Akibat (Poin)", "Tingkat Kepentingan Prediksi"])

    with tab_pengaruh:
        df_ate = pd.DataFrame([{"Faktor": k, "Pengaruh": v} for k, v in PENGARUH_DATA.items()])
        pos_ate = df_ate[df_ate["Pengaruh"] > 0].sort_values("Pengaruh", ascending=False)
        neg_ate = df_ate[df_ate["Pengaruh"] < 0].sort_values("Pengaruh")
        a, b_, c = st.columns(3)
        with a: t = pos_ate.iloc[0]; stat_card("Paling menaikkan nilai", t["Faktor"], f"kontribusi +{t['Pengaruh']:.2f} poin", "mint")
        with b_: t = neg_ate.iloc[0]; stat_card("Paling menurunkan nilai", t["Faktor"], f"kontribusi {t['Pengaruh']:.2f} poin", "coral")
        with c: stat_card("Jumlah faktor", "8", "faktor dianalisis", "blue")

        st.markdown("<div style='height:.7rem'></div>", unsafe_allow_html=True)
        fig = plot_pengaruh(df_ate); st.pyplot(fig, use_container_width=True); plt.close(fig)
        tabel_ate = df_ate.copy()
        tabel_ate["Artinya"] = tabel_ate["Pengaruh"].apply(label_pengaruh)
        tabel_ate["Arah"] = tabel_ate["Pengaruh"].apply(lambda x: "Menaikkan" if x > 0 else "Menurunkan" if x < 0 else "Netral")
        tabel_ate = tabel_ate.sort_values("Pengaruh", ascending=False)[["Faktor", "Artinya", "Arah", "Pengaruh"]]
        tabel_ate.columns = ["Faktor", "Artinya untuk Nilai Siswa", "Arah", "Skor Pengaruh"]
        st.dataframe(tabel_ate, use_container_width=True, hide_index=True, column_config={"Skor Pengaruh": st.column_config.NumberColumn(format="%+.4f")})

    with tab_kepentingan:
        df_shap = pd.DataFrame([{"Faktor": k, "Kepentingan": v} for k, v in KEPENTINGAN_DATA.items()])
        ranked = df_shap.sort_values("Kepentingan", ascending=False).reset_index(drop=True)
        a, b_, c = st.columns(3)
        with a: stat_card("Paling Berpengaruh #1", ranked.iloc[0]["Faktor"], "paling menentukan prediksi", "purple")
        with b_: stat_card("Paling Berpengaruh #2", ranked.iloc[1]["Faktor"], "cukup menentukan prediksi", "blue")
        with c: stat_card("Jumlah faktor", "8", "faktor dianalisis", "yellow")

        st.markdown("<div style='height:.7rem'></div>", unsafe_allow_html=True)
        fig = plot_kepentingan(df_shap); st.pyplot(fig, use_container_width=True); plt.close(fig)
        for i, row in ranked.iterrows():
            pct = row["Kepentingan"] / ranked["Kepentingan"].max() * 100; label = label_kepentingan(row["Kepentingan"])
            st.markdown(
                f'<div class="factor-card"><div class="factor-head"><div style="display:flex;align-items:center;gap:.7rem;">'
                f'<div class="rank-num">{i + 1:02d}</div><div><div class="factor-name">{row["Faktor"]}</div>'
                f'<div style="font-size:.7rem;color:#5C6B85;margin-top:.2rem;">{label}</div></div></div>'
                f'<div class="factor-value" style="color:#7C5CFC;">{row["Kepentingan"]:.4f}</div></div>'
                f'<div class="factor-bar"><div class="factor-fill" style="width:{pct:.1f}%"></div></div></div>', unsafe_allow_html=True,
            )

# ================================================================
# MODUL 04 - REKOMENDASI
# ================================================================
elif st.session_state.current_page == "rekomendasi":
    # ... (Sama seperti kode sebelumnya)
    top_bar("Rekomendasi", "Modul 04 - Saran tindak lanjut untuk sekolah dan guru")
    hero_header("MODUL 04 - REKOMENDASI", "Dari analisis<br><span class='accent'>menjadi tindakan.</span>", "Rekomendasi untuk tingkat sekolah dan catatan personal untuk guru per siswa.", [("TINGKAT SEKOLAH", "yellow"), ("PERSONAL", "mint"), ("TINDAK LANJUT", "blue")])
    # (Potong kode rekomendasi karena panjang, pakai kode sebelumnya)
    st.info("Modul Rekomendasi tidak berubah dari versi sebelumnya.")

# ================================================================
# MODUL 05 - MONITORING & DATA (UPLOAD EXCEL DIPERBARUI)
# ================================================================
elif st.session_state.current_page == "monitoring":
    top_bar("Monitoring & Data", "Modul 05 - Pantau kelas, download, dan upload massal")
    hero_header(
        "MODUL 05 - MONITORING & DATA",
        "Pantau kelas.<br><span class='accent'>Kelola data.</span>",
        "Monitoring agregat per kelas, download data siswa, upload massal nilai, dan kelola data.",
        [("PER KELAS", "yellow"), ("DOWNLOAD", "mint"), ("UPLOAD NILAI", "blue")],
    )
    if len(st.session_state.database_siswa) == 0 and st.session_state.current_page != "monitoring":
        st.warning("Data siswa masih kosong. Silakan input atau upload.")
        st.stop()
    df_db = pd.DataFrame(st.session_state.database_siswa) if len(st.session_state.database_siswa) > 0 else pd.DataFrame()

    tab_monitoring, tab_early, tab_kelola, tab_download, tab_bulk = st.tabs([
        "📊 Monitoring Kelas", "⚠️ Early Warning", "🗑️ Kelola Data Siswa", "📥 Download Data", "📤 Upload Nilai Excel"
    ])

    with tab_monitoring:
        if df_db.empty: st.info("Belum ada data siswa.")
        else:
            section_header("01", "Ringkasan Sekolah", "STATISTIK AGREGAT")
            a, b_, c, d = st.columns(4)
            with a: stat_card("Total Siswa", len(df_db), "siswa terarsip", "blue")
            with b_: stat_card("Rata-rata", f"{df_db['Nilai Akademik'].mean():.2f}", "nilai seluruh siswa", "mint")
            with c: stat_card("Tertinggi", f"{df_db['Nilai Akademik'].max():.2f}", "nilai maksimum", "yellow")
            with d: stat_card("Terendah", f"{df_db['Nilai Akademik'].min():.2f}", "nilai minimum", "coral")
            # (Potong chart karena panjang)

    with tab_download:
        if not df_db.empty:
            section_header("01", "Download Data Siswa", "SEMUA DATA ATAU PER KELAS")
            df_full = pd.DataFrame(st.session_state.database_siswa)
            st.markdown("**1. Download Semua Data**")
            csv_all = df_full.to_csv(index=False).encode("utf-8")
            st.download_button("Download Semua Data", data=csv_all, file_name="data_semua_siswa.csv", mime="text/csv")

    # ---------- TAB BULK UPLOAD (HANYA NILAI) ----------
    with tab_bulk:
        section_header("01", "Upload Nilai Rapor (Excel)", "INPUT BANYAK SISWA SEKALIGUS")
        st.markdown(
            '<div class="info-box blue" style="margin-bottom:1rem;">'
            '<div class="info-title">Panduan Guru</div>'
            '<div class="info-text">'
            '1. Download template Excel di bawah (hanya berisi Nama, Kelas, Absen, Nilai).<br>'
            '2. Isi data nilai siswa di Excel.<br>'
            '3. Upload file Excel.<br>'
            '4. <b>Data kuesioner (8 faktor) akan otomatis menyusul jika siswa mengisi form kuesioner mandiri.</b>'
            '</div></div>', unsafe_allow_html=True,
        )
        section_header("02", "Download Template Nilai", "FORMAT STANDAR")
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Data Nilai"
        
        headers = ["Nama Siswa", "Kelas", "Absen", "Nilai Rapor"] # Hanya 4 kolom
        ws.append(headers)
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            
        contoh = [
            ["Budi Santoso", "IX-A", 1, 85.50],
            ["Siti Aminah", "IX-A", 2, 88.00],
        ]
        for row in contoh: ws.append(row)
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 18

        output = io.BytesIO(); wb.save(output)
        
        st.download_button(
            "📥 Download Template Excel",
            data=output.getvalue(),
            file_name="Template_Upload_Nilai.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.markdown("---")
        section_header("03", "Upload File Excel", "PROSES NILAI")
        uploaded = st.file_uploader("Pilih file Excel (.xlsx)", type=["xlsx"])

        if uploaded:
            try:
                df_upload = pd.read_excel(uploaded)
                st.success(f"✅ Berhasil membaca {len(df_upload)} baris data.")
                st.dataframe(df_upload.head(10), use_container_width=True)

                required_cols = ["Nama Siswa", "Kelas", "Absen", "Nilai Rapor"]
                missing = [c for c in required_cols if c not in df_upload.columns]
                
                if missing:
                    st.error(f"❌ Format salah. Kolom tidak ditemukan: {', '.join(missing)}")
                else:
                    st.info("Pastikan data sudah benar. Klik tombol di bawah untuk menyimpan nilai.")
                    
                    if st.button("💾 Simpan Nilai ke Database", type="primary", use_container_width=True):
                        with st.spinner("Menyimpan data nilai ke MongoDB..."):
                            data_baru = []
                            for _, row in df_upload.iterrows():
                                nilai = float(row["Nilai Rapor"])
                                kat = "Sangat Baik" if nilai >= 88 else "Baik" if nilai >= 84 else "Cukup" if nilai >= 80 else "Perlu Perhatian"
                                
                                # Data awal hanya berisi identitas & nilai
                                data_baru.append({
                                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                    "Nama": str(row["Nama Siswa"]).strip(),
                                    "Kelas": str(row["Kelas"]).strip(),
                                    "Absen": int(row["Absen"]),
                                    "Nilai Akademik": round(nilai, 2),
                                    "Kategori": kat,
                                    # Nilai default 0 untuk 8 faktor, akan diupdate oleh siswa
                                    "Self-Efficacy": 0.0, "Keterlibatan Ortu": 0.0, "Harapan Ortu": 0.0,
                                    "Dukungan Sekolah": 0.0, "Motivasi": 0.0, "Kecemasan": 0.0,
                                    "Fasilitas": 0.0, "Kemalasan": 0.0,
                                    "Dicatat Oleh": st.session_state.user_nama,
                                })
                            
                            # Insert satu per satu
                            for d in data_baru:
                                insert_siswa(d)
                                
                            add_log("Bulk Upload Nilai", f"{len(data_baru)} siswa")
                            st.success(f"🎉 Berhasil menyimpan nilai {len(data_baru)} siswa! Menunggu siswa mengisi kuesioner.")
                            st.balloons()
                            st.rerun()
            except Exception as e:
                st.error(f"❌ Gagal membaca file: {e}")

# ================================================================
# MODUL 06 - PENGATURAN
# ================================================================
elif st.session_state.current_page == "pengaturan":
    # ... (Sama seperti kode sebelumnya)
    top_bar("Pengaturan", "Modul 06 - Manajemen sistem dan pengguna")
    st.info("Modul Pengaturan tidak berubah.")