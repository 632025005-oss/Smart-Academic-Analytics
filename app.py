# -*- coding: utf-8 -*-
# ================================================================
# SMART ACADEMIC ANALYTICS (SAA) - SMP Negeri 6 Salatiga
# Sub-brand: SIA.Prestasi - Magister Sains Data UKSW
# MongoDB Edition + Bulk Upload Excel + Bahasa Awam
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
    insert_siswa, hapus_semua_siswa
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
    "Self-Efficacy Akademik": {"question": "Seberapa yakin siswa terhadap kemampuan akademiknya?", "options": [
        ("1","Hampir tidak pernah percaya diri"),("2","Jarang percaya diri"),
        ("3","Kadang-kadang percaya diri"),("4","Sering percaya diri"),("5","Hampir selalu percaya diri")]},
    "Keterlibatan Orang Tua": {"question": "Seberapa aktif orang tua mendampingi siswa belajar?", "options": [
        ("1","Hampir tidak pernah mendampingi"),("2","Jarang mendampingi"),
        ("3","Kadang-kadang mendampingi"),("4","Sering mendampingi"),("5","Hampir selalu mendampingi")]},
    "Harapan Orang Tua": {"question": "Seberapa tinggi tuntutan orang tua terhadap nilai siswa?", "options": [
        ("1","Sangat rendah"),("2","Rendah"),("3","Sedang / wajar"),
        ("4","Tinggi"),("5","Sangat tinggi / menekan")]},
    "Dukungan Sekolah": {"question": "Seberapa besar dukungan sekolah kepada siswa?", "options": [
        ("1","Sangat kurang"),("2","Kurang"),("3","Cukup"),("4","Baik"),("5","Sangat baik")]},
    "Motivasi Belajar": {"question": "Seberapa besar motivasi dan semangat belajar siswa?", "options": [
        ("1","Tidak ada motivasi"),("2","Kurang termotivasi"),("3","Cukup termotivasi"),
        ("4","Termotivasi"),("5","Sangat termotivasi")]},
    "Kecemasan Akademik": {"question": "Seberapa sering siswa merasa cemas saat belajar atau ujian?", "options": [
        ("1","Hampir tidak pernah cemas"),("2","Jarang cemas"),("3","Kadang-kadang cemas"),
        ("4","Sering cemas"),("5","Hampir selalu cemas")]},
    "Fasilitas Sekolah": {"question": "Seberapa memadai fasilitas belajar di sekolah?", "options": [
        ("1","Sangat kurang memadai"),("2","Kurang memadai"),("3","Cukup memadai"),
        ("4","Memadai"),("5","Sangat memadai")]},
    "Kemalasan Belajar": {"question": "Seberapa sering siswa menunjukkan sikap malas belajar?", "options": [
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
# HALAMAN LOGIN
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

    tab_login, tab_daftar = st.tabs(["🔐 Masuk", "📝 Daftar Akun Baru"])

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

            if lupa:
                st.info("Hubungi Administrator sekolah untuk reset password.")

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
                    st.success(
                        f"Pendaftaran berhasil, {reg_nama}. Akun Anda akan ditinjau oleh Administrator."
                    )

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
                insert_siswa(data) # Menggunakan fungsi insert MongoDB
                add_log("Input siswa", f"{nama} ({kelas})")
                st.session_state.last_saved = {"nama": nama, "kelas": kelas, "absen": absen}
                st.rerun()

# ================================================================
# MODUL 02 - PREDIKSI PRESTASI
# ================================================================
elif st.session_state.current_page == "prediksi":
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
        nama = b["Nama"]
        kelas = b["Kelas"]
        absen = int(b["Absen"])
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
    else:
        nama, kelas, absen, nilai, profil = form_profil_siswa("m2")

    pred = prediksi_nilai(profil)
    potensi = potensi_maksimal(profil)
    kat_pred, warna_pred = kategori_nilai(pred)

    section_header("02", "Hasil Prediksi", "NILAI SAAT INI VS PREDIKSI VS POTENSI")

    a, b_, c = st.columns(3)
    with a:
        st.markdown(
            f'<div class="card stat-card"><div class="topline blue"></div>'
            f'<div class="card-label">NILAI SAAT INI</div>'
            f'<div class="card-value" style="color:#2563EB;">{nilai:.2f}</div>'
            f'<div class="card-note">nilai rata-rata rapor</div></div>',
            unsafe_allow_html=True,
        )
    with b_:
        delta = pred - nilai
        st.markdown(
            f'<div class="card stat-card"><div class="topline purple"></div>'
            f'<div class="card-label">PREDIKSI NILAI</div>'
            f'<div class="card-value" style="color:#7C5CFC;">{pred:.2f}</div>'
            f'<div class="card-note">{delta:+.2f} poin dari nilai saat ini</div></div>',
            unsafe_allow_html=True,
        )
    with c:
        gap = potensi - nilai
        st.markdown(
            f'<div class="card stat-card"><div class="topline mint"></div>'
            f'<div class="card-label">POTENSI MAKSIMAL</div>'
            f'<div class="card-value" style="color:#10B981;">{potensi:.2f}</div>'
            f'<div class="card-note">{gap:+.2f} poin jika faktor dioptimalkan</div></div>',
            unsafe_allow_html=True,
        )

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
        f'<b style="color:#10B981;">{potensi:.2f}</b> '
        f'(selisih <b>+{potensi - nilai:.2f}</b> dari nilai saat ini).'
        f'</div></div>',
        unsafe_allow_html=True,
    )

    section_header("03", "Posisi Nilai Siswa", "PERBANDINGAN")

    fig, ax = plt.subplots(figsize=(10, 3.6))
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")
    posisi = [nilai, pred, potensi]
    labels = ["Nilai Sekarang", "Prediksi", "Potensi Maksimal"]
    colors = ["#2563EB", "#7C5CFC", "#10B981"]
    bars = ax.barh([0, 1, 2], posisi, height=.5, color=colors, alpha=.85)
    for i, p in enumerate(posisi):
        ax.text(p + 0.5, i, f"{p:.2f}", va="center", fontsize=11,
                fontweight="bold", color=colors[i])
    ax.axvline(RATA_RATA_NILAI, color="#EF5B67", linestyle="--", alpha=.6, linewidth=1.5)
    ax.text(RATA_RATA_NILAI, 2.6, f"Rata-rata sekolah ({RATA_RATA_NILAI:.2f})",
            color="#EF5B67", fontsize=9, fontweight="bold", ha="center")
    ax.set_yticks([0, 1, 2])
    ax.set_yticklabels(labels, fontsize=10)
    ax.set_xlim(60, 100)
    ax.tick_params(axis="y", length=0)
    ax.grid(axis="x", alpha=.15)
    ax.set_axisbelow(True)
    for s in ["top", "right", "left"]:
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color("#D8E0EB")
    ax.set_xlabel("Nilai Akademik (skala 60-100)", fontsize=9.5, color="#5C6B85", labelpad=8)
    plt.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

    with st.expander("Penjelasan tambahan untuk guru"):
        st.markdown(
            "Nilai prediksi diperoleh dari perhitungan yang mempertimbangkan kondisi 8 faktor siswa "
            "dibandingkan dengan rata-rata sekolah. Jika faktor-faktor yang mendukung sudah baik, "
            "prediksi akan cenderung naik. Jika ada faktor yang masih di bawah rata-rata, "
            "prediksi bisa tertekan. Gunakan informasi ini sebagai bahan pertimbangan untuk "
            "membantu siswa, bukan sebagai keputusan mutlak."
        )

# ================================================================
# MODUL 03 - FAKTOR PENGARUH (BAHASA DIRUBAH)
# ================================================================
elif st.session_state.current_page == "kausal":
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
        '</div></div>',
        unsafe_allow_html=True,
    )

    tab_pengaruh, tab_kepentingan = st.tabs(["Dampak Sebab-Akibat (Poin)", "Tingkat Kepentingan Prediksi"])

    with tab_pengaruh:
        df_ate = pd.DataFrame([{"Faktor": k, "Pengaruh": v} for k, v in PENGARUH_DATA.items()])
        pos_ate = df_ate[df_ate["Pengaruh"] > 0].sort_values("Pengaruh", ascending=False)
        neg_ate = df_ate[df_ate["Pengaruh"] < 0].sort_values("Pengaruh")

        a, b_, c = st.columns(3)
        with a:
            t = pos_ate.iloc[0]
            stat_card("Paling menaikkan nilai", t["Faktor"],
                      f"kontribusi +{t['Pengaruh']:.2f} poin", "mint")
        with b_:
            t = neg_ate.iloc[0]
            stat_card("Paling menurunkan nilai", t["Faktor"],
                      f"kontribusi {t['Pengaruh']:.2f} poin", "coral")
        with c:
            stat_card("Jumlah faktor", "8", "faktor dianalisis", "blue")

        st.markdown("<div style='height:.7rem'></div>", unsafe_allow_html=True)
        fig = plot_pengaruh(df_ate)
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

        tabel_ate = df_ate.copy()
        tabel_ate["Artinya"] = tabel_ate["Pengaruh"].apply(label_pengaruh)
        tabel_ate["Arah"] = tabel_ate["Pengaruh"].apply(
            lambda x: "Menaikkan" if x > 0 else "Menurunkan" if x < 0 else "Netral"
        )
        tabel_ate = tabel_ate.sort_values("Pengaruh", ascending=False)[
            ["Faktor", "Artinya", "Arah", "Pengaruh"]
        ]
        tabel_ate.columns = ["Faktor", "Artinya untuk Nilai Siswa", "Arah", "Skor Pengaruh"]
        st.dataframe(
            tabel_ate, use_container_width=True, hide_index=True,
            column_config={"Skor Pengaruh": st.column_config.NumberColumn(format="%+.4f")},
        )

    with tab_kepentingan:
        df_shap = pd.DataFrame([{"Faktor": k, "Kepentingan": v} for k, v in KEPENTINGAN_DATA.items()])
        ranked = df_shap.sort_values("Kepentingan", ascending=False).reset_index(drop=True)

        a, b_, c = st.columns(3)
        with a:
            stat_card("Paling Berpengaruh #1", ranked.iloc[0]["Faktor"], "paling menentukan prediksi", "purple")
        with b_:
            stat_card("Paling Berpengaruh #2", ranked.iloc[1]["Faktor"], "cukup menentukan prediksi", "blue")
        with c:
            stat_card("Jumlah faktor", "8", "faktor dianalisis", "yellow")

        st.markdown("<div style='height:.7rem'></div>", unsafe_allow_html=True)
        fig = plot_kepentingan(df_shap)
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

        for i, row in ranked.iterrows():
            pct = row["Kepentingan"] / ranked["Kepentingan"].max() * 100
            label = label_kepentingan(row["Kepentingan"])
            st.markdown(
                f'<div class="factor-card">'
                f'<div class="factor-head">'
                f'<div style="display:flex;align-items:center;gap:.7rem;">'
                f'<div class="rank-num">{i + 1:02d}</div>'
                f'<div><div class="factor-name">{row["Faktor"]}</div>'
                f'<div style="font-size:.7rem;color:#5C6B85;margin-top:.2rem;">{label}</div>'
                f'</div></div>'
                f'<div class="factor-value" style="color:#7C5CFC;">{row["Kepentingan"]:.4f}</div>'
                f'</div>'
                f'<div class="factor-bar"><div class="factor-fill" style="width:{pct:.1f}%"></div></div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    with st.expander("Penjelasan tambahan untuk guru"):
        st.markdown(
            "**Dampak Sebab-Akibat** menunjukkan apakah faktor tersebut benar-benar bisa mengubah nilai "
            "siswa jika diperbaiki. Misalnya: siswa yang lebih percaya diri cenderung mendapat nilai lebih baik.\n\n"
            "**Tingkat Kepentingan Prediksi** menunjukkan seberapa sering faktor tersebut dipakai oleh sistem "
            "untuk memperkirakan nilai siswa.\n\n"
            "**Catatan**: Gunakan hasil dampak sebab-akibat untuk memutuskan apa yang perlu diperbaiki. "
            "Gunakan tingkat kepentingan untuk memahami faktor mana yang paling sering dipertimbangkan."
        )

# ================================================================
# MODUL 04 - REKOMENDASI
# ================================================================
elif st.session_state.current_page == "rekomendasi":
    top_bar("Rekomendasi", "Modul 04 - Saran tindak lanjut untuk sekolah dan guru")

    hero_header(
        "MODUL 04 - REKOMENDASI",
        "Dari analisis<br><span class='accent'>menjadi tindakan.</span>",
        "Rekomendasi untuk tingkat sekolah dan catatan personal untuk guru per siswa.",
        [("TINGKAT SEKOLAH", "yellow"), ("PERSONAL", "mint"), ("TINDAK LANJUT", "blue")],
    )

    tab_umum, tab_personal = st.tabs(["Rekomendasi Tingkat Sekolah", "Catatan untuk Guru (Personal)"])

    with tab_umum:
        section_header("01", "Prioritas Intervensi", "FAKTOR DENGAN PENGARUH TERKUAT")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown(
                '<div class="card card-mint">'
                '<div class="card-label">FAKTOR YANG PERLU DIPERKUAT</div>'
                '<div style="font-family:Manrope;font-size:1.25rem;font-weight:800;margin-top:.45rem;">'
                'Faktor pendorong prestasi</div></div>',
                unsafe_allow_html=True,
            )
            pos_priority = [
                ("Self-Efficacy Akademik", "Dorong kepercayaan diri akademik melalui mentoring dan apresiasi proses."),
                ("Keterlibatan Orang Tua", "Perkuat komunikasi dan pendampingan belajar antara sekolah dan keluarga."),
                ("Kemalasan Belajar", "Pantau agar siswa tidak terjebak pola belajar pasif."),
            ]
            for i, (name, desc) in enumerate(pos_priority, 1):
                label = label_pengaruh(PENGARUH_DATA[name])
                st.markdown(
                    f'<div style="padding:1rem 0;border-bottom:1px solid #CBEBDD;">'
                    f'<div style="display:flex;justify-content:space-between;gap:.7rem;">'
                    f'<div style="font-weight:800;font-size:.88rem;">{i:02d} - {name}</div>'
                    f'<div style="font-weight:800;color:#10B981;font-size:.78rem;">{label}</div>'
                    f'</div>'
                    f'<div style="font-size:.76rem;line-height:1.55;margin-top:.45rem;color:#475467;">{desc}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

        with c2:
            st.markdown(
                '<div class="card card-yellow">'
                '<div class="card-label">FAKTOR YANG PERLU DIEVALUASI</div>'
                '<div style="font-family:Manrope;font-size:1.25rem;font-weight:800;margin-top:.45rem;">'
                'Faktor penghambat prestasi</div></div>',
                unsafe_allow_html=True,
            )
            neg_priority = [
                ("Motivasi Belajar", "Identifikasi hambatan belajar dan kaitkan materi dengan kehidupan nyata."),
                ("Dukungan Sekolah", "Evaluasi bentuk pendampingan agar tidak mengurangi kemandirian siswa."),
                ("Harapan Orang Tua", "Dorong target akademik realistis dan komunikasi tanpa tekanan berlebih."),
            ]
            for i, (name, desc) in enumerate(neg_priority, 1):
                label = label_pengaruh(PENGARUH_DATA[name])
                st.markdown(
                    f'<div style="padding:1rem 0;border-bottom:1px solid #F1DF96;">'
                    f'<div style="display:flex;justify-content:space-between;gap:.7rem;">'
                    f'<div style="font-weight:800;font-size:.88rem;">{i:02d} - {name}</div>'
                    f'<div style="font-weight:800;color:#EF5B67;font-size:.78rem;">{label}</div>'
                    f'</div>'
                    f'<div style="font-size:.76rem;line-height:1.55;margin-top:.45rem;color:#475467;">{desc}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

        st.markdown(
            '<div class="card card-dark" style="margin-top:1rem;">'
            '<div class="card-label">CATATAN</div>'
            '<div style="font-family:Manrope;font-size:1.35rem;line-height:1.25;font-weight:800;margin-top:.7rem;">'
            'Gunakan hasil sebab-akibat untuk memahami apa yang perlu diubah, '
            'dan hasil kepentingan prediksi untuk memahami faktor mana yang paling menentukan.'
            '</div></div>',
            unsafe_allow_html=True,
        )

    with tab_personal:
        if len(st.session_state.database_siswa) == 0:
            st.warning("Data siswa masih kosong. Input terlebih dahulu di Modul 01.")
        else:
            df_db = pd.DataFrame(st.session_state.database_siswa)
            opsi = df_db.apply(
                lambda x: f"{x['Nama']} - {x['Kelas']} (Absen {x['Absen']})", axis=1
            ).tolist()
            pilihan = st.selectbox("Pilih Siswa", opsi, key="rek_pilih")
            idx = opsi.index(pilihan)
            b = df_db.iloc[idx]

            nama = b["Nama"]; kelas = b["Kelas"]
            absen = int(b["Absen"]); nilai = float(b["Nilai Akademik"])
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
            neg = df_k[df_k["Kontribusi"] < -0.2].sort_values("Kontribusi")
            pos = df_k[df_k["Kontribusi"] > 0.2].sort_values("Kontribusi", ascending=False)

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
                f'<div class="card" style="border-left:5px solid {warna_kategori};margin-top:1rem;">'
                f'<div style="display:flex;align-items:center;gap:1rem;margin-bottom:1rem;">'
                f'<div class="profile-avatar" style="width:56px;height:56px;font-size:1.2rem;">'
                f'{"".join(x[0] for x in nama.split()[:2]).upper()}</div>'
                f'<div><div style="font-family:Manrope;font-weight:800;font-size:1.3rem;">{nama}</div>'
                f'<div style="color:#5C6B85;font-size:.8rem;">{kelas} - Absen {absen:02d} - Nilai {nilai:.2f}</div>'
                f'</div></div>'
                f'<div style="font-size:.95rem;line-height:1.75;color:#1F2A44;">'
                f'Nilai saat ini <b>{abs(selisih):.2f} poin {posisi} rata-rata sekolah</b> ({RATA_RATA_NILAI:.2f}).'
                f'<br><br>{narasi_neg}<br><br>{narasi_pos}<br><br>'
                f'<b>Kategori:</b> <span style="color:{warna_kategori};">{kategori}</span>'
                f'</div></div>',
                unsafe_allow_html=True,
            )

            section_header("02", "Prioritas Perbaikan", "FOKUSKAN PADA 3 HAL INI")
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
                            f'<div class="card" style="border-top:4px solid #EF5B67;min-height:200px;">'
                            f'<div style="font-size:.65rem;font-weight:800;letter-spacing:1.5px;color:#EF5B67;">PRIORITAS {i + 1}</div>'
                            f'<div style="font-family:Manrope;font-weight:800;font-size:1rem;margin:.6rem 0 .3rem;">{row["Aspek"]}</div>'
                            f'<div style="font-size:.7rem;color:#5C6B85;margin-bottom:.5rem;">{level}</div>'
                            f'<div style="font-size:.75rem;line-height:1.6;color:#475467;">'
                            f'Nilai siswa: <b>{row["Nilai_Siswa"]:.2f}</b><br>'
                            f'Rata-rata: <b>{row["Baseline"]:.2f}</b><br>'
                            f'<span style="color:#EF5B67;font-weight:700;">Selisih: {row["Selisih"]:+.2f}</span>'
                            f'</div></div>',
                            unsafe_allow_html=True,
                        )

            section_header("03", "Tindak Lanjut", "CHECKLIST GURU DAN SISWA")
            TINDAK = {
                "Self-Efficacy Akademik": {
                    "guru": ["Berikan tugas bertahap", "Pujian spesifik atas usaha", "Ajak refleksi mingguan", "Pasangkan dengan peer-mentor"],
                    "siswa": ["Jurnal harian 1 hal yang berhasil", "Tetapkan target kecil mingguan"],
                },
                "Keterlibatan Orang Tua": {
                    "guru": ["Kirim kabar positif ke orang tua", "Ajak orang tua ikut sesi belajar", "Panduan mendampingi belajar 15 menit/hari", "Komunikasi rutin 2 minggu sekali"],
                    "siswa": ["Ceritakan 1 hal yang dipelajari", "Minta orang tua periksa PR"],
                },
                "Harapan Orang Tua": {
                    "guru": ["Pertemuan ekspektasi realistis", "Bantu pahami tahap perkembangan anak", "Sarankan fokus pada usaha", "Contoh memotivasi tanpa menekan"],
                    "siswa": ["Belajar menyampaikan perasaan", "Fokus pada usaha yang bisa dikontrol"],
                },
                "Dukungan Sekolah": {
                    "guru": ["Refleksi bantuan berlebih", "Kurangi bantuan yang bisa dilakukan sendiri", "Berikan kesempatan mencoba", "Fokus membimbing, bukan menggantikan"],
                    "siswa": ["Coba selesaikan tugas 10 menit sebelum bertanya", "Catat apa yang sudah dicoba"],
                },
                "Motivasi Belajar": {
                    "guru": ["Kaitkan materi dengan kehidupan nyata", "Berikan pilihan tugas", "Apresiasi proses", "Ciptakan suasana kelas menyenangkan"],
                    "siswa": ["Cari 1 hal menarik dari tiap pelajaran", "Belajar bersama teman"],
                },
                "Kecemasan Akademik": {
                    "guru": ["Ajarkan teknik relaksasi", "Ubah suasana ujian lebih santai", "Ujian formatif yang tidak menakutkan", "Normalisasi cemas itu wajar"],
                    "siswa": ["Latihan pernapasan 4-7-8", "Persiapan lebih awal"],
                },
                "Fasilitas Sekolah": {
                    "guru": ["Informasikan fasilitas yang tersedia", "Bantu akses perpustakaan atau lab", "Cek hambatan akses"],
                    "siswa": ["Manfaatkan perpustakaan", "Tanyakan ke guru jika butuh bantuan"],
                },
                "Kemalasan Belajar": {
                    "guru": ["Cari akar kemalasan", "Beri tugas lebih menantang jika bosan", "Pecah tugas besar jadi langkah kecil", "Buat sistem reward sederhana"],
                    "siswa": ["Mulai dari tugas 5 menit", "Gunakan teknik Pomodoro"],
                },
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
                                st.markdown(
                                    f'<div style="display:flex;gap:.7rem;padding:.6rem 0;border-bottom:1px solid #EEF2F7;">'
                                    f'<div style="width:22px;height:22px;border-radius:6px;background:#EAF1FF;color:#2563EB;'
                                    f'display:flex;align-items:center;justify-content:center;font-size:.7rem;font-weight:800;">{j}</div>'
                                    f'<div style="font-size:.82rem;line-height:1.55;color:#1F2A44;">{t}</div></div>',
                                    unsafe_allow_html=True,
                                )
                        with c2:
                            st.markdown("**Yang bisa dilakukan SISWA:**")
                            for j, s in enumerate(tugas["siswa"], 1):
                                st.markdown(
                                    f'<div style="display:flex;gap:.7rem;padding:.6rem 0;border-bottom:1px solid #EEF2F7;">'
                                    f'<div style="width:22px;height:22px;border-radius:6px;background:#E8F8F1;color:#10B981;'
                                    f'display:flex;align-items:center;justify-content:center;font-size:.7rem;font-weight:800;">{j}</div>'
                                    f'<div style="font-size:.82rem;line-height:1.55;color:#1F2A44;">{s}</div></div>',
                                    unsafe_allow_html=True,
                                )

            section_header("04", "Download Catatan", "ARSIP GURU")
            faktor_utama = prioritas.iloc[0]["Aspek"] if len(prioritas) > 0 else "-"
            faktor_kekuatan = pos.iloc[0]["Aspek"] if len(pos) > 0 else "-"
            txt = (
                "CATATAN KONSULTASI GURU - SAA\n"
                "=================================\n"
                f"Nama Siswa  : {nama}\nKelas       : {kelas}\nAbsen       : {absen}\n"
                f"Nilai       : {nilai:.2f}\nKategori    : {kategori}\n\n"
                "KESIMPULAN\n----------\n"
                f"Nilai siswa {abs(selisih):.2f} poin {'di atas' if selisih >= 0 else 'di bawah'} rata-rata sekolah.\n\n"
                f"Faktor paling menekan : {faktor_utama}\n"
                f"Faktor kekuatan       : {faktor_kekuatan}\n\n"
                "PRIORITAS PERBAIKAN\n-------------------\n"
            )
            for i, (_, row) in enumerate(prioritas.iterrows(), 1):
                txt += f"{i}. {row['Aspek']} (selisih {row['Selisih']:+.2f})\n"
            txt += f"\nDibuat oleh : {st.session_state.user_nama}\n"
            txt += f"Tanggal     : {datetime.now().strftime('%d %B %Y, %H:%M')}\n"
            st.download_button(
                "Download Catatan (.txt)",
                data=txt.encode("utf-8"),
                file_name=f"catatan_{nama.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.txt",
                mime="text/plain",
                use_container_width=True,
            )

# ================================================================
# MODUL 05 - MONITORING KELAS & KELOLA DATA (TAMBAH BULK UPLOAD)
# ================================================================
elif st.session_state.current_page == "monitoring":
    top_bar("Monitoring & Data", "Modul 05 - Pantau kelas, download, dan upload massal")

    hero_header(
        "MODUL 05 - MONITORING & DATA",
        "Pantau kelas.<br><span class='accent'>Kelola data.</span>",
        "Monitoring agregat per kelas, download data siswa, upload massal dari Excel, dan kelola data.",
        [("PER KELAS", "yellow"), ("DOWNLOAD", "mint"), ("UPLOAD EXCEL", "blue")],
    )

    if len(st.session_state.database_siswa) == 0 and st.session_state.current_page != "monitoring":
        st.warning("Data siswa masih kosong. Silakan input atau upload.")
        st.stop()

    df_db = pd.DataFrame(st.session_state.database_siswa) if len(st.session_state.database_siswa) > 0 else pd.DataFrame()

    if st.session_state.user_role == "wali_kelas" and st.session_state.user_kelas and not df_db.empty:
        df_db = df_db[df_db["Kelas"].isin(st.session_state.user_kelas)]
        st.info(f"Anda login sebagai Wali Kelas. Data hanya untuk kelas: {', '.join(st.session_state.user_kelas)}")

    tab_monitoring, tab_early, tab_kelola, tab_download, tab_bulk = st.tabs([
        "📊 Monitoring Kelas", "⚠️ Early Warning", "🗑️ Kelola Data Siswa", "📥 Download Data", "📤 Upload Bulk Excel"
    ])

    with tab_monitoring:
        if df_db.empty:
            st.info("Belum ada data siswa untuk dimonitoring.")
        else:
            section_header("01", "Ringkasan Sekolah", "STATISTIK AGREGAT")
            a, b_, c, d = st.columns(4)
            with a:
                stat_card("Total Siswa", len(df_db), "siswa terarsip", "blue")
            with b_:
                stat_card("Rata-rata", f"{df_db['Nilai Akademik'].mean():.2f}", "nilai seluruh siswa", "mint")
            with c:
                stat_card("Tertinggi", f"{df_db['Nilai Akademik'].max():.2f}", "nilai maksimum", "yellow")
            with d:
                stat_card("Terendah", f"{df_db['Nilai Akademik'].min():.2f}", "nilai minimum", "coral")

            section_header("02", "Distribusi Kategori Nilai", "SEBARAN SELURUH SISWA")
            dist = df_db["Kategori"].value_counts().to_dict()
            kategori_list = ["Sangat Baik", "Baik", "Cukup", "Perlu Perhatian"]
            warna_map = {
                "Sangat Baik": "#10B981", "Baik": "#2563EB",
                "Cukup": "#F0B900", "Perlu Perhatian": "#EF5B67",
            }
            cols = st.columns(4)
            for i, kat in enumerate(kategori_list):
                jml = dist.get(kat, 0)
                pct = jml / len(df_db) * 100 if len(df_db) > 0 else 0
                with cols[i]:
                    st.markdown(
                        f'<div class="card stat-card" style="border-left:5px solid {warna_map[kat]};">'
                        f'<div class="card-label" style="color:{warna_map[kat]};">{kat.upper()}</div>'
                        f'<div class="card-value" style="color:{warna_map[kat]};">{jml}</div>'
                        f'<div class="card-note">{pct:.1f}% dari {len(df_db)} siswa</div>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

            section_header("03", "Perbandingan Antar Kelas", "RATA-RATA PER KELAS")
            per_kelas = df_db.groupby("Kelas").agg(
                Jumlah=("Nama", "count"),
                Rata2=("Nilai Akademik", "mean"),
                Tertinggi=("Nilai Akademik", "max"),
                Terendah=("Nilai Akademik", "min"),
            ).reset_index().sort_values("Rata2", ascending=False)
            per_kelas["Rata2"] = per_kelas["Rata2"].round(2)
            per_kelas["Tertinggi"] = per_kelas["Tertinggi"].round(2)
            per_kelas["Terendah"] = per_kelas["Terendah"].round(2)

            fig, ax = plt.subplots(figsize=(11, max(4, len(per_kelas) * 0.42)))
            fig.patch.set_alpha(0)
            ax.set_facecolor("none")
            colors = ["#10B981" if x >= RATA_RATA_NILAI else "#EF5B67" for x in per_kelas["Rata2"]]
            bars = ax.barh(per_kelas["Kelas"], per_kelas["Rata2"], color=colors, alpha=.85)
            ax.axvline(RATA_RATA_NILAI, color="#2563EB", linestyle="--", alpha=.6, linewidth=1.5)
            for bar, val in zip(bars, per_kelas["Rata2"]):
                ax.text(val + 0.15, bar.get_y() + bar.get_height() / 2, f"{val:.2f}",
                        va="center", fontsize=9.5, fontweight="bold")
            ax.set_xlim(70, 100)
            ax.tick_params(axis="y", length=0)
            ax.grid(axis="x", alpha=.15)
            ax.set_axisbelow(True)
            for s in ["top", "right", "left"]:
                ax.spines[s].set_visible(False)
            ax.spines["bottom"].set_color("#D8E0EB")
            ax.set_xlabel(f"Rata-rata nilai (garis biru = rata-rata sekolah {RATA_RATA_NILAI:.2f})",
                          fontsize=9.5, color="#5C6B85", labelpad=8)
            plt.tight_layout()
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)

            st.dataframe(per_kelas, use_container_width=True, hide_index=True)

            section_header("04", "Distribusi Nilai Per Kelas", "SEBARAN")
            kelas_pilih = st.selectbox("Pilih Kelas",
                                       sorted(df_db["Kelas"].unique().tolist()), key="mon_kelas")
            df_k = df_db[df_db["Kelas"] == kelas_pilih]
            if len(df_k) > 0:
                fig, ax = plt.subplots(figsize=(11, 4))
                fig.patch.set_alpha(0)
                ax.set_facecolor("none")
                ax.hist(df_k["Nilai Akademik"], bins=10, color="#2563EB", alpha=.7, edgecolor="white")
                ax.axvline(df_k["Nilai Akademik"].mean(), color="#EF5B67", linestyle="--", linewidth=2,
                           label=f"Rata-rata kelas: {df_k['Nilai Akademik'].mean():.2f}")
                ax.axvline(RATA_RATA_NILAI, color="#10B981", linestyle=":", linewidth=2,
                           label=f"Rata-rata sekolah: {RATA_RATA_NILAI:.2f}")
                ax.legend(fontsize=9)
                ax.grid(axis="y", alpha=.15)
                ax.set_axisbelow(True)
                for s in ["top", "right", "left"]:
                    ax.spines[s].set_visible(False)
                ax.spines["bottom"].set_color("#D8E0EB")
                ax.set_xlabel("Nilai Akademik", fontsize=9.5, color="#5C6B85")
                ax.set_ylabel("Jumlah Siswa", fontsize=9.5, color="#5C6B85")
                plt.tight_layout()
                st.pyplot(fig, use_container_width=True)
                plt.close(fig)

    with tab_early:
        if df_db.empty:
            st.info("Belum ada data siswa.")
        else:
            section_header("01", "Early Warning", "SISWA YANG PERLU PERHATIAN")
            st.markdown(
                '<div class="info-box coral" style="margin-bottom:1rem;">'
                '<div class="info-title">Kriteria</div>'
                '<div class="info-text">Siswa dengan nilai di bawah rata-rata sekolah lebih dari 1 standar deviasi.</div>'
                '</div>',
                unsafe_allow_html=True,
            )

            df_perhatian = df_db[df_db["Nilai Akademik"] < RATA_RATA_NILAI - STD_NILAI].copy()
            df_perhatian = df_perhatian.sort_values("Nilai Akademik")

            if len(df_perhatian) == 0:
                st.success("Tidak ada siswa yang memerlukan perhatian khusus saat ini.")
            else:
                st.markdown(f"**{len(df_perhatian)} siswa** memerlukan perhatian:")
                tabel_warn = df_perhatian[["Nama", "Kelas", "Absen", "Nilai Akademik",
                                           "Kategori", "Dicatat Oleh"]].copy()
                tabel_warn.columns = ["Nama", "Kelas", "Absen", "Nilai", "Kategori", "Dicatat Oleh"]
                st.dataframe(tabel_warn, use_container_width=True, hide_index=True)

                csv = tabel_warn.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "Download Daftar Early Warning (CSV)", data=csv,
                    file_name=f"early_warning_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv", use_container_width=True,
                )

    with tab_kelola:
        if df_db.empty:
            st.info("Belum ada data siswa.")
        else:
            section_header("01", "Kelola Data Siswa", "HAPUS SATU PER SATU ATAU SEMUA")

            st.markdown(
                '<div class="info-box coral" style="margin-bottom:1rem;">'
                '<div class="info-title">Perhatian</div>'
                '<div class="info-text">Data yang dihapus tidak bisa dikembalikan. '
                'Pastikan Anda sudah mendownload backup terlebih dahulu di tab <b>Download Data</b>.</div>'
                '</div>',
                unsafe_allow_html=True,
            )

            df_full = pd.DataFrame(st.session_state.database_siswa)

            col1, col2 = st.columns([2, 1])
            with col1:
                filter_hapus_kelas = st.selectbox(
                    "Filter Kelas", ["Semua"] + sorted(df_full["Kelas"].unique().tolist()),
                    key="kelola_filter",
                )
            with col2:
                cari_hapus = st.text_input("Cari nama", placeholder="Ketik nama siswa...",
                                           key="kelola_cari")

            df_hapus = df_full.copy()
            if filter_hapus_kelas != "Semua":
                df_hapus = df_hapus[df_hapus["Kelas"] == filter_hapus_kelas]
            if cari_hapus:
                df_hapus = df_hapus[df_hapus["Nama"].str.contains(cari_hapus, case=False, na=False)]

            st.markdown(f"**{len(df_hapus)} siswa** ditemukan")
            df_hapus_display = df_hapus[["Nama", "Kelas", "Absen", "Nilai Akademik", "Kategori", "Dicatat Oleh"]].copy()
            df_hapus_display.columns = ["Nama", "Kelas", "Absen", "Nilai", "Kategori", "Dicatat Oleh"]
            st.dataframe(df_hapus_display, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("**Hapus satu siswa:**")

            if len(df_hapus) == 0:
                st.info("Tidak ada siswa yang cocok dengan filter.")
            else:
                df_hapus_reset = df_hapus.reset_index(drop=True)
                opsi_hapus = df_hapus_reset.apply(
                    lambda x: f"{x['Nama']} - {x['Kelas']} (Absen {x['Absen']}) - Nilai {x['Nilai Akademik']:.2f}",
                    axis=1
                ).tolist()
                siswa_dipilih = st.selectbox("Pilih siswa yang akan dihapus", opsi_hapus,
                                             key="kelola_pilih")

                idx_hapus_local = opsi_hapus.index(siswa_dipilih)
                row_hapus = df_hapus_reset.iloc[idx_hapus_local]

                st.markdown(
                    f'<div class="card" style="border-left:5px solid #EF5B67;margin-top:.8rem;">'
                    f'<div class="card-label">SISWA YANG AKAN DIHAPUS</div>'
                    f'<div style="font-family:Manrope;font-weight:800;font-size:1.05rem;margin-top:.5rem;">'
                    f'{row_hapus["Nama"]}</div>'
                    f'<div style="font-size:.75rem;color:#5C6B85;margin-top:.3rem;">'
                    f'{row_hapus["Kelas"]} - Absen {int(row_hapus["Absen"]):02d} - Nilai {row_hapus["Nilai Akademik"]:.2f}'
                    f'</div></div>',
                    unsafe_allow_html=True,
                )

                konfirmasi_hapus = st.checkbox(
                    f"Saya yakin ingin menghapus data ini",
                    key="kelola_konfirmasi"
                )

                col_a, col_b = st.columns([1, 1])
                with col_a:
                    if st.button("Hapus Siswa Ini", use_container_width=True,
                                 type="primary", disabled=not konfirmasi_hapus, key="btn_hapus_satu"):
                        nama_hapus = row_hapus["Nama"]
                        kelas_hapus = row_hapus["Kelas"]
                        absen_hapus = int(row_hapus["Absen"])
                        for i, s in enumerate(st.session_state.database_siswa):
                            if (s["Nama"] == nama_hapus and s["Kelas"] == kelas_hapus
                                    and int(s["Absen"]) == absen_hapus):
                                st.session_state.database_siswa.pop(i)
                                break
                        save_database(st.session_state.database_siswa)
                        add_log("Hapus siswa", f"{nama_hapus} ({kelas_hapus})")
                        st.success(f"Data siswa {nama_hapus} berhasil dihapus.")
                        st.rerun()

            st.markdown("---")
            with st.expander("⚠️ Hapus SEMUA data siswa"):
                st.warning(
                    "Tombol di bawah ini akan menghapus SELURUH data siswa dari sistem. "
                    "Tindakan ini tidak bisa dibatalkan. Pastikan Anda sudah mendownload backup."
                )
                konfirmasi_semua = st.checkbox(
                    f"Saya mengerti, hapus SEMUA {len(st.session_state.database_siswa)} data siswa",
                    key="kelola_konfirmasi_semua",
                )
                if st.button("Hapus Semua Data", use_container_width=True,
                             type="primary", disabled=not konfirmasi_semua, key="btn_hapus_semua"):
                    jumlah = hapus_semua_siswa()
                    st.session_state.database_siswa = []
                    add_log("Hapus semua data", f"{jumlah} siswa dihapus")
                    st.success(f"{jumlah} data siswa berhasil dihapus.")
                    st.rerun()

    with tab_download:
        if df_db.empty:
            st.info("Belum ada data siswa.")
        else:
            section_header("01", "Download Data Siswa", "SEMUA DATA ATAU PER KELAS")

            st.markdown(
                '<div class="info-box blue" style="margin-bottom:1rem;">'
                '<div class="info-title">Format Download</div>'
                '<div class="info-text">Data akan diunduh dalam format CSV yang bisa dibuka dengan Excel, '
                'Google Sheets, atau aplikasi pengolah data lainnya.</div>'
                '</div>',
                unsafe_allow_html=True,
            )

            df_full = pd.DataFrame(st.session_state.database_siswa)

            st.markdown("**1. Download Semua Data**")
            csv_all = df_full.to_csv(index=False).encode("utf-8")
            st.download_button(
                f"Download Semua Data ({len(df_full)} siswa)",
                data=csv_all,
                file_name=f"data_semua_siswa_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True,
                key="dl_semua",
            )

            st.markdown("---")
            st.markdown("**2. Download Per Kelas**")
            kelas_pilih_dl = st.selectbox(
                "Pilih Kelas", sorted(df_full["Kelas"].unique().tolist()),
                key="dl_kelas_pilih",
            )
            df_kelas_dl = df_full[df_full["Kelas"] == kelas_pilih_dl]
            csv_kelas = df_kelas_dl.to_csv(index=False).encode("utf-8")
            st.download_button(
                f"Download Kelas {kelas_pilih_dl} ({len(df_kelas_dl)} siswa)",
                data=csv_kelas,
                file_name=f"data_{kelas_pilih_dl}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True,
                key="dl_kelas",
            )

    # ---------- TAB BULK UPLOAD ----------
    with tab_bulk:
        section_header("01", "Upload Data Massal (Excel)", "INPUT BANYAK SISWA SEKALIGUS")

        st.markdown(
            '<div class="info-box blue" style="margin-bottom:1rem;">'
            '<div class="info-title">Panduan Guru</div>'
            '<div class="info-text">'
            '1. Download template Excel di bawah<br>'
            '2. Isi data siswa di Excel (bisa puluhan/ratusan sekaligus)<br>'
            '3. Upload file Excel ke form di bawah<br>'
            '4. Klik "Simpan ke Database"'
            '</div></div>',
            unsafe_allow_html=True,
        )

        section_header("02", "Download Template", "FORMAT STANDAR")
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Data Siswa"
        
        headers = ["Nama Siswa", "Kelas", "Absen", "Nilai Rapor", 
                   "Self-Efficacy", "Keterlibatan Ortu", "Harapan Ortu", 
                   "Dukungan Sekolah", "Motivasi", "Kecemasan", 
                   "Fasilitas", "Kemalasan"]
        ws.append(headers)
        
        # Styling header
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            
        # Contoh baris
        contoh = [
            ["Budi Santoso", "IX-A", 1, 85.50, 4.5, 4.0, 3.5, 4.8, 3.2, 2.5, 4.5, 1.5],
            ["Siti Aminah", "IX-A", 2, 88.00, 4.8, 4.2, 3.8, 4.9, 3.5, 2.0, 4.8, 1.2],
        ]
        for row in contoh:
            ws.append(row)
            
        # Lebar kolom
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 18

        output = io.BytesIO()
        wb.save(output)
        
        st.download_button(
            "📥 Download Template Excel",
            data=output.getvalue(),
            file_name="Template_Upload_Siswa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("Hapus 2 baris contoh di atas sebelum mengisi data asli.")

        st.markdown("---")
        section_header("03", "Upload File Excel", "PROSES DATA")

        uploaded = st.file_uploader("Pilih file Excel (.xlsx)", type=["xlsx"])

        if uploaded:
            try:
                df_upload = pd.read_excel(uploaded)
                st.success(f"✅ Berhasil membaca {len(df_upload)} baris data.")
                
                st.markdown("**Preview Data:**")
                st.dataframe(df_upload.head(10), use_container_width=True) # Tampilkan 10 teratas

                # Validasi kolom
                required_cols = ["Nama Siswa", "Kelas", "Absen", "Nilai Rapor", 
                                 "Self-Efficacy", "Keterlibatan Ortu", "Harapan Ortu", 
                                 "Dukungan Sekolah", "Motivasi", "Kecemasan", 
                                 "Fasilitas", "Kemalasan"]
                missing = [c for c in required_cols if c not in df_upload.columns]
                
                if missing:
                    st.error(f"❌ Format salah. Kolom berikut tidak ditemukan: {', '.join(missing)}")
                else:
                    st.info("Pastikan data sudah benar. Klik tombol di bawah untuk menyimpan.")
                    
                    if st.button("💾 Simpan Semua Data ke Database", type="primary", use_container_width=True):
                        with st.spinner("Menyimpan data ke MongoDB..."):
                            data_baru = []
                            for _, row in df_upload.iterrows():
                                nilai = float(row["Nilai Rapor"])
                                kat = "Sangat Baik" if nilai >= 88 else "Baik" if nilai >= 84 else "Cukup" if nilai >= 80 else "Perlu Perhatian"
                                
                                data_baru.append({
                                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                    "Nama": str(row["Nama Siswa"]).strip(),
                                    "Kelas": str(row["Kelas"]).strip(),
                                    "Absen": int(row["Absen"]),
                                    "Nilai Akademik": round(nilai, 2),
                                    "Kategori": kat,
                                    "Self-Efficacy": float(row["Self-Efficacy"]),
                                    "Keterlibatan Ortu": float(row["Keterlibatan Ortu"]),
                                    "Harapan Ortu": float(row["Harapan Ortu"]),
                                    "Dukungan Sekolah": float(row["Dukungan Sekolah"]),
                                    "Motivasi": float(row["Motivasi"]),
                                    "Kecemasan": float(row["Kecemasan"]),
                                    "Fasilitas": float(row["Fasilitas"]),
                                    "Kemalasan": float(row["Kemalasan"]),
                                    "Dicatat Oleh": st.session_state.user_nama,
                                })
                            
                            # Simpan ke MongoDB sekaligus
                            for d in data_baru:
                                insert_siswa(d)
                                
                            add_log("Bulk Upload", f"{len(data_baru)} siswa")
                            st.success(f"🎉 Berhasil menyimpan {len(data_baru)} siswa ke database!")
                            st.balloons()
                            st.rerun()
            except Exception as e:
                st.error(f"❌ Gagal membaca file: {e}")

# ================================================================
# MODUL 06 - PENGATURAN
# ================================================================
elif st.session_state.current_page == "pengaturan":
    top_bar("Pengaturan", "Modul 06 - Manajemen sistem dan pengguna")

    if not can_access_pengaturan(st.session_state.user_role):
        st.error("Akses ditolak. Modul Pengaturan hanya untuk Administrator.")
        st.stop()

    hero_header(
        "MODUL 06 - PENGATURAN",
        "Konfigurasi<br><span class='accent'>dan tata kelola.</span>",
        "Kelola akun pengguna, konfigurasi sekolah, dan pantau aktivitas sistem.",
        [("ADMIN ONLY", "coral"), ("KONFIGURASI", "blue"), ("AUDIT", "purple")],
    )

    reload_users()
    users = st.session_state.users

    pending = [u for u, d in users.items() if d.get("status") == "pending"]
    approved = [u for u, d in users.items() if d.get("status") == "approved"]
    rejected = [u for u, d in users.items() if d.get("status") == "rejected"]

    tab_konfig, tab_user, tab_log = st.tabs([
        f"🏫 Konfigurasi Sekolah",
        f"👥 Manajemen Pengguna" + (f" ({len(pending)} menunggu)" if pending else ""),
        f"📋 Log Aktivitas",
    ])

    with tab_konfig:
        section_header("01", "Identitas Sekolah", "INFORMASI UMUM")
        cfg = st.session_state.config
        c1, c2 = st.columns(2)
        with c1:
            cfg["nama_sekolah"] = st.text_input("Nama Sekolah", value=cfg.get("nama_sekolah", ""))
            cfg["tahun_ajaran"] = st.text_input("Tahun Ajaran", value=cfg.get("tahun_ajaran", ""))
        with c2:
            cfg["semester"] = st.selectbox(
                "Semester", ["Ganjil", "Genap"],
                index=0 if cfg.get("semester") == "Ganjil" else 1,
            )
            cfg["kepala_sekolah"] = st.text_input("Nama Kepala Sekolah",
                                                  value=cfg.get("kepala_sekolah", ""))

        if st.button("Simpan Konfigurasi", type="primary"):
            st.session_state.config = cfg
            save_config(cfg)
            add_log("Update konfigurasi sekolah")
            st.success("Konfigurasi berhasil disimpan.")
            st.rerun()

    with tab_user:
        section_header("01", "Pengguna Aktif", f"{len(approved)} akun aktif")

        if len(approved) == 0:
            st.info("Belum ada pengguna aktif.")
        else:
            for u in approved:
                d = users[u]
                is_super = u in SUPER_ADMIN
                role_cls = d.get("role", "guru")
                st.markdown(
                    f'<div class="card" style="margin-bottom:.7rem;">'
                    f'<div style="display:flex;justify-content:space-between;align-items:center;gap:1rem;flex-wrap:wrap;">'
                    f'<div>'
                    f'<div style="font-family:Manrope;font-weight:800;font-size:1rem;">{d["nama"]}</div>'
                    f'<div style="font-size:.72rem;color:#5C6B85;margin-top:.25rem;">'
                    f'@{u} - {ROLE_LABEL.get(role_cls, role_cls)}'
                    + (f' - Kelas: {", ".join(d["kelas_ampu"])}' if d.get("kelas_ampu") else '')
                    + (f' - {d.get("email","")}' if d.get("email") else '')
                    + f'</div></div>'
                    f'<div>'
                    + (f'<span class="status-badge approved">Super Admin</span>' if is_super
                       else f'<span class="status-badge approved">Aktif</span>')
                    + f'</div></div></div>',
                    unsafe_allow_html=True,
                )
                if not is_super:
                    cc1, cc2, cc3, cc4 = st.columns([1, 1, 1, 3])
                    with cc1:
                        if st.button("Reset Password", key=f"reset_{u}", use_container_width=True):
                            st.session_state[f"show_reset_{u}"] = not st.session_state.get(f"show_reset_{u}", False)
                    with cc2:
                        if st.button("Nonaktifkan", key=f"reject_{u}", use_container_width=True):
                            users[u]["status"] = "rejected"
                            save_users(users)
                            reload_users()
                            add_log("Nonaktifkan user", u)
                            st.success(f"User {u} dinonaktifkan.")
                            st.rerun()
                    with cc3:
                        if st.button("Hapus", key=f"del_{u}", use_container_width=True):
                            del users[u]
                            save_users(users)
                            reload_users()
                            add_log("Hapus user", u)
                            st.success(f"User {u} dihapus.")
                            st.rerun()
                    if st.session_state.get(f"show_reset_{u}", False):
                        with st.form(f"form_reset_{u}"):
                            new_pw = st.text_input(f"Kata sandi baru untuk {u}", type="password",
                                                   key=f"newpw_{u}")
                            if st.form_submit_button("Simpan Password Baru"):
                                if len(new_pw) < 6:
                                    st.error("Kata sandi minimal 6 karakter.")
                                else:
                                    users[u]["password"] = hash_password(new_pw)
                                    save_users(users)
                                    reload_users()
                                    add_log("Reset password", u)
                                    st.session_state[f"show_reset_{u}"] = False
                                    st.success(f"Password {u} berhasil direset.")
                                    st.rerun()

        st.markdown("---")
        section_header("02", "Menunggu Persetujuan", f"{len(pending)} akun menunggu")

        if len(pending) == 0:
            st.success("Tidak ada akun yang menunggu persetujuan.")
        else:
            for u in pending:
                d = users[u]
                role_cls = d.get("role", "guru")
                st.markdown(
                    f'<div class="card" style="border-left:5px solid #F6C945;margin-bottom:.7rem;">'
                    f'<div style="display:flex;justify-content:space-between;align-items:center;gap:1rem;flex-wrap:wrap;">'
                    f'<div>'
                    f'<div style="font-family:Manrope;font-weight:800;font-size:1rem;">{d["nama"]}</div>'
                    f'<div style="font-size:.72rem;color:#5C6B85;margin-top:.25rem;">'
                    f'@{u} - {ROLE_LABEL.get(role_cls, role_cls)}'
                    + (f' - Kelas: {", ".join(d["kelas_ampu"])}' if d.get("kelas_ampu") else '')
                    + (f' - {d.get("email","")}' if d.get("email") else '')
                    + f'</div>'
                    f'<div style="font-size:.68rem;color:#92400E;margin-top:.3rem;">Didaftarkan: {d.get("created","-")}</div>'
                    f'</div>'
                    f'<div><span class="status-badge pending">Menunggu</span></div>'
                    f'</div></div>',
                    unsafe_allow_html=True,
                )
                cc1, cc2, cc3 = st.columns([1, 1, 3])
                with cc1:
                    if st.button("Setujui", key=f"approve_{u}", use_container_width=True, type="primary"):
                        users[u]["status"] = "approved"
                        users[u]["approved_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                        save_users(users)
                        reload_users()
                        add_log("Approve user", f"{u} ({d['nama']})")
                        st.success(f"Akun {d['nama']} disetujui.")
                        st.rerun()
                with cc2:
                    if st.button("Tolak", key=f"rej_{u}", use_container_width=True):
                        users[u]["status"] = "rejected"
                        save_users(users)
                        reload_users()
                        add_log("Reject user", f"{u} ({d['nama']})")
                        st.warning(f"Akun {d['nama']} ditolak.")
                        st.rerun()

        if len(rejected) > 0:
            st.markdown("---")
            section_header("03", "Akun Ditolak", f"{len(rejected)} akun ditolak")
            for u in rejected:
                d = users[u]
                role_cls = d.get("role", "guru")
                st.markdown(
                    f'<div class="card" style="border-left:5px solid #EF5B67;margin-bottom:.7rem;">'
                    f'<div style="display:flex;justify-content:space-between;align-items:center;gap:1rem;flex-wrap:wrap;">'
                    f'<div>'
                    f'<div style="font-family:Manrope;font-weight:800;font-size:1rem;">{d["nama"]}</div>'
                    f'<div style="font-size:.72rem;color:#5C6B85;margin-top:.25rem;">'
                    f'@{u} - {ROLE_LABEL.get(role_cls, role_cls)}</div>'
                    f'</div><div><span class="status-badge rejected">Ditolak</span></div>'
                    f'</div></div>',
                    unsafe_allow_html=True,
                )
                cc1, cc2, cc3 = st.columns([1, 1, 3])
                with cc1:
                    if st.button("Aktifkan", key=f"react_{u}", use_container_width=True):
                        users[u]["status"] = "approved"
                        save_users(users)
                        reload_users()
                        add_log("Reaktivasi user", u)
                        st.success(f"Akun {u} diaktifkan kembali.")
                        st.rerun()
                with cc2:
                    if st.button("Hapus Permanen", key=f"delrej_{u}", use_container_width=True):
                        del users[u]
                        save_users(users)
                        reload_users()
                        add_log("Hapus user ditolak", u)
                        st.success(f"Akun {u} dihapus permanen.")
                        st.rerun()

        st.markdown("---")
        section_header("04", "Tambah Pengguna Manual", "OPSIONAL")
        st.caption("Gunakan form ini jika ingin menambah pengguna langsung tanpa proses approval.")
        with st.form("form_user_baru"):
            c1, c2 = st.columns(2)
            with c1:
                new_u = st.text_input("Username", placeholder="huruf kecil, tanpa spasi")
                new_nama = st.text_input("Nama Lengkap")
                new_role = st.selectbox("Peran",
                                        ["admin", "kepala_sekolah", "wali_kelas", "guru"],
                                        format_func=lambda x: ROLE_LABEL.get(x, x))
            with c2:
                new_pw = st.text_input("Kata Sandi", type="password")
                new_email = st.text_input("Email (opsional)")
                new_kelas = st.text_input("Kelas Diampu (pisah koma)",
                                          placeholder="IX-A, IX-B")
            submit_user = st.form_submit_button("Tambah Pengguna", use_container_width=True)

            if submit_user:
                if not new_u or not new_pw or not new_nama:
                    st.error("Username, nama, dan kata sandi wajib diisi.")
                elif len(new_pw) < 6:
                    st.error("Kata sandi minimal 6 karakter.")
                elif new_u in users:
                    st.error(f"Username '{new_u}' sudah ada.")
                else:
                    kelas_ampu = [k.strip() for k in new_kelas.split(",") if k.strip()] if new_kelas else None
                    users[new_u] = {
                        "password": hash_password(new_pw),
                        "nama": new_nama, "role": new_role,
                        "kelas_ampu": kelas_ampu,
                        "email": new_email.strip(),
                        "status": "approved",
                        "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "approved_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    }
                    save_users(users)
                    reload_users()
                    add_log("Tambah user manual", f"{new_u} ({new_role})")
                    st.success(f"Pengguna '{new_u}' berhasil ditambahkan.")
                    st.rerun()

    with tab_log:
        section_header("01", "Log Aktivitas", "CATATAN SISTEM")
        st.markdown(
            '<div class="info-box blue" style="margin-bottom:1rem;">'
            '<div class="info-title">Info</div>'
            '<div class="info-text">Log menampilkan 200 aktivitas terakhir.</div>'
            '</div>',
            unsafe_allow_html=True,
        )

        if len(st.session_state.log) == 0:
            st.info("Belum ada aktivitas tercatat.")
        else:
            df_log = pd.DataFrame(st.session_state.log)
            st.dataframe(df_log, use_container_width=True, hide_index=True)

            csv_log = df_log.to_csv(index=False).encode("utf-8")
            st.download_button(
                "Download Log (CSV)", data=csv_log,
                file_name=f"log_saa_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv", use_container_width=True,
            )

            if st.button("Bersihkan Log"):
                st.session_state.log = []
                save_log([])
                st.rerun()

# ================================================================
# END OF FILE
# ================================================================