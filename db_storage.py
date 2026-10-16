# db_storage.py
# ================================================================
# Layer MongoDB untuk Streamlit App SAA - SMP Negeri 6 Salatiga
# Semua operasi baca/tulis ke MongoDB ada di sini
# ================================================================
import streamlit as st
from pymongo import MongoClient
from bson.objectid import ObjectId
from datetime import datetime

# ================================================================
# KONEKSI (cached, aman untuk Streamlit)
# ================================================================
@st.cache_resource
def _get_client():
    """Buat koneksi ke MongoDB (hanya sekali, lalu di-cache)."""
    return MongoClient(st.secrets["MONGODB_URI"])

def _db():
    """Ambil database saa_smpn6."""
    return _get_client()["saa_smpn6"]

# ================================================================
# USERS — Manajemen Akun
# ================================================================
def load_users():
    """Return dict: {username: {password, nama, role, ...}}"""
    users = {}
    for doc in _db().users.find():
        username = doc.pop("_id")
        users[username] = doc
    return users

def save_users(users_dict):
    """Replace seluruh collection users."""
    col = _db().users
    col.delete_many({})
    if users_dict:
        docs = [{"_id": u, **data} for u, data in users_dict.items()]
        col.insert_many(docs)

def get_user(username):
    """Ambil satu user berdasarkan username."""
    doc = _db().users.find_one({"_id": username})
    if doc:
        doc.pop("_id")
    return doc

def upsert_user(username, data):
    """Insert atau update user."""
    _db().users.replace_one(
        {"_id": username},
        {"_id": username, **data},
        upsert=True
    )

def delete_user(username):
    """Hapus user."""
    _db().users.delete_one({"_id": username})

# ================================================================
# SISWA — Data Siswa + 8 Faktor + Nilai
# ================================================================
def load_database():
    """Return list of siswa (kompatibel dengan kode lama)."""
    siswa = []
    for doc in _db().siswa.find().sort("Timestamp", 1):
        doc["_id"] = str(doc["_id"])
        siswa.append(doc)
    return siswa

def save_database(siswa_list):
    """Replace seluruh collection siswa."""
    col = _db().siswa
    col.delete_many({})
    if siswa_list:
        docs = []
        for s in siswa_list:
            d = {k: v for k, v in s.items() if k != "_id"}
            docs.append(d)
        col.insert_many(docs)

def insert_siswa(data):
    """Tambah 1 siswa (lebih efisien untuk operasi tunggal)."""
    return _db().siswa.insert_one(dict(data)).inserted_id

def cek_duplikat(kelas, absen):
    """Cek apakah kelas+absen sudah dipakai."""
    return _db().siswa.find_one({"Kelas": kelas, "Absen": int(absen)})

def hapus_siswa(kelas, absen):
    """Hapus satu siswa berdasarkan kelas + absen."""
    return _db().siswa.delete_one({"Kelas": kelas, "Absen": int(absen)})

def hapus_semua_siswa():
    """Hapus semua data siswa."""
    return _db().siswa.delete_many({}).deleted_count

def get_siswa_by_kelas(kelas):
    """Ambil semua siswa di kelas tertentu."""
    return list(_db().siswa.find({"Kelas": kelas}))

# ================================================================
# CONFIG — Konfigurasi Sekolah
# ================================================================
def load_config():
    """Ambil konfigurasi sekolah."""
    doc = _db().config.find_one({"_id": "main"})
    if doc:
        doc.pop("_id")
        return doc
    return {
        "nama_sekolah": "SMP Negeri 6 Salatiga",
        "tahun_ajaran": "2025/2026",
        "semester": "Genap",
        "kepala_sekolah": "Kepala SMPN 6 Salatiga",
    }

def save_config(cfg):
    """Simpan konfigurasi sekolah."""
    _db().config.replace_one(
        {"_id": "main"},
        {**cfg, "_id": "main"},
        upsert=True
    )

# ================================================================
# LOG — Audit Trail Aktivitas
# ================================================================
def load_log(limit=200):
    """Ambil 200 log terbaru."""
    return list(_db().log.find({}, {"_id": 0}).sort("waktu", -1).limit(limit))

def save_log(log_list):
    """Replace seluruh collection log."""
    col = _db().log
    col.delete_many({})
    if log_list:
        col.insert_many([dict(l) for l in log_list[:200]])

def tambah_log(entry):
    """Tambah 1 log entry (lebih efisien dari replace)."""
    _db().log.insert_one(dict(entry))
    # Batasi 200 terbaru
    total = _db().log.count_documents({})
    if total > 200:
        oldest = list(_db().log.find().sort("waktu", 1).limit(total - 200))
        ids = [d["_id"] for d in oldest]
        _db().log.delete_many({"_id": {"$in": ids}})

def bersihkan_log():
    """Hapus semua log."""
    _db().log.delete_many({})

# ================================================================
# HASIL_ANALISIS — Hasil dari Script Analisis di Laptop
# ================================================================
def simpan_hasil(judul, metode, deskripsi, dataframe, user="-"):
    """Simpan hasil analisis (dari laptop) ke MongoDB."""
    data_records = dataframe.to_dict(orient="records")
    _db().hasil_analisis.insert_one({
        "judul": judul,
        "metode": metode,
        "deskripsi": deskripsi,
        "kolom": list(dataframe.columns),
        "jumlah_baris": len(data_records),
        "data": data_records,
        "dibuat_oleh": user,
        "tanggal": datetime.now(),
    })

def ambil_semua_hasil():
    """Ambil semua hasil analisis, terbaru dulu."""
    hasil = []
    for doc in _db().hasil_analisis.find().sort("tanggal", -1):
        doc["_id"] = str(doc["_id"])
        hasil.append(doc)
    return hasil

def ambil_hasil_by_id(hasil_id):
    """Ambil satu hasil analisis berdasarkan ID."""
    doc = _db().hasil_analisis.find_one({"_id": ObjectId(hasil_id)})
    if doc:
        doc["_id"] = str(doc["_id"])
    return doc

def hapus_hasil(hasil_id):
    """Hapus hasil analisis berdasarkan ID."""
    _db().hasil_analisis.delete_one({"_id": ObjectId(hasil_id)})

def hapus_semua_hasil():
    """Hapus semua hasil analisis."""
    return _db().hasil_analisis.delete_many({}).deleted_count

# ================================================================
# UTILITAS
# ================================================================
def test_koneksi():
    """Test apakah MongoDB bisa diakses."""
    try:
        _db().command("ping")
        return True, "Koneksi MongoDB OK"
    except Exception as e:
        return False, str(e)

def get_statistik_db():
    """Ambil statistik database (untuk dashboard admin)."""
    db = _db()
    return {
        "users": db.users.count_documents({}),
        "siswa": db.siswa.count_documents({}),
        "log": db.log.count_documents({}),
        "hasil_analisis": db.hasil_analisis.count_documents({}),
    }