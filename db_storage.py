import streamlit as st
from pymongo import MongoClient
from bson.objectid import ObjectId
from datetime import datetime

@st.cache_resource
def _get_client():
    return MongoClient(st.secrets["MONGODB_URI"])

def _db():
    return _get_client()["saa_smpn6"]

# ================================================================
# USERS
# ================================================================
def load_users():
    users = {}
    for doc in _db().users.find():
        username = doc.pop("_id")
        users[username] = doc
    return users

def save_users(users_dict):
    col = _db().users
    col.delete_many({})
    if users_dict:
        docs = [{"_id": u, **data} for u, data in users_dict.items()]
        col.insert_many(docs)

# ================================================================
# SISWA (DENGAN FILTER PRIVASI GURU)
# ================================================================
def load_database(guru_nama=None):
    """Hanya memuat data milik guru yang sedang login."""
    query = {}
    if guru_nama:
        query["Dicatat Oleh"] = guru_nama
        
    siswa = []
    for doc in _db().siswa.find(query).sort("Timestamp", 1):
        doc["_id"] = str(doc["_id"])
        siswa.append(doc)
    return siswa

def insert_siswa(data):
    # Pastikan 'Dicatat Oleh' diisi dengan nama guru
    if "Dicatat Oleh" not in data or not data["Dicatat Oleh"]:
        data["Dicatat Oleh"] = "Sistem"
    return _db().siswa.insert_one(dict(data)).inserted_id

def hapus_siswa(kelas, absen, guru_nama):
    """Hapus hanya jika milik guru tersebut."""
    return _db().siswa.delete_one({
        "Kelas": kelas, 
        "Absen": int(absen),
        "Dicatat Oleh": guru_nama
    })

def hapus_semua_siswa(guru_nama):
    """Hapus semua data milik guru tertentu saja."""
    return _db().siswa.delete_many({"Dicatat Oleh": guru_nama}).deleted_count

def update_kuesioner_siswa(nama, kelas, absen, profil_faktor):
    db = _db().siswa
    # Cari siswa berdasarkan Nama, Kelas, dan Absen (case-insensitive)
    siswa = db.find_one({
        "Nama": {"$regex": f"^{nama.strip()}$", "$options": "i"},
        "Kelas": kelas,
        "Absen": int(absen)
    })
    
    if siswa:
        db.update_one({"_id": siswa["_id"]}, {"$set": profil_faktor})
        return True, "Data kuesioner berhasil ditambahkan ke data siswa yang sudah ada."
    else:
        data_baru = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Nama": nama.strip(),
            "Kelas": kelas,
            "Absen": int(absen),
            "Nilai Akademik": 0.0,
            "Kategori": "Belum Dinilai",
            **profil_faktor,
            "Dicatat Oleh": "Siswa (Mandiri)"
        }
        db.insert_one(data_baru)
        return True, "Siswa belum ada di database guru, data kuesioner disimpan sebagai data baru."

# ================================================================
# CONFIG & LOG
# ================================================================
def load_config():
    doc = _db().config.find_one({"_id": "main"})
    if doc:
        doc.pop("_id")
        return doc
    return {"nama_sekolah": "SMP Negeri 6 Salatiga", "tahun_ajaran": "2025/2026"}

def save_config(cfg):
    _db().config.replace_one({"_id": "main"}, {**cfg, "_id": "main"}, upsert=True)

def load_log(limit=200):
    return list(_db().log.find({}, {"_id": 0}).sort("waktu", -1).limit(limit))

def save_log(log_list):
    col = _db().log
    col.delete_many({})
    if log_list:
        col.insert_many([dict(l) for l in log_list[:200]])

def tambah_log(entry):
    _db().log.insert_one(dict(entry))
    total = _db().log.count_documents({})
    if total > 200:
        oldest = list(_db().log.find().sort("waktu", 1).limit(total - 200))
        ids = [d["_id"] for d in oldest]
        _db().log.delete_many({"_id": {"$in": ids}})