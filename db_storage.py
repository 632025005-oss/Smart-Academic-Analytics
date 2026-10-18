import streamlit as st
from pymongo import MongoClient
from bson.objectid import ObjectId
from datetime import datetime

# ================================================================
# KONEKSI (cached, aman untuk Streamlit)
# ================================================================
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
# SISWA
# ================================================================
def load_database(guru_nama=None):
    query = {}
    if guru_nama:
        query["Dicatat Oleh"] = guru_nama
        
    siswa = []
    for doc in _db().siswa.find(query).sort("Timestamp", 1):
        doc["_id"] = str(doc["_id"])
        siswa.append(doc)
    return siswa

def insert_siswa(data):
    existing = _db().siswa.find_one({
        "Kelas": data.get("Kelas"),
        "Absen": int(data.get("Absen", 0)),
        "Dicatat Oleh": data.get("Dicatat Oleh")
    })
    
    if existing:
        _db().siswa.update_one({"_id": existing["_id"]}, {"$set": dict(data)})
        return existing["_id"]
    else:
        return _db().siswa.insert_one(dict(data)).inserted_id

def hapus_semua_siswa(guru_nama):
    return _db().siswa.delete_many({"Dicatat Oleh": guru_nama}).deleted_count

# FUNGSI BARU UNTUK ADMIN
def load_all_siswa():
    """Memuat SEMUA data siswa dari semua guru (Khusus Admin)."""
    siswa = []
    for doc in _db().siswa.find().sort("Kelas", 1):
        doc["_id"] = str(doc["_id"])
        siswa.append(doc)
    return siswa

def hapus_siswa_by_id(siswa_id):
    """Menghapus data siswa berdasarkan ID (Khusus Admin)."""
    _db().siswa.delete_one({"_id": ObjectId(siswa_id)})

def update_kuesioner_siswa(nama, kelas, absen, profil_faktor):
    db = _db().siswa
    try:
        siswa_absen = db.find_one({"Kelas": kelas, "Absen": int(absen)})
        
        if siswa_absen:
            if siswa_absen["Nama"].strip().lower() == nama.strip().lower():
                db.update_one({"_id": siswa_absen["_id"]}, {"$set": profil_faktor})
                return True, "Data kuesioner berhasil diperbarui."
            else:
                return False, f"❌ Gagal: Kelas {kelas} Absen {absen} sudah diisi oleh {siswa_absen['Nama']}. Silakan periksa kembali nomor absenmu."
        
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
        return True, "Data kuesioner berhasil disimpan. (Menunggu guru menginput nilai)."
    except Exception as e:
        return False, f"Gagal menyimpan: Terjadi kesalahan sistem. ({str(e)})"

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

def load_log(user=None, limit=200):
    query = {}
    if user:
        query["user"] = user
    return list(_db().log.find(query, {"_id": 0}).sort("waktu", -1).limit(limit))

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