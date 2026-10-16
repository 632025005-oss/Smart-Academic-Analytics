# init_mongodb.py
# ================================================================
# SETUP AWAL MongoDB untuk SAA - SMP Negeri 6 Salatiga
# Jalankan SEKALI: python init_mongodb.py
# ================================================================
import hashlib
from datetime import datetime
from pymongo import MongoClient, ASCENDING, DESCENDING
import toml
import os

# Baca secrets.toml manual
def baca_secrets():
    path = os.path.join(".streamlit", "secrets.toml")
    with open(path, "r", encoding="utf-8") as f:
        return toml.load(f)

secrets_data = baca_secrets()
MONGODB_URI = secrets_data["MONGODB_URI"]
SUPER_ADMIN = secrets_data.get("super_admin", {})

print("=" * 60)
print("INISIALISASI DATABASE SAA - SMPN 6 SALATIGA")
print("=" * 60)

client = MongoClient(MONGODB_URI)
db = client.saa_smpn6

# ================================================================
# 1. USERS
# ================================================================
print("\n[1/5] Setup collection: users")
users_col = db.users
users_col.create_index("role")
users_col.create_index("status")

for username, info in SUPER_ADMIN.items():
    if not users_col.find_one({"_id": username}):
        users_col.insert_one({
            "_id": username,
            "password": hashlib.sha256(info["password"].encode()).hexdigest(),
            "nama": info.get("nama", username.capitalize()),
            "role": "admin",
            "kelas_ampu": None,
            "email": info.get("email", ""),
            "status": "approved",
            "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
        })
        print(f"   ✅ Super admin '{username}' dibuat")
    else:
        print(f"   ⏭️  Super admin '{username}' sudah ada")

# ================================================================
# 2. SISWA (data mentah: 37 item + nilai rapor)
# ================================================================
print("\n[2/5] Setup collection: siswa")
siswa_col = db.siswa
siswa_col.create_index([("Kelas", ASCENDING), ("Absen", ASCENDING)], unique=True)
siswa_col.create_index("Nama")
siswa_col.create_index("Nilai_Akademik")
siswa_col.create_index("Timestamp")
print("   ✅ Index dibuat")

# ================================================================
# 3. CONFIG
# ================================================================
print("\n[3/5] Setup collection: config")
config_col = db.config
if not config_col.find_one({"_id": "main"}):
    config_col.insert_one({
        "_id": "main",
        "nama_sekolah": "SMP Negeri 6 Salatiga",
        "tahun_ajaran": "2025/2026",
        "semester": "Genap",
        "kepala_sekolah": "Kepala SMPN 6 Salatiga",
    })
    print("   ✅ Config default dibuat")
else:
    print("   ⏭️  Config sudah ada")

# ================================================================
# 4. LOG
# ================================================================
print("\n[4/5] Setup collection: log")
log_col = db.log
log_col.create_index([("waktu", DESCENDING)])
log_col.create_index("user")
print("   ✅ Index dibuat")

# ================================================================
# 5. HASIL_ANALISIS
# ================================================================
print("\n[5/5] Setup collection: hasil_analisis")
hasil_col = db.hasil_analisis
hasil_col.create_index([("tanggal", DESCENDING)])
hasil_col.create_index("jenis")  # 'ate' | 'shap' | 'cei' | 'kuadran' | 'konstruk' | 'model'
print("   ✅ Index dibuat")

# ================================================================
# RINGKASAN
# ================================================================
print("\n" + "=" * 60)
print("✅ INISIALISASI SELESAI")
print("=" * 60)
print(f"\nDatabase: {db.name}")
print("Collections:")
for col_name in db.list_collection_names():
    count = db[col_name].count_documents({})
    print(f"  - {col_name:20s} : {count} dokumen")
print("\n🎉 MongoDB siap digunakan!")