import toml
from pymongo import MongoClient, ASCENDING

secrets = toml.load(".streamlit/secrets.toml")
client = MongoClient(secrets["MONGODB_URI"])
db = client.saa_smpn6

# 1. Hapus index lama yang bikin error
try:
    db.siswa.drop_index("Kelas_1_Absen_1")
    print("✅ Index lama (Kelas + Absen) berhasil dihapus.")
except Exception as e:
    print("⚠️ Index lama tidak ditemukan atau sudah dihapus.")

# 2. Bersihkan data siswa lama (karena kamu memang mau menghapus 117 siswa sebelumnya)
db.siswa.delete_many({})
print("✅ Data siswa lama dibersihkan.")

# 3. Buat index baru (Kelas + Absen + Nama Guru)
# Ini memungkinkan Guru A dan Guru B punya siswa IX-A absen 1 masing-masing.
try:
    db.siswa.create_index(
        [("Kelas", ASCENDING), ("Absen", ASCENDING), ("Dicatat Oleh", ASCENDING)], 
        unique=True
    )
    print("✅ Index baru berhasil dibuat.")
except Exception as e:
    print("❌ Gagal membuat index baru:", e)

print("\n🎉 Perbaikan Selesai! Silakan coba input lagi di Streamlit.")