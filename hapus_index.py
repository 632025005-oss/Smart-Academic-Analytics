import toml
from pymongo import MongoClient
from pymongo.errors import OperationFailure

secrets = toml.load(".streamlit/secrets.toml")
client = MongoClient(secrets["MONGODB_URI"])
db = client.saa_smpn6

print("Mencoba menghapus index lama...")
try:
    db.siswa.drop_index("Kelas_1_Absen_1")
    print("✅ Index (Kelas + Absen) berhasil dihapus.")
except OperationFailure:
    print("⚠️ Index (Kelas + Absen) sudah tidak ada atau tidak ditemukan.")

# Cek index yang tersisa
print("\nIndex yang tersisa:")
for idx in db.siswa.list_indexes():
    print(f" - {idx['name']}: {idx['key']}")