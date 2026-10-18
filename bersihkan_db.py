import toml
from pymongo import MongoClient

secrets = toml.load(".streamlit/secrets.toml")
client = MongoClient(secrets["MONGODB_URI"])
db = client.saa_smpn6

hasil = db.siswa.delete_many({})
print(f"✅ {hasil.deleted_count} data siswa lama berhasil dihapus!")