# test_koneksi.py
import os
import toml
from pymongo import MongoClient

print("=" * 60)
print("TEST KONEKSI MONGODB ATLAS")
print("=" * 60)

path = os.path.join(".streamlit", "secrets.toml")
with open(path, "r", encoding="utf-8") as f:
    secrets_data = toml.load(f)

uri = secrets_data["MONGODB_URI"]
print(f"\n📡 URI: {uri[:60]}...")

try:
    print("\n⏳ Menghubungkan...")
    client = MongoClient(uri, serverSelectionTimeoutMS=10000)
    client.admin.command('ping')
    print("✅ KONEKSI BERHASIL!")

    db = client.saa_smpn6
    print(f"\n📊 Database: {db.name}")
    cols = db.list_collection_names()
    if cols:
        for c in cols:
            print(f"  - {c}: {db[c].count_documents({})} dokumen")
    else:
        print("  (belum ada collection — jalankan init_mongodb.py)")

    client.close()
    print("\n🎉 Selesai!")

except Exception as e:
    print(f"\n❌ GAGAL: {type(e).__name__}")
    print(f"   {str(e)[:200]}")