# test_storage.py
# Test db_storage.py tanpa Streamlit
import sys
import toml
from unittest.mock import MagicMock

# ================================================================
# MOCK STREAMLIT
# ================================================================
st_mock = MagicMock()

# Baca secrets
secrets_data = toml.load(".streamlit/secrets.toml")
st_mock.secrets = secrets_data

# Mock cache_resource — harus support dua cara pakai:
#   1. @st.cache_resource          (tanpa kurung)
#   2. @st.cache_resource()        (pakai kurung)
def cache_resource_mock(func=None, **kwargs):
    """Mock yang support pemakaian dengan atau tanpa kurung."""
    if func is not None and callable(func):
        # Dipakai sebagai @st.cache_resource (langsung)
        return func
    # Dipakai sebagai @st.cache_resource() — return decorator
    def decorator(f):
        return f
    return decorator

st_mock.cache_resource = cache_resource_mock

sys.modules['streamlit'] = st_mock

# ================================================================
# IMPORT & TEST
# ================================================================
from db_storage import (
    load_users, load_database, load_config, load_log,
    get_statistik_db, test_koneksi,
)

print("=" * 60)
print("TEST DB_STORAGE")
print("=" * 60)

# Test koneksi
ok, msg = test_koneksi()
print(f"\n🔌 Koneksi: {'✅' if ok else '❌'} {msg}")

# Statistik
print("\n📊 Statistik Database:")
stats = get_statistik_db()
for k, v in stats.items():
    print(f"  - {k:20s} : {v} dokumen")

# Users
print("\n👥 Users:")
for u, data in load_users().items():
    print(f"  - {u:15s} | {data['nama']:25s} | {data['role']}")

# Config
print("\n⚙️ Config:")
print(f"  {load_config()}")

# Siswa
print(f"\n📚 Siswa: {len(load_database())} data")

# Log
print(f"📋 Log: {len(load_log())} entries")

print("\n✅ Selesai!")