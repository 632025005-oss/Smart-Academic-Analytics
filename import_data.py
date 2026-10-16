# import_data.py
# ================================================================
# Import Data Thesis dari Excel ke MongoDB
# Jalankan: python import_data.py
# ================================================================
import os
import toml
import pandas as pd
from datetime import datetime
from pymongo import MongoClient
import certifi  

# ================================================================
# KONFIGURASI
# ================================================================
EXCEL_FILE = "DATA THESIS.xlsx"
SHEET_NAME = "Form Responses 1"

# Mapping konstruk → indeks kolom (0-based)
# A=0, B=1, ..., K=10, L=11, dst
MAPPING = {
    "Self-Efficacy":      {"cols": [32, 33, 34, 35], "reverse": [36]},  # AG,AH,AI,AJ + AK(reverse)
    "Motivasi":           {"cols": [16, 17, 18, 19], "reverse": [20, 21]},  # Q,R,S,T + U,V(reverse)
    "Dukungan Sekolah":   {"cols": [27, 28, 30, 31], "reverse": []},  # AB,AC,AE,AF
    "Harapan Ortu":       {"cols": [25, 26], "reverse": []},  # Z,AA
    "Fasilitas":          {"cols": [29], "reverse": []},  # AD
    "Kecemasan":          {"cols": [23, 24], "reverse": []},  # X,Y
    "Keterlibatan Ortu":  {"cols": [10, 11, 12, 13, 14, 15], "reverse": []},  # K,L,M,N,O,P
    "Kemalasan":          {"cols": [22], "reverse": []},  # W
}

RATA_RATA_NILAI = 83.78

def hitung_konstruk(row):
    """Hitung 8 konstruk dari 1 baris data."""
    hasil = {}
    for nama, cfg in MAPPING.items():
        nilai_positif = [float(row.iloc[i]) for i in cfg["cols"]]
        nilai_reverse = [6.0 - float(row.iloc[i]) for i in cfg["reverse"]]
        semua = nilai_positif + nilai_reverse
        hasil[nama] = round(sum(semua) / len(semua), 3) if semua else 0
    return hasil

def kategori_nilai(n):
    if n >= 88: return "Sangat Baik"
    if n >= 84: return "Baik"
    if n >= 80: return "Cukup"
    return "Perlu Perhatian"

# ================================================================
# MAIN
# ================================================================
print("=" * 60)
print("IMPORT DATA THESIS KE MONGODB")
print("=" * 60)

# Koneksi MongoDB
secrets = toml.load(".streamlit/secrets.toml")
client = MongoClient(
    secrets["MONGODB_URI"],
    serverSelectionTimeoutMS=30000,
    connectTimeoutMS=30000,
    socketTimeoutMS=30000,
        tlsCAFile=certifi.where(),
)
db = client.saa_smpn6
siswa_col = db.siswa

# Baca Excel
print(f"\n📂 Membaca file: {EXCEL_FILE}")
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
print(f"✅ Berhasil baca {len(df)} baris")

# Cek duplikat
existing = siswa_col.count_documents({})
if existing > 0:
    print(f"\n⚠️  Sudah ada {existing} siswa di MongoDB.")
    konfirmasi = input("   Hapus dan ganti dengan data baru? (y/n): ")
    if konfirmasi.lower() != "y":
        print("❌ Dibatalkan.")
        exit()
    siswa_col.delete_many({})
    print("✅ Data lama dihapus.")

# Proses setiap baris
print("\n📝 Memproses data...")
data_insert = []

for idx, row in df.iterrows():
    try:
        nilai = float(row.iloc[37])  # Kolom AL = RATA-RATA
        konstruk = hitung_konstruk(row)
        selisih = round(nilai - RATA_RATA_NILAI, 2)
        
        doc = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Nama": f"Siswa-{idx+1:03d}",
            "Kelas": "IX-A",
            "Absen": idx + 1,
            "Nilai Akademik": round(nilai, 2),
            "Selisih": selisih,
            "Kategori": kategori_nilai(nilai),
            # Demografi
            "Usia": int(row.iloc[0]) if pd.notna(row.iloc[0]) else None,
            "Pendidikan_Ayah": int(row.iloc[1]) if pd.notna(row.iloc[1]) else None,
            "Pendidikan_Ibu": int(row.iloc[2]) if pd.notna(row.iloc[2]) else None,
            "Pendapatan_Ortu": int(row.iloc[3]) if pd.notna(row.iloc[3]) else None,
            "Jarak_Sekolah": float(row.iloc[4]) if pd.notna(row.iloc[4]) else None,
            "Fasilitas_Rumah": int(row.iloc[5]) if pd.notna(row.iloc[5]) else None,
            "Jam_Gadget": float(row.iloc[6]) if pd.notna(row.iloc[6]) else None,
            "Jam_Belajar": float(row.iloc[7]) if pd.notna(row.iloc[7]) else None,
            "Jam_Les": float(row.iloc[8]) if pd.notna(row.iloc[8]) else None,
            "Ekstrakurikuler": int(row.iloc[9]) if pd.notna(row.iloc[9]) else None,
            # 8 Konstruk
            "Self-Efficacy":      konstruk["Self-Efficacy"],
            "Motivasi":           konstruk["Motivasi"],
            "Dukungan Sekolah":   konstruk["Dukungan Sekolah"],
            "Harapan Ortu":       konstruk["Harapan Ortu"],
            "Fasilitas":          konstruk["Fasilitas"],
            "Kecemasan":          konstruk["Kecemasan"],
            "Keterlibatan Ortu":  konstruk["Keterlibatan Ortu"],
            "Kemalasan":          konstruk["Kemalasan"],
            "Dicatat Oleh": "import_data.py",
        }
        data_insert.append(doc)
        
    except Exception as e:
        print(f"   ⚠️  Baris {idx+1} error: {e}")

# Insert ke MongoDB
print(f"\n📤 Mengirim {len(data_insert)} siswa ke MongoDB...")
if data_insert:
    result = siswa_col.insert_many(data_insert)
    print(f"✅ Berhasil insert {len(result.inserted_ids)} dokumen")

# Ringkasan
print("\n" + "=" * 60)
print("✅ IMPORT SELESAI")
print("=" * 60)
print(f"\nTotal siswa di MongoDB: {siswa_col.count_documents({})}")
print(f"\nContoh 3 siswa pertama:")
for s in siswa_col.find().limit(3):
    print(f"  • {s['Nama']} | Nilai: {s['Nilai Akademik']} | Self-Efficacy: {s['Self-Efficacy']}")

print("\n🎉 Buka dashboard → Modul 02 untuk lihat data!")