import pandas as pd
import numpy as np
from konstruk_config import KONSTRUK_MAP
from db_storage import insert_siswa
from datetime import datetime
import pytz

def get_wib_time():
    return datetime.now(pytz.timezone('Asia/Jakarta')).strftime("%Y-%m-%d %H:%M")

# 1. Baca CSV
df = pd.read_csv("Kuisioner Thesis.csv")
# Hapus duplikat (ZUWANTYAN) dan data iseng (Hohoho)
df = df[df['Nama'] != 'Hohoho']
df = df.drop_duplicates(subset=['Nama'], keep='first')
df = df.reset_index(drop=True)

likert_cols = df.columns[-27:].tolist()
df_likert = df[likert_cols].apply(pd.to_numeric, errors='coerce')

print(f"Memproses {len(df)} siswa...")

for idx, row in df.iterrows():
    # Hitung 8 Konstruk
    profil = {}
    for nama, cfg in KONSTRUK_MAP.items():
        idxs = [i - 1 for i in cfg["items"]]
        sub = df_likert.iloc[idx, idxs].copy()
        
        # Reverse item
        for r_idx in cfg["reverse"]:
            col_pos = cfg["items"].index(r_idx)
            sub.iloc[col_pos] = 6 - sub.iloc[col_pos]
            
        profil[nama] = float(sub.mean())
    
    # 2. BUAT NILAI DUMMY (Karena Dapodik belum ada)
    # Rumus dummy: Nilai dasar 70 + pengaruh Self-Efficacy, Ortu, dan Motivasi - Kemalasan
    nilai_dummy = (
        70.0 
        + (profil["Self-Efficacy Akademik"] * 2.0) 
        + (profil["Keterlibatan Orang Tua"] * 1.5) 
        + (profil["Motivasi Belajar"] * 1.0)
        - (profil["Kemalasan Belajar"] * 1.5)
        + np.random.normal(0, 2) # Tambah noise sedikit agar realistis
    )
    nilai_dummy = min(max(nilai_dummy, 60.0), 98.0) # Batasi 60-98
    
    # 3. Simpan ke Database MongoDB
    insert_siswa({
        "Timestamp": get_wib_time(),
        "Nama": row['Nama'],
        "Kelas": row['Kelas'],
        "Absen": idx + 1,
        "Nilai Akademik": round(nilai_dummy, 2),
        "Kategori": "Baik" if nilai_dummy >= 80 else "Cukup",
        "Self-Efficacy": profil["Self-Efficacy Akademik"],
        "Keterlibatan Ortu": profil["Keterlibatan Orang Tua"],
        "Harapan Ortu": profil["Harapan Orang Tua"],
        "Dukungan Sekolah": profil["Dukungan Sekolah"],
        "Motivasi": profil["Motivasi Belajar"],
        "Kecemasan": profil["Kecemasan Akademik"],
        "Fasilitas": profil["Fasilitas Sekolah"],
        "Kemalasan": profil["Kemalasan Belajar"],
        "Dicatat Oleh": "Sistem NLP (Inject)"
    })

print(f"✅ Berhasil memasukkan {len(df)} data siswa ke MongoDB!")