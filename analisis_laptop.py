# analisis_laptop.py
# ================================================================
# Script Analisis di Laptop
# Ambil data dari MongoDB → Analisis → Kirim hasil ke MongoDB
# Jalankan: python analisis_laptop.py
# ================================================================
import os
import toml
import pandas as pd
import numpy as np
from datetime import datetime
from pymongo import MongoClient
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import cross_val_score
from sklearn.metrics import mean_absolute_error

# ================================================================
# 1. KONEKSI MONGODB
# ================================================================
print("=" * 60)
print("ANALISIS DATA SISWA - SAA SMPN 6 SALATIGA")
print("=" * 60)

secrets = toml.load(".streamlit/secrets.toml")
client = MongoClient(secrets["MONGODB_URI"])
db = client.saa_smpn6

# ================================================================
# 2. AMBIL DATA DARI MONGODB
# ================================================================
print("\n📥 Mengambil data dari MongoDB...")
siswa_list = list(db.siswa.find())

if not siswa_list:
    print("❌ Tidak ada data siswa. Input dulu di dashboard (Modul 01).")
    exit()

df = pd.DataFrame(siswa_list).drop(columns=["_id"], errors="ignore")
print(f"✅ Berhasil ambil {len(df)} siswa")

# Kolom 8 faktor
faktor_cols = ["Self-Efficacy", "Keterlibatan Ortu", "Harapan Ortu",
               "Dukungan Sekolah", "Motivasi", "Kecemasan",
               "Fasilitas", "Kemalasan"]

# ================================================================
# 3. ANALISIS
# ================================================================
print("\n🔬 Menganalisis...\n")

# --- Analisis 1: Statistik Deskriptif ---
print("📊 [1/4] Statistik deskriptif...")
statistik = df[["Nilai Akademik"] + faktor_cols].describe().round(3).reset_index()
statistik.columns = ["Statistik"] + ["Nilai Akademik"] + faktor_cols

# --- Analisis 2: Statistik per Kelas ---
print("📊 [2/4] Statistik per kelas...")
per_kelas = df.groupby("Kelas").agg(
    Jumlah_Siswa=("Nama", "count"),
    Rata_rata=("Nilai Akademik", "mean"),
    Tertinggi=("Nilai Akademik", "max"),
    Terendah=("Nilai Akademik", "min"),
    Std_Dev=("Nilai Akademik", "std"),
).round(2).reset_index()

# --- Analisis 3: Ranking Siswa ---
print("🏆 [3/4] Ranking siswa...")
ranking = df[["Nama", "Kelas", "Absen", "Nilai Akademik", "Kategori"]].copy()
ranking = ranking.sort_values("Nilai Akademik", ascending=False).reset_index(drop=True)
ranking.insert(0, "Peringkat", range(1, len(ranking) + 1))

# --- Analisis 4: Random Forest + Feature Importance ---
print("🌲 [4/4] Random Forest + Feature Importance...")

X = df[faktor_cols].values
y = df["Nilai Akademik"].values

rf = RandomForestRegressor(n_estimators=100, max_depth=5, random_state=42)

# Cross-validation R² (kalau data cukup)
if len(df) >= 5:
    cv_scores = cross_val_score(rf, X, y, cv=min(5, len(df)), scoring="r2")
    r2_mean = cv_scores.mean().round(4)
    r2_std = cv_scores.std().round(4)
else:
    r2_mean = r2_std = None

# Fit full model
rf.fit(X, y)
y_pred = rf.predict(X)
mae = round(mean_absolute_error(y, y_pred), 4)

# Feature importance
feature_imp = pd.DataFrame({
    "Faktor": faktor_cols,
    "Importance": rf.feature_importances_.round(4)
}).sort_values("Importance", ascending=False).reset_index(drop=True)
feature_imp.insert(0, "Peringkat", range(1, len(feature_imp) + 1))

# Info model
info_model = pd.DataFrame([{
    "Metode": "Random Forest",
    "n_estimators": 100,
    "max_depth": 5,
    "Jumlah_Sampel": len(df),
    "R2_CV_Mean": r2_mean,
    "R2_CV_Std": r2_std,
    "MAE_Train": mae,
}])

print(f"   ✅ R² CV: {r2_mean} ± {r2_std}")
print(f"   ✅ MAE Train: {mae}")

# ================================================================
# 4. KIRIM HASIL KE MONGODB
# ================================================================
print("\n📤 Mengirim hasil ke MongoDB...\n")

def kirim(judul, metode, deskripsi, dataframe):
    db.hasil_analisis.insert_one({
        "judul": judul,
        "metode": metode,
        "deskripsi": deskripsi,
        "kolom": list(dataframe.columns),
        "jumlah_baris": len(dataframe),
        "data": dataframe.to_dict(orient="records"),
        "dibuat_oleh": "analisis_laptop.py",
        "tanggal": datetime.now(),
    })
    print(f"   ✅ {judul}")

kirim(
    "Statistik Deskriptif",
    "Statistik Deskriptif",
    f"Ringkasan statistik 8 faktor + nilai akademik per {datetime.now().strftime('%d %B %Y')}",
    statistik,
)

kirim(
    "Statistik per Kelas",
    "Group By Kelas",
    f"Perbandingan nilai antar kelas ({len(per_kelas)} kelas)",
    per_kelas,
)

kirim(
    f"Ranking Siswa ({len(ranking)} siswa)",
    "Sorting",
    "Peringkat siswa berdasarkan nilai akademik",
    ranking,
)

kirim(
    "Feature Importance (Random Forest)",
    "Random Forest",
    "Kontribusi prediktif setiap faktor terhadap nilai akademik",
    feature_imp,
)

kirim(
    "Info Model Random Forest",
    "Random Forest",
    "Hyperparameter dan performa model prediktif",
    info_model,
)

# ================================================================
# 5. SIMPAN CSV LOKAL (BACKUP)
# ================================================================
os.makedirs("hasil_analisis", exist_ok=True)
statistik.to_csv("hasil_analisis/statistik.csv", index=False)
per_kelas.to_csv("hasil_analisis/per_kelas.csv", index=False)
ranking.to_csv("hasil_analisis/ranking.csv", index=False)
feature_imp.to_csv("hasil_analisis/feature_importance.csv", index=False)

print(f"\n💾 CSV backup disimpan di folder 'hasil_analisis/'")
print("\n" + "=" * 60)
print("🎉 SELESAI! Buka dashboard → Modul 04 untuk lihat hasilnya.")
print("=" * 60)