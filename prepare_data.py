# prepare_data.py
import pandas as pd
import numpy as np
from konstruk_config import KONSTRUK_MAP

def process_survey(csv_path="Kuisioner Thesis.csv"):
    df = pd.read_csv(csv_path)
    
    # 1. Ambil 27 kolom Likert terakhir
    likert_cols = df.columns[-27:].tolist()
    df_likert = df[likert_cols].apply(pd.to_numeric, errors='coerce')
    
    # 2. Bentuk 8 Konstruk
    df_konstruk = pd.DataFrame()
    df_konstruk['Nama'] = df['Nama']
    df_konstruk['Kelas'] = df['Kelas']
    
    for nama, cfg in KONSTRUK_MAP.items():
        idxs = [i - 1 for i in cfg["items"]] # 0-based
        sub = df_likert.iloc[:, idxs].copy()
        
        # Reverse item jika perlu
        for r_idx in cfg["reverse"]:
            col_pos = cfg["items"].index(r_idx)
            sub.iloc[:, col_pos] = 6 - sub.iloc[:, col_pos]
            
        df_konstruk[nama] = sub.mean(axis=1)
        
    # 3. Ekstraksi Fitur Non-Likert (Demografi & Kebiasaan)
    # Kita ubah teks menjadi angka sederhana (bisa disesuaikan nanti)
    # Contoh: Jam Gadget, Jam Belajar (kategorikal ke numerik)
    def parse_jam(val):
        if pd.isna(val): return 0
        val = str(val).lower()
        if 'kurang dari 1' in val: return 0.5
        if '1 - 2' in val or '1  - 2' in val: return 1.5
        if '2 - 3' in val: return 2.5
        if '3 - 4' in val: return 3.5
        if 'lebih dari 4' in val: return 4.5
        return 0

    df_konstruk['Jam_Gadget'] = df['Rata-rata berapa jam per hari kamu menggunakan gadget/HP di luar keperluan belajar (bermain game, media sosial, menonton video)?'].apply(parse_jam)
    df_konstruk['Jam_Belajar'] = df['Rata-rata berapa jam kamu BELAJAR DI RUMAH setiap hari (di luar jam sekolah)?'].apply(parse_jam)
    
    # 4. (NANTI) Merge dengan Nilai Dapodik berdasarkan NISN
    # Untuk sekarang, kita buat Y dummy karena data Dapodik belum ada
    # df_konstruk['Nilai_Akademik'] = ...
    
    return df_konstruk

# Test jalankan
if __name__ == "__main__":
    df_clean = process_survey()
    print(df_clean.head())
    print(df_clean.describe())