import pandas as pd
from konstruk_config import KONSTRUK_MAP

# Baca CSV
df = pd.read_csv("Kuisioner Thesis.csv")
likert_cols = df.columns[-27:].tolist()
df_likert = df[likert_cols].apply(pd.to_numeric, errors='coerce')

print("=== RATA-RATA KONSTRUK ===")
for nama, cfg in KONSTRUK_MAP.items():
    idxs = [i - 1 for i in cfg["items"]]
    sub = df_likert.iloc[:, idxs].copy()
    
    # Reverse item (misal 5 jadi 1)
    for r_idx in cfg["reverse"]:
        col_pos = cfg["items"].index(r_idx)
        sub.iloc[:, col_pos] = 6 - sub.iloc[:, col_pos]
        
    skor = sub.mean(axis=1)
    print(f"{nama:25} : {skor.mean():.3f} (Tesis: cek Tabel 4.2)")