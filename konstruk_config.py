# konstruk_config.py
# -*- coding: utf-8 -*-

# Indeks 1-based sesuai urutan kolom Likert di CSV
KONSTRUK_MAP = {
    "Self-Efficacy Akademik": {"items": [23, 24, 25, 26, 27, 8], "reverse": [27]},
    "Motivasi Belajar":       {"items": [7, 9, 10, 11, 12, 17],  "reverse": [12]},
    "Dukungan Sekolah":       {"items": [18, 19, 21, 22],        "reverse": []},
    "Harapan Orang Tua":      {"items": [2, 16],                 "reverse": []},
    "Fasilitas Sekolah":      {"items": [20],                    "reverse": []},
    "Kecemasan Akademik":     {"items": [14, 15],                "reverse": []},
    "Keterlibatan Orang Tua": {"items": [1, 3, 4, 5, 6],         "reverse": []},
    "Kemalasan Belajar":      {"items": [13],                    "reverse": []},
}

# Daftar nama mapel untuk Dapodik (11 mapel)
MAPEL_LIST = [
    "Agama", "PPKn", "B.Indo", "Matematika", "IPA", 
    "IPS", "B.Inggris", "Seni", "PJOK", "Prakarya", "Mulok"
]