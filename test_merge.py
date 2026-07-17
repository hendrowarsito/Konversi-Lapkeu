"""
Test penggabungan laporan keuangan multi-tahun (union by account name).

Skenario utama (sesuai permintaan Pak Hendro):
- Neraca 2025 punya akun "Piutang usaha" DAN "Piutang lain-lain"
- Neraca 2024 hanya punya "Piutang usaha"
- Hasil gabungan harus menampilkan KEDUA akun, dan "Piutang lain-lain"
  pada kolom 2024 (dari dokumen 2024) diisi 0.

Jalankan: python test_merge.py
"""
import sys
import types
from pathlib import Path
import importlib.util

_APP_PATH = Path(__file__).resolve().parent / "app.py"

# Mock streamlit agar app.py bisa di-import tanpa UI
mock_st = types.ModuleType('streamlit')
mock_st.set_page_config = lambda **kwargs: None
mock_st.cache_resource = lambda f: f
mock_st.warning = lambda *a, **k: None
mock_st.info = lambda *a, **k: None
mock_st.error = lambda *a, **k: None
sys.modules['streamlit'] = mock_st

mock_components = types.ModuleType('streamlit.components')
mock_components_v1 = types.ModuleType('streamlit.components.v1')
mock_components_v1.html = lambda *a, **k: None
mock_components.v1 = mock_components_v1
sys.modules['streamlit.components'] = mock_components
sys.modules['streamlit.components.v1'] = mock_components_v1


class MockColConfig:
    def __getattr__(self, name):
        return lambda *a, **k: None
mock_st.column_config = MockColConfig()

spec = importlib.util.spec_from_file_location("app_module", str(_APP_PATH))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

import pandas as pd

FAILED = []


def check(name, cond, detail=""):
    status = "✅" if cond else "❌"
    print(f"{status} {name}" + (f"  [{detail}]" if detail and not cond else ""))
    if not cond:
        FAILED.append(name)


def make_df(rows):
    """rows: list of (tipe, keterangan, catatan, nilai1, nilai2)"""
    return pd.DataFrame([
        {"Tipe": t, "Keterangan": k, "Catatan": c, "Nilai_1": v1, "Nilai_2": v2}
        for (t, k, c, v1, v2) in rows
    ])


print("=" * 70)
print("TEST 1: Skenario Pak Hendro — akun beda antar tahun → union + isi 0")
print("=" * 70)

# Neraca dari laporan tahun 2025 (kolom: 2025, 2024)
neraca_2025 = {
    "df": make_df([
        ("section", "ASET",                "",    None,          None),
        ("section", "ASET LANCAR",         "",    None,          None),
        ("item",    "Kas dan setara kas",  "4",   1_000_000.0,   900_000.0),
        ("item",    "Piutang usaha",       "5",   500_000.0,     450_000.0),
        ("item",    "Piutang lain-lain",   "6",   120_000.0,     100_000.0),
        ("total",   "TOTAL ASET LANCAR",   "",    1_620_000.0,   1_450_000.0),
    ]),
    "years": ["2025", "2024"],
}

# Neraca dari laporan tahun 2024 (kolom: 2024, 2023) — TANPA piutang lain-lain,
# tapi ADA "Persediaan" yang tidak muncul di laporan 2025
neraca_2024 = {
    "df": make_df([
        ("section", "ASET",                "",    None,        None),
        ("section", "ASET LANCAR",         "",    None,        None),
        ("item",    "Kas dan setara kas",  "4",   900_000.0,   800_000.0),
        ("item",    "Piutang usaha",       "5",   450_000.0,   400_000.0),
        ("item",    "Persediaan",          "7",   200_000.0,   180_000.0),
        ("total",   "TOTAL ASET LANCAR",   "",    1_550_000.0, 1_380_000.0),
    ]),
    "years": ["2024", "2023"],
}

merged = app._merge_reports_by_account([neraca_2024, neraca_2025])  # urutan sengaja dibalik
df = merged["df"]
years = merged["years"]

print(f"\nTahun gabungan: {years}")
print(df.to_string(index=False))
print()

check("Tahun gabungan = ['2025','2024','2023'] (urut menurun)",
      years == ["2025", "2024", "2023"], str(years))

labels = list(df["Keterangan"])
check("'Piutang lain-lain' ada di hasil gabungan", "Piutang lain-lain" in labels)
check("'Persediaan' (hanya di dok. 2024) ada di hasil gabungan", "Persediaan" in labels)
check("Tidak ada duplikat 'Piutang usaha'",
      labels.count("Piutang usaha") == 1, str(labels))
check("Tidak ada duplikat section 'ASET LANCAR'",
      labels.count("ASET LANCAR") == 1, str(labels))

# Posisi: Persediaan harus di DALAM section (sebelum TOTAL), bukan di akhir tabel
idx_persediaan = labels.index("Persediaan")
idx_total = labels.index("TOTAL ASET LANCAR")
check("'Persediaan' disisipkan SEBELUM baris TOTAL (di dalam section-nya)",
      idx_persediaan < idx_total, f"persediaan={idx_persediaan}, total={idx_total}")

def get_row(label):
    return df[df["Keterangan"] == label].iloc[0]

pl = get_row("Piutang lain-lain")
check("Piutang lain-lain 2025 = 120.000", pl["Nilai_1"] == 120_000.0, str(pl["Nilai_1"]))
check("Piutang lain-lain 2024 = 100.000 (dari kolom pembanding dok. 2025)",
      pl["Nilai_2"] == 100_000.0, str(pl["Nilai_2"]))
check("Piutang lain-lain 2023 kosong (belum diisi 0)",
      pd.isna(pl["Nilai_3"]), str(pl["Nilai_3"]))

ps = get_row("Persediaan")
check("Persediaan 2025 kosong (belum diisi 0)", pd.isna(ps["Nilai_1"]), str(ps["Nilai_1"]))
check("Persediaan 2024 = 200.000", ps["Nilai_2"] == 200_000.0, str(ps["Nilai_2"]))
check("Persediaan 2023 = 180.000", ps["Nilai_3"] == 180_000.0, str(ps["Nilai_3"]))

# Overlap tahun 2024: nilai dari dokumen TERBARU (2025) yang menang
pu = get_row("Piutang usaha")
check("Piutang usaha 2024 = 450.000 (dokumen terbaru menang)",
      pu["Nilai_2"] == 450_000.0, str(pu["Nilai_2"]))


print()
print("=" * 70)
print("TEST 2: _fill_missing_values_with_zero — akun hilang → 0")
print("=" * 70)

df_filled = app._fill_missing_values_with_zero(df)
pl_f = df_filled[df_filled["Keterangan"] == "Piutang lain-lain"].iloc[0]
ps_f = df_filled[df_filled["Keterangan"] == "Persediaan"].iloc[0]
sec_f = df_filled[df_filled["Keterangan"] == "ASET LANCAR"].iloc[0]

check("Piutang lain-lain 2023 = 0 setelah fill", pl_f["Nilai_3"] == 0, str(pl_f["Nilai_3"]))
check("Persediaan 2025 = 0 setelah fill", ps_f["Nilai_1"] == 0, str(ps_f["Nilai_1"]))
check("Nilai yang sudah ada tidak berubah", pl_f["Nilai_1"] == 120_000.0, str(pl_f["Nilai_1"]))
check("Baris section TETAP kosong (tidak diisi 0)",
      pd.isna(sec_f["Nilai_1"]) and pd.isna(sec_f["Nilai_2"]) and pd.isna(sec_f["Nilai_3"]))


print()
print("=" * 70)
print("TEST 3: Normalisasi nama akun — variasi penulisan tetap tergabung")
print("=" * 70)

norm = app._normalize_keterangan
check("'Piutang lain-lain' == 'Piutang lain lain'",
      norm("Piutang lain-lain") == norm("Piutang lain lain"))
check("'AKTIVA LANCAR' == 'Aset Lancar' (sinonim)",
      norm("AKTIVA LANCAR") == norm("Aset Lancar"))
check("'Kewajiban jangka pendek' == 'Liabilitas jangka pendek'",
      norm("Kewajiban jangka pendek") == norm("Liabilitas jangka pendek"))
check("'Kas dan setara kas 2h' → nomor catatan dibuang",
      norm("Kas dan setara kas 2h") == norm("Kas dan setara kas"))
check("'Utang PPh pasal 21' != 'Utang PPh pasal 25' (angka bermakna dipertahankan)",
      norm("Utang PPh pasal 21") != norm("Utang PPh pasal 25"))


print()
print("=" * 70)
print("TEST 4: Akun bernama sama di section berbeda TIDAK tergabung")
print("=" * 70)

neraca_a = {
    "df": make_df([
        ("section", "ASET LANCAR",       "", None,       None),
        ("item",    "Lainnya",           "", 10_000.0,   9_000.0),
        ("total",   "TOTAL ASET LANCAR", "", 10_000.0,   9_000.0),
        ("section", "LIABILITAS",        "", None,       None),
        ("item",    "Lainnya",           "", 5_000.0,    4_000.0),
        ("total",   "TOTAL LIABILITAS",  "", 5_000.0,    4_000.0),
    ]),
    "years": ["2025", "2024"],
}
neraca_b = {
    "df": make_df([
        ("section", "ASET LANCAR",       "", None,       None),
        ("item",    "Lainnya",           "", 9_000.0,    8_000.0),
        ("total",   "TOTAL ASET LANCAR", "", 9_000.0,    8_000.0),
        ("section", "LIABILITAS",        "", None,       None),
        ("item",    "Lainnya",           "", 4_000.0,    3_000.0),
        ("total",   "TOTAL LIABILITAS",  "", 4_000.0,    3_000.0),
    ]),
    "years": ["2024", "2023"],
}

merged2 = app._merge_reports_by_account([neraca_a, neraca_b])
df2 = merged2["df"]
print(df2.to_string(index=False))
lainnya_rows = df2[df2["Keterangan"] == "Lainnya"]
check("Ada 2 baris 'Lainnya' (satu per section)", len(lainnya_rows) == 2,
      f"jumlah={len(lainnya_rows)}")
if len(lainnya_rows) == 2:
    r_aset = lainnya_rows.iloc[0]
    r_liab = lainnya_rows.iloc[1]
    check("'Lainnya' (aset) 2023 = 8.000 dari dok. B", r_aset["Nilai_3"] == 8_000.0,
          str(r_aset["Nilai_3"]))
    check("'Lainnya' (liabilitas) 2023 = 3.000 dari dok. B", r_liab["Nilai_3"] == 3_000.0,
          str(r_liab["Nilai_3"]))


print()
print("=" * 70)
print("TEST 5: Merge 3 dokumen sekaligus (2025, 2024, 2023 — masing-masing 2 kolom)")
print("=" * 70)

doc_2023 = {
    "df": make_df([
        ("section", "ASET LANCAR",        "", None,       None),
        ("item",    "Kas dan setara kas", "", 800_000.0,  700_000.0),
        ("item",    "Uang muka",          "", 50_000.0,   40_000.0),
        ("total",   "TOTAL ASET LANCAR",  "", 850_000.0,  740_000.0),
    ]),
    "years": ["2023", "2022"],
}
merged3 = app._merge_reports_by_account([doc_2023, neraca_2025, neraca_2024])
df3 = app._fill_missing_values_with_zero(merged3["df"])
print(f"Tahun: {merged3['years']}")
print(df3.to_string(index=False))

check("Tahun gabungan 4 kolom: 2025..2022",
      merged3["years"] == ["2025", "2024", "2023", "2022"], str(merged3["years"]))
um = df3[df3["Keterangan"] == "Uang muka"].iloc[0]
check("'Uang muka' 2025 = 0 dan 2024 = 0 (tidak ada di dok. lain)",
      um["Nilai_1"] == 0 and um["Nilai_2"] == 0, f"{um['Nilai_1']}, {um['Nilai_2']}")
check("'Uang muka' 2023 = 50.000", um["Nilai_3"] == 50_000.0, str(um["Nilai_3"]))
kas = df3[df3["Keterangan"] == "Kas dan setara kas"].iloc[0]
check("'Kas dan setara kas' terisi di semua 4 tahun",
      kas["Nilai_1"] == 1_000_000.0 and kas["Nilai_2"] == 900_000.0
      and kas["Nilai_3"] == 800_000.0 and kas["Nilai_4"] == 700_000.0)


print()
print("=" * 70)
if FAILED:
    print(f"❌ {len(FAILED)} TEST GAGAL:")
    for f in FAILED:
        print(f"   - {f}")
    sys.exit(1)
else:
    print("✅ SEMUA TEST LULUS")
