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
print("TEST 6: End-to-end — 2 halaman neraca asli (2024/2023 dan 2023/2022)")
print("=" * 70)

# Replika teks halaman 1 laporan asli (layout: label | 2024 | Catatan | 2023)
PAGE_2024 = """
                                        2024      Catatan/      2023
ASET

Aset lancar
Kas di bank                          3,640,682        5       5,682,551
Piutang usaha                       63,679,413        6      35,748,619
Persediaan                          21,234,913        7      14,315,782
Aset lancar lain-lain                  350,598                   22,120

                                    88,905,606               55,769,072
Aset tidak lancar
Aset tetap                             116,888                  158,259
Aset hak-guna                          631,617                  247,987
Aset pajak tangguhan                    84,842                   53,408
Aset tidak lancar lain-lain            177,864                  248,213

                                     1,011,211                  707,867

JUMLAH ASET                         89,916,817               56,476,939

LIABILITAS

Liabilitas jangka pendek
Utang usaha                         47,917,492        8      22,537,498
Utang lain-lain                        168,277                  173,241
Akrual                               1,120,405                  806,487
Utang pajak:                                          9a
- Pajak penghasilan badan            1,122,963                  887,302
- Pajak lain-lain                      864,845                  534,816
Bagian jangka pendek
   dari liabilitas sewa                287,348                  201,034

                                    51,481,330               25,140,378
Liabilitas jangka panjang
Liabilitas sewa                        263,622                        -
Kewajiban imbalan kerja                385,645                  242,765

                                       649,267                  242,765

JUMLAH LIABILITAS                   52,130,597               25,383,143

EKUITAS
Modal saham - 255.750 lembar
   modal dasar, ditempatkan dan
   disetor penuh dengan nilai
   nominal Rp 100.000
   (satuan penuh) per lembar
   saham                            25,575,000       10      25,575,000
Saldo laba                          12,211,220                5,518,796

JUMLAH EKUITAS                      37,786,220               31,093,796

JUMLAH LIABILITAS
   DAN EKUITAS                      89,916,817               56,476,939
"""

# Replika halaman 2 (2023/2022) — perhatikan perbedaan dari halaman 1:
# 'Aset lancar lainnya' (bukan 'lain-lain'), ada 'Pajak dibayar di muka'
# dengan nilai '-' di 2023, dan 'Liabilitas sewa bagian jangka pendek'
# (urutan kata beda dengan 'Bagian jangka pendek dari liabilitas sewa')
PAGE_2023 = """
                                        2023      Catatan/      2022
ASET
Aset lancar
Kas di bank                          5,682,551        5       1,477,996
Piutang usaha                       35,748,619        6      19,230,251
Persediaan                          14,315,782        7       9,812,237
Pajak dibayar di muka                        -       9a           1,544
Aset lancar lainnya                     22,120                   68,598

                                    55,769,072               30,590,626
Aset tidak lancar
Aset tetap                             158,259                  123,208
Aset hak-guna                          247,987                  316,404
Aset pajak tangguhan                    53,408                   28,152
Aset tidak lancar lainnya              248,213                   21,665

                                       707,867                  489,429

JUMLAH ASET                         56,476,939               31,080,055

LIABILITAS

Liabilitas jangka pendek
Utang usaha                         22,537,498        8      20,387,156
Akrual                                 806,487                  418,418
Utang pajak:                                          9b
- Pajak penghasilan badan              887,302                  995,054
- Pajak lain-lain                      534,816                   79,779
Liabilitas sewa
   bagian jangka pendek                201,034                  136,659
Utang lain-lain                        173,241                  123,634

                                    25,140,378               22,140,700
Liabilitas jangka panjang
Liabilitas sewa                              -                  214,914
Kewajiban imbalan kerja                242,765                  127,964

                                       242,765                  342,878

JUMLAH LIABILITAS                   25,383,143               22,483,578

EKUITAS
Modal saham - 255.750 lembar
   modal dasar, ditempatkan dan
   disetor penuh dengan nilai
   nominal Rp 100.000
   (satuan penuh) per lembar
   saham (2022: 50.000 lembar
   modal dasar, ditempatkan dan
   disetor penuh dengan nilai
   nominal Rp 100.000 (satuan
   penuh) per lembar saham)         25,575,000       10       5,000,000
Saldo laba                           5,518,796                3,596,477

JUMLAH EKUITAS                      31,093,796                8,596,477

JUMLAH LIABILITAS DAN
   EKUITAS                          56,476,939               31,080,055
"""

df_p1 = app.parse_ocr_text(PAGE_2024, num_value_cols=2)
df_p2 = app.parse_ocr_text(PAGE_2023, num_value_cols=2)

merged_e2e = app._merge_reports_by_account([
    {"df": df_p1, "years": ["2024", "2023"]},
    {"df": df_p2, "years": ["2023", "2022"]},
])
df_e2e = app._fill_missing_values_with_zero(merged_e2e["df"])
print(f"Tahun: {merged_e2e['years']}")
print(df_e2e.to_string(index=False))

check("Tahun gabungan = 2024,2023,2022",
      merged_e2e["years"] == ["2024", "2023", "2022"], str(merged_e2e["years"]))

def val(df, label_contains, col):
    # Exact match dulu (hindari 'Utang usaha' tertangkap 'Piutang usaha'),
    # baru fallback ke substring
    rows = df[df["Keterangan"].str.strip() == label_contains]
    if len(rows) == 0:
        rows = df[df["Keterangan"].str.contains(label_contains, case=False, regex=False, na=False)]
    if len(rows) == 0:
        return None
    return rows.iloc[0][col]

def row_count(df, label_contains):
    return len(df[df["Keterangan"].str.contains(label_contains, case=False, regex=False, na=False)])

# Nilai kunci vs laporan asli
e2e_checks = [
    ("Kas di bank",            3_640_682, 5_682_551, 1_477_996),
    ("Piutang usaha",         63_679_413, 35_748_619, 19_230_251),
    ("Persediaan",            21_234_913, 14_315_782, 9_812_237),
    ("Aset tetap",               116_888,    158_259,   123_208),
    ("Aset hak-guna",            631_617,    247_987,   316_404),
    ("Aset pajak tangguhan",      84_842,     53_408,    28_152),
    ("Utang usaha",           47_917_492, 22_537_498, 20_387_156),
    ("Akrual",                 1_120_405,    806_487,   418_418),
    ("Pajak penghasilan badan",1_122_963,    887_302,   995_054),
    ("Pajak lain-lain",          864_845,    534_816,    79_779),
    ("Kewajiban imbalan kerja",  385_645,    242_765,   127_964),
    ("Saldo laba",            12_211_220,  5_518_796, 3_596_477),
]
for label, v24, v23, v22 in e2e_checks:
    got = (val(df_e2e, label, "Nilai_1"), val(df_e2e, label, "Nilai_2"), val(df_e2e, label, "Nilai_3"))
    check(f"{label}: {v24:,} | {v23:,} | {v22:,}",
          got == (float(v24), float(v23), float(v22)), str(got))

# 'lain-lain' dan 'lainnya' tergabung jadi SATU akun
check("'Aset lancar lain-lain/lainnya' = 1 baris",
      row_count(df_e2e, "Aset lancar lain") == 1,
      f"count={row_count(df_e2e, 'Aset lancar lain')}")
check("Aset lancar lain-lain: 350.598 | 22.120 | 68.598",
      (val(df_e2e, "Aset lancar lain", "Nilai_1"),
       val(df_e2e, "Aset lancar lain", "Nilai_2"),
       val(df_e2e, "Aset lancar lain", "Nilai_3")) == (350_598.0, 22_120.0, 68_598.0))
check("'Aset tidak lancar lain-lain/lainnya' = 1 baris",
      row_count(df_e2e, "Aset tidak lancar lain") == 1)

# Nilai '-' (nihil) tidak menggeser kolom
check("Pajak dibayar di muka: 0 | 0 | 1.544 (tanda '-' = nihil)",
      (val(df_e2e, "Pajak dibayar di muka", "Nilai_1"),
       val(df_e2e, "Pajak dibayar di muka", "Nilai_2"),
       val(df_e2e, "Pajak dibayar di muka", "Nilai_3")) == (0.0, 0.0, 1_544.0),
      str((val(df_e2e, "Pajak dibayar di muka", "Nilai_1"),
           val(df_e2e, "Pajak dibayar di muka", "Nilai_2"),
           val(df_e2e, "Pajak dibayar di muka", "Nilai_3"))))
ls_rows = df_e2e[df_e2e["Keterangan"] == "Liabilitas sewa"]
check("'Liabilitas sewa' (jangka panjang): 263.622 | 0 | 214.914",
      len(ls_rows) == 1 and
      (ls_rows.iloc[0]["Nilai_1"], ls_rows.iloc[0]["Nilai_2"], ls_rows.iloc[0]["Nilai_3"])
      == (263_622.0, 0.0, 214_914.0),
      ls_rows.to_string())

# Urutan kata beda → tetap satu akun
n_sewa_pendek = row_count(df_e2e, "jangka pendek") - row_count(df_e2e, "Liabilitas jangka pendek")
check("'Bagian jangka pendek dari liabilitas sewa' ≡ "
      "'Liabilitas sewa bagian jangka pendek' (1 baris)",
      val(df_e2e, "Bagian jangka pendek", "Nilai_1") == 287_348.0 and
      val(df_e2e, "Bagian jangka pendek", "Nilai_2") == 201_034.0 and
      val(df_e2e, "Bagian jangka pendek", "Nilai_3") == 136_659.0,
      str((val(df_e2e, "Bagian jangka pendek", "Nilai_1"),
           val(df_e2e, "Bagian jangka pendek", "Nilai_2"),
           val(df_e2e, "Bagian jangka pendek", "Nilai_3"))))

# Section Title Case terdeteksi — tidak menempel ke item
check("'Aset lancar' jadi baris section",
      "section" in list(df_e2e[df_e2e["Keterangan"] == "Aset lancar"]["Tipe"]),
      str(df_e2e[df_e2e["Keterangan"].str.contains("Aset lancar", na=False)][["Tipe","Keterangan"]].values))
check("'Aset tidak lancar' section — TIDAK menempel ke 'Aset tetap'",
      row_count(df_e2e, "Aset tidak lancar Aset tetap") == 0)

# Label total terpecah dua baris tergabung
jle = df_e2e[df_e2e["Keterangan"].str.contains("JUMLAH LIABILITAS DAN", na=False)]
check("'JUMLAH LIABILITAS DAN EKUITAS' = 1 baris total",
      len(jle) == 1 and jle.iloc[0]["Tipe"] == "total",
      jle.to_string())
if len(jle) == 1:
    check("JUMLAH LIABILITAS DAN EKUITAS: 89.916.817 | 56.476.939 | 31.080.055",
          (jle.iloc[0]["Nilai_1"], jle.iloc[0]["Nilai_2"], jle.iloc[0]["Nilai_3"])
          == (89_916_817.0, 56_476_939.0, 31_080_055.0),
          str((jle.iloc[0]["Nilai_1"], jle.iloc[0]["Nilai_2"], jle.iloc[0]["Nilai_3"])))

# Narasi modal saham: angka '255.750 lembar' & 'Rp 100.000' BUKAN nilai akun
modal_rows = df_e2e[df_e2e["Keterangan"].str.contains("Modal saham", na=False)]
check("Baris modal saham = 1 (halaman 2024 & 2023 tergabung)",
      len(modal_rows) == 1, f"count={len(modal_rows)}")
if len(modal_rows) >= 1:
    mr = modal_rows.iloc[0]
    check("Modal saham: 25.575.000 | 25.575.000 | 5.000.000 (bukan 255.750/100.000)",
          (mr["Nilai_1"], mr["Nilai_2"], mr["Nilai_3"])
          == (25_575_000.0, 25_575_000.0, 5_000_000.0),
          str((mr["Nilai_1"], mr["Nilai_2"], mr["Nilai_3"])))

# Nomor catatan di tengah (layout 2024|Catatan|2023) tertangkap
check("Catatan 'Kas di bank' = 5",
      str(val(df_e2e.assign(), "Kas di bank", "Catatan")) == "5",
      str(val(df_e2e, "Kas di bank", "Catatan")))


print()
print("=" * 70)
if FAILED:
    print(f"❌ {len(FAILED)} TEST GAGAL:")
    for f in FAILED:
        print(f"   - {f}")
    sys.exit(1)
else:
    print("✅ SEMUA TEST LULUS")
