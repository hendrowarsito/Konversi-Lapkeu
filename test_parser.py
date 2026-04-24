"""
Test unit untuk komponen inti parser (tanpa perlu jalankan Streamlit).
Jalankan: python test_parser.py
"""
import sys
sys.path.insert(0, '/home/claude')

# Import functions tanpa trigger streamlit UI
import importlib.util
spec = importlib.util.spec_from_file_location("app_module", "/home/claude/app.py")

# Patch streamlit agar tidak error saat import
import types
mock_st = types.ModuleType('streamlit')
mock_st.set_page_config = lambda **kwargs: None
mock_st.cache_resource = lambda f: f  # no-op decorator
mock_st.warning = lambda *a, **k: None
mock_st.info = lambda *a, **k: None
mock_st.error = lambda *a, **k: None
sys.modules['streamlit'] = mock_st

# Mock streamlit sub-modules
mock_components = types.ModuleType('streamlit.components')
mock_components_v1 = types.ModuleType('streamlit.components.v1')
mock_components_v1.html = lambda *a, **k: None
mock_components.v1 = mock_components_v1
sys.modules['streamlit.components'] = mock_components
sys.modules['streamlit.components.v1'] = mock_components_v1

# Column config juga di-mock
class MockColConfig:
    def __getattr__(self, name):
        return lambda *a, **k: None
mock_st.column_config = MockColConfig()

spec.loader.exec_module(sys.modules.setdefault('app_module', importlib.util.module_from_spec(spec)))
app = sys.modules['app_module']
spec.loader.exec_module(app)


def test(name, actual, expected):
    status = "✅" if actual == expected else "❌"
    print(f"{status} {name}")
    if actual != expected:
        print(f"   Expected: {expected!r}")
        print(f"   Got:      {actual!r}")
    return actual == expected


print("="*70)
print("TEST 1: parse_indonesian_number — kedua format pemisah ribuan")
print("="*70)
cases = [
    # Format titik (Indonesia)
    ("1.068.980.860.803",  1068980860803.0),
    ("982.621.700.996",    982621700996.0),
    ("(9.531.540.607)",    -9531540607.0),
    # Format koma (OCR barat / campuran) ← kasus dari screenshot
    ("375,272,610",        375272610.0),
    ("435,531,332",        435531332.0),
    ("(96,508,401)",       -96508401.0),
    ("360,724,480",        360724480.0),
    ("457,311,369",        457311369.0),
    # Format desimal campuran
    ("1.234,56",           1234.56),
    # Edge cases
    ("-",                  None),
    ("",                   None),
    ("Rp 1.000.000",       1000000.0),
]
passed = 0
for inp, exp in cases:
    result = app.parse_indonesian_number(inp)
    if test(f"'{inp}' → {exp}", result, exp):
        passed += 1
print(f"\nPassed: {passed}/{len(cases)}\n")


print("="*70)
print("TEST 2: _PAT_FINANCIAL_NUM — deteksi angka format titik DAN koma")
print("="*70)
import re as _re
pat = app._PAT_FINANCIAL_NUM
match_cases = [
    # harus match
    ("1.068.980.860.803",  True),
    ("375,272,610",        True),
    ("(96,508,401)",       True),
    ("457,311,369",        True),
    ("360,724,480",        True),
    # tidak boleh match (nomor catatan, tahun, dsb)
    ("4,33",               False),
    ("2017",               False),
    ("19a",                False),
    ("28",                 False),
]
for s, should_match in match_cases:
    found = bool(pat.search(s))
    status = "✅" if found == should_match else "❌"
    label = "match" if should_match else "no match"
    print(f"  {status} '{s}' → {label} (got: {'match' if found else 'no match'})")
print()


print("="*70)
print("TEST 3: LANGKAH 1 — detect_header_row")
print("="*70)
header_cases = [
    (["PT ABC", "LAPORAN", "Catatan  31 Desember 2022  31 Desember 2021"],
     ["2022", "2021"], 2),
    (["Tahun berakhir 2017 dan 2016"],
     ["2017", "2016"], 0),
    (["Tanggal 31 Desember 2022", "lanjutan..."],
     ["2022"], 0),
    (["Tidak ada tahun"],
     ["Tahun 1", "Tahun 2"], -1),
]
for lines, exp_years, exp_idx in header_cases:
    result = app.detect_header_row(lines)
    y_ok = sorted(result['years']) == sorted(exp_years)
    i_ok = (exp_idx == -1 and result['header_idx'] == -1) or \
           (exp_idx >= 0 and result['header_idx'] >= 0)
    status = "✅" if (y_ok and i_ok) else "❌"
    print(f"  {status} years={result['years']} idx={result['header_idx']}"
          f"  (exp years={exp_years})")
print()


print("="*70)
print("TEST 4: parse_ocr_text — simulasi kasus dari screenshot (format koma)")
print("="*70)
# Simulasi teks OCR dari screenshot Pak Hendro
# Format koma sebagai pemisah ribuan (bukan titik)
ocr_screenshot = """
LAPORAN LABA RUGI
Untuk tahun yang berakhir 31 Desember 2017 dan 2016

                                     Catatan    2017          2016
PENDAPATAN USAHA                        22    820,545,985   871,062,666
BEBAN POKOK PENDAPATAN                  23   (445,273,375) (435,531,334)
LABA BRUTO                                    375,272,610   435,531,332
PENDAPATAN (BEBAN)
LAIN-LAIN - NETO                        28     82,038,759    40,107,434
LABA SEBELUM ZAKAT                            457,311,369   475,638,766
LABA SEBELUM PAJAK
PENGHASILAN                                   457,232,881   475,203,810
BEBAN PAJAK PENGHASILAN                 29    (96,508,401) (108,791,214)
LABA BERSIH TAHUN BERJALAN                    360,724,480   366,412,596
PENGHASILAN KOMPREHENSIF
LAIN                                          (78,401)       (209,954)
TOTAL PENGHASILAN KOMPREHENSIF                360,646,079   366,202,642
"""

df = app.parse_ocr_text(ocr_screenshot, num_value_cols=2)
print(f"Total baris ter-ekstrak: {len(df)}")
print()
print(f"  {'Tipe':8s} {'Keterangan':50s} {'Nilai_1':>18} {'Nilai_2':>18}")
print(f"  {'─'*8} {'─'*50} {'─'*18} {'─'*18}")
for _, row in df.iterrows():
    v1 = f"{row['Nilai_1']:>18,.0f}" if row['Nilai_1'] is not None and str(row['Nilai_1']) != 'nan' else f"{'–':>18}"
    v2 = f"{row['Nilai_2']:>18,.0f}" if row['Nilai_2'] is not None and str(row['Nilai_2']) != 'nan' else f"{'–':>18}"
    print(f"  {str(row['Tipe']):8s} {str(row['Keterangan'])[:50]:50s} {v1} {v2}")

# Validasi nilai kunci
print("\n--- Validasi nilai kunci ---")
items = df[df['Tipe'] == 'item']
totals = df[df['Tipe'] == 'total']
all_data = df[df['Nilai_1'].notna()]

checks = [
    ("PENDAPATAN USAHA",          "Nilai_1", 820545985.0),
    ("BEBAN POKOK",               "Nilai_1", -445273375.0),
    ("LABA BRUTO",                "Nilai_1", 375272610.0),
    ("LABA BERSIH",               "Nilai_1", 360724480.0),
]
for kw, col, exp in checks:
    row = df[df['Keterangan'].str.contains(kw, case=False, na=False)]
    if len(row) > 0:
        val = row.iloc[0][col]
        ok = abs(val - exp) < 1 if val is not None else False
        print(f"  {'✅' if ok else '❌'} {kw}: {val:,.0f} (exp: {exp:,.0f})")
    else:
        print(f"  ⚠  '{kw}' tidak ditemukan di hasil")

# Pastikan tidak ada angka yang tertinggal di kolom Keterangan
print("\n--- Cek: tidak ada angka tertinggal di Keterangan ---")
has_number_in_label = df[df['Keterangan'].str.contains(
    r'\d{3},\d{3}|\d{3}\.\d{3}', regex=True, na=False
)]
if len(has_number_in_label) == 0:
    print("  ✅ Tidak ada angka keuangan yang tertinggal di kolom Keterangan")
else:
    print(f"  ❌ {len(has_number_in_label)} baris masih ada angka di Keterangan:")
    for _, r in has_number_in_label.iterrows():
        print(f"     '{r['Keterangan']}'")


print("\n" + "="*70)
print("TEST 5: _is_section_header — pastikan baris berAngka bukan section")
print("="*70)
sh_cases = [
    ("ASET",                                           True),
    ("PENDAPATAN (BEBAN)",                             True),
    ("PENGHASILAN KOMPREHENSIF",                       True),
    ("LABA BERSIH TAHUN BERJALAN 360,724,480",         False),  # ada angka → bukan
    ("LABAUSAHA OPERASIONAL 375,272,610 435,531,332",  False),  # ada angka → bukan
    ("Kas dan setara kas",                             False),
    ("",                                               False),
]
for label, exp in sh_cases:
    result = app._is_section_header(label)
    status = "✅" if result == exp else "❌"
    print(f"  {status} '{label[:55]}' → {result} (exp {exp})")


print("\n" + "="*70)
print("TEST 6: _clean_label")
print("="*70)
clean_cases = [
    ("Kas dan setara kas 4,33",                "Kas dan setara kas"),
    ("TOTAL ASET LANCAR",                      "TOTAL ASET LANCAR"),
    ("Aset keuangan lancar lainnya 12,14,33",  "Aset keuangan lancar lainnya"),
    ("LABA BERSIH TAHUN BERJALAN",             "LABA BERSIH TAHUN BERJALAN"),
]
for inp, exp in clean_cases:
    result = app._clean_label(inp)
    test(f"clean('{inp[:45]}') → '{exp}'", result, exp)


print("\n" + "="*70)
print("SELESAI")
print("="*70)



print("="*70)
print("TEST 2: LANGKAH 1 — detect_header_row")
print("="*70)
header_cases = [
    (["PT ABC", "LAPORAN POSISI KEUANGAN", "Tanggal 31 Desember 2022",
      "                      Catatan  31 Desember 2022  31 Desember 2021"],
     ["2022", "2021"], 2),
    (["Judul", "Tahun berakhir 31 Desember 2023 dan 2022"],
     ["2023", "2022"], 1),
    (["Tidak ada tahun di sini"],
     ["Tahun 1", "Tahun 2"], -1),
]
for lines, exp_years, exp_idx in header_cases:
    result = app.detect_header_row(lines)
    y_ok  = result['years'] == exp_years
    i_ok  = (exp_idx == -1 and result['header_idx'] == -1) or \
            (exp_idx >= 0 and result['header_idx'] >= 0)
    status = "✅" if (y_ok and i_ok) else "❌"
    print(f"{status} Years={result['years']} idx={result['header_idx']}")
print()


print("="*70)
print("TEST 3: LANGKAH 2 — filter_numeric_lines")
print("="*70)
lines_test = [
    "LAPORAN POSISI KEUANGAN",
    "31 Desember 2022  31 Desember 2021",    # ← header (idx 1)
    "ASET",
    "ASET LANCAR",
    "Kas dan setara kas  4,33  1.068.980.860.803  982.621.700.996",
    "Piutang usaha - pihak ketiga",
    "  (setelah dikurangi penyisihan)",
    "  kerugian ekspektasian)  5,14  142.101.369.697  126.714.548.456",
    "TOTAL ASET LANCAR    1.730.514.103.614  1.641.297.422.821",
]
numeric = app.filter_numeric_lines(lines_test, header_idx=1)
print(f"  Baris berAngka: {len(numeric)} (expected 3)")
for idx, line, nums in numeric:
    print(f"    baris {idx}: {nums}")
test("Filter: 3 baris berAngka", len(numeric), 3)
print()


print("="*70)
print("TEST 4: LANGKAH 3 — reconstruct_labels (KASUS MULTI-BARIS)")
print("="*70)
multiline_text = """
PT NIRVANA WASTU PRATAMA DAN ENTITAS ANAKNYA
LAPORAN POSISI KEUANGAN KONSOLIDASIAN
31 Desember 2022  31 Desember 2021

ASET
ASET LANCAR
Kas dan setara kas                   4,33    1.068.980.860.803    982.621.700.996
Piutang usaha - pihak ketiga
  (setelah dikurangi penyisihan
   kerugian kredit ekspektasian)     5,14      142.101.369.697    126.714.548.456
Piutang lain-lain
  Pihak berelasi                      32        1.351.523.984      1.145.603.985
  Pihak ketiga - neto               1c,5d       9.531.540.607      7.665.477.517
Bagian lancar biaya
  dibayar di muka                      8       13.720.855.958      9.730.774.850
TOTAL ASET LANCAR                          1.730.514.103.614  1.641.297.422.821
"""

df = app.parse_ocr_text(multiline_text, num_value_cols=2)
print(f"\nHasil parse ({len(df)} baris):")
for _, row in df.iterrows():
    v1 = f"{row['Nilai_1']:,.0f}" if row['Nilai_1'] else "-"
    v2 = f"{row['Nilai_2']:,.0f}" if row['Nilai_2'] else "-"
    print(f"  [{row['Tipe']:7s}] {str(row['Keterangan'])[:55]:55s} | {v1:>22} | {v2:>22}")

# Validasi kasus kritis
print("\n--- Validasi kasus multi-baris ---")
items = df[df['Tipe'] == 'item']
# Cek "Piutang usaha" dengan label multi-baris sudah tergabung
piutang = items[items['Keterangan'].str.contains('iutang usaha', na=False)]
if len(piutang) > 0:
    label = piutang.iloc[0]['Keterangan']
    val   = piutang.iloc[0]['Nilai_1']
    print(f"  Piutang usaha label: '{label}'")
    print(f"  Piutang usaha nilai: {val:,.0f}")
    test("Piutang usaha nilai 2022 benar", val, 142101369697.0)
    # Label seharusnya TIDAK hanya "kerugian kredit ekspektasian)"
    label_ok = 'iutang' in label.lower()
    test("Label multi-baris tergabung dengan benar", label_ok, True)

# Cek "Bagian lancar biaya dibayar di muka"
bagian = items[items['Keterangan'].str.contains('iaya', na=False)]
if len(bagian) > 0:
    label = bagian.iloc[0]['Keterangan']
    val   = bagian.iloc[0]['Nilai_1']
    print(f"\n  Biaya dibayar label: '{label}'")
    print(f"  Biaya dibayar nilai: {val:,.0f}")
    test("Biaya dibayar nilai benar", val, 13720855958.0)

# Cek TOTAL
totals = df[df['Tipe'] == 'total']
if len(totals) > 0:
    total_val = totals.iloc[0]['Nilai_1']
    test("TOTAL ASET LANCAR 2022 benar", total_val, 1730514103614.0)

# Cek section header TIDAK ikut terbawa sebagai bagian label item
sections = df[df['Tipe'] == 'section']
print(f"\n  Section headers terdeteksi: {list(sections['Keterangan'])}")
kas_row = items[items['Keterangan'].str.contains('Kas dan setara', na=False)]
if len(kas_row) > 0:
    label = kas_row.iloc[0]['Keterangan']
    aset_in_label = 'ASET' in label
    test(f"'ASET LANCAR' tidak ikut ke label 'Kas' (label='{label[:40]}')",
         aset_in_label, False)


print("\n" + "="*70)
print("TEST 5: _is_section_header")
print("="*70)
sh_cases = [
    ("ASET", True),
    ("ASET LANCAR", True),
    ("TOTAL ASET", True),
    ("Kas dan setara kas", False),
    ("1.068.980.860.803", False),
    ("LIABILITAS JANGKA PENDEK", True),
    ("", False),
]
for label, exp in sh_cases:
    result = app._is_section_header(label)
    test(f"_is_section_header('{label}') == {exp}", result, exp)


print("\n" + "="*70)
print("TEST 6: _clean_label (hapus nomor catatan dari label)")
print("="*70)
clean_cases = [
    ("Kas dan setara kas 4,33", "Kas dan setara kas"),
    ("Persediaan 6", "Persediaan"),
    ("Piutang usaha - pihak ketiga (setelah dikurangi) 5,14", "Piutang usaha - pihak ketiga (setelah dikurangi)"),
    ("TOTAL ASET LANCAR", "TOTAL ASET LANCAR"),
    ("Aset keuangan lancar lainnya 12,14,33", "Aset keuangan lancar lainnya"),
]
for inp, exp in clean_cases:
    result = app._clean_label(inp)
    test(f"clean('{inp[:40]}') → '{exp[:40]}'", result, exp)


print("\n" + "="*70)
print("SELESAI — Semua test algoritma 3-langkah")
print("="*70)
