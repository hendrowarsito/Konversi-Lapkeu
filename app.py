"""
Konverter Laporan Keuangan - Gambar & PDF ke Excel
===================================================
Mendukung:
  • Upload gambar PNG/JPG (OCR lokal atau Claude Vision API)
  • Upload file PDF (scan maupun digital)
  • Mode PDF cerdas: scan halaman, temukan judul laporan keuangan,
    ekstrak tahun, ambil HANYA tahun yang diminta per file
  • Paste clipboard (Ctrl+V)
  • Export Excel multi-sheet dengan formula SUM

Author: Untuk Hendro Warsito / KJPP SRR
"""

import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import base64
import json
from pathlib import Path
from datetime import datetime
from PIL import Image
import streamlit.components.v1 as components

# PDF
try:
    from pypdf import PdfReader
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False

try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

# OCR engines
try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False

try:
    import easyocr
    EASYOCR_AVAILABLE = True
except ImportError:
    EASYOCR_AVAILABLE = False

# Excel
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =============================================================================
# KONSTANTA & KONFIGURASI
# =============================================================================

st.set_page_config(
    page_title="Konverter Laporan Keuangan",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Kata kunci untuk deteksi jenis laporan (bahasa Indonesia)
KEYWORDS_NERACA = [
    "laporan posisi keuangan", "neraca", "aset lancar", "aset tidak lancar",
    "total aset", "liabilitas", "ekuitas", "modal saham",
]
KEYWORDS_LABA_RUGI = [
    "laba rugi", "penghasilan komprehensif", "pendapatan", "beban pokok",
    "laba bruto", "laba usaha", "laba sebelum pajak", "laba tahun berjalan",
    "laba per saham",
]
KEYWORDS_ARUS_KAS = [
    "arus kas", "aktivitas operasi", "aktivitas investasi", "aktivitas pendanaan",
    "kas dan setara kas awal", "kas dan setara kas akhir",
]

# Kata kunci baris section/total (untuk styling)
SECTION_KEYWORDS = [
    "aset", "aset lancar", "aset tidak lancar", "liabilitas", "liabilitas jangka pendek",
    "liabilitas jangka panjang", "ekuitas",
    "pendapatan", "beban", "arus kas dari",
]
TOTAL_KEYWORDS = ["total", "jumlah", "laba bruto", "laba usaha", "laba sebelum",
                  "laba tahun", "kenaikan", "penurunan"]

# Judul halaman laporan keuangan yang dicari (lowercase)
TARGET_REPORT_TITLES = {
    "neraca":    ["laporan posisi keuangan", "neraca konsolidasian", "neraca"],
    "laba_rugi": ["laporan laba rugi", "laporan laba-rugi", "laba rugi komprehensif"],
    "arus_kas":  ["laporan arus kas"],
}


# =============================================================================
# MODUL 0: PDF PROCESSOR  (pendekatan: PDF → gambar per halaman)
# =============================================================================

def pdf_to_images(pdf_bytes: bytes, dpi: int = 200) -> list[Image.Image]:
    """
    Konversi semua halaman PDF menjadi list PIL.Image.

    Strategi:
    - Gunakan pdf2image (poppler) jika tersedia → kualitas terbaik
    - Fallback ke pypdf + PIL jika pdf2image tidak ada
    """
    images = []

    if PDF2IMAGE_AVAILABLE:
        try:
            imgs = convert_from_bytes(pdf_bytes, dpi=dpi)
            return imgs
        except Exception as e:
            st.warning(f"pdf2image gagal ({e}), mencoba fallback...")

    # Fallback: render via pypdf → cukup untuk PDF digital
    if PYPDF_AVAILABLE:
        try:
            from pypdf import PdfReader
            import struct, zlib

            reader = PdfReader(io.BytesIO(pdf_bytes))
            for page in reader.pages:
                # Ambil teks sebagai gambar teks putih di atas hitam (dummy)
                # Fallback ini rendah kualitasnya, hanya untuk PDF digital
                txt = page.extract_text() or ""
                # Buat gambar sederhana berisi teks
                from PIL import ImageDraw, ImageFont
                img = Image.new("RGB", (1200, 1600), color=(255, 255, 255))
                draw = ImageDraw.Draw(img)
                # Tulis teks dengan font kecil
                y = 20
                for line in txt.split("\n")[:80]:
                    draw.text((20, y), line[:120], fill=(0, 0, 0))
                    y += 18
                images.append(img)
            return images
        except Exception as e:
            st.error(f"Fallback PDF render gagal: {e}")

    return images


def pdf_extract_page_text(pdf_bytes: bytes) -> list[str]:
    """
    Ekstrak teks dari setiap halaman PDF (untuk PDF digital/selectable text).
    Returns list[str], satu string per halaman.
    """
    if not PYPDF_AVAILABLE:
        return []
    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(pdf_bytes))
        return [page.extract_text() or "" for page in reader.pages]
    except Exception as e:
        st.warning(f"Gagal ekstrak teks PDF: {e}")
        return []


def detect_report_pages(page_texts: list[str]) -> dict[str, list[int]]:
    """
    Scan teks semua halaman, deteksi halaman mana yang berisi
    laporan target (Posisi Keuangan, Laba Rugi, Arus Kas).

    Returns:
        {'neraca': [0, 1], 'laba_rugi': [2, 3], 'arus_kas': [4]}
        Halaman yang tidak dikenali tidak dimasukkan.
    """
    result: dict[str, list[int]] = {}
    current_type = None

    for pg_idx, text in enumerate(page_texts):
        t = text.lower()
        found_type = None

        for rtype, keywords in TARGET_REPORT_TITLES.items():
            if any(kw in t for kw in keywords):
                found_type = rtype
                break

        if found_type:
            current_type = found_type
            result.setdefault(current_type, [])
            result[current_type].append(pg_idx)
        elif current_type and text.strip():
            # Halaman lanjutan dari laporan sebelumnya
            # Stop jika halaman ini berisi judul laporan lain atau catatan
            is_notes = any(kw in t for kw in [
                "catatan atas", "notes to", "auditor", "akuntan publik",
                "laporan auditor", "pernyataan direksi"
            ])
            if not is_notes:
                result[current_type].append(pg_idx)
            else:
                current_type = None  # reset, halaman ini bukan laporan keuangan

    return result


def detect_years_in_pages(page_texts: list[str], page_indices: list[int]) -> list[str]:
    """Kumpulkan semua tahun dari halaman-halaman tertentu, urutkan menurun."""
    all_years = []
    for idx in page_indices:
        if idx < len(page_texts):
            years = re.findall(r"\b(20[0-3]\d)\b", page_texts[idx])
            all_years.extend(years)
    return sorted(set(all_years), reverse=True)

@st.cache_resource
def get_easyocr_reader():
    """Cache EasyOCR reader (load sekali saja, model ~100MB)."""
    if EASYOCR_AVAILABLE:
        return easyocr.Reader(['id', 'en'], gpu=False)
    return None


def ocr_tesseract(image: Image.Image) -> str:
    """OCR dengan Tesseract. Config untuk tabel angka Indonesia."""
    if not TESSERACT_AVAILABLE:
        return ""
    # PSM 6 = assume uniform block of text (baik untuk tabel)
    # Bahasa: ind + eng (untuk istilah inggris)
    config = "--psm 6 -l ind+eng"
    try:
        return pytesseract.image_to_string(image, config=config)
    except Exception as e:
        st.warning(f"Tesseract gagal: {e}")
        return ""


def ocr_easyocr(image: Image.Image) -> str:
    """OCR dengan EasyOCR sebagai fallback."""
    reader = get_easyocr_reader()
    if reader is None:
        return ""
    try:
        img_array = np.array(image)
        results = reader.readtext(img_array, detail=0, paragraph=True)
        return "\n".join(results)
    except Exception as e:
        st.warning(f"EasyOCR gagal: {e}")
        return ""


def ocr_image(image: Image.Image, engine: str = "tesseract") -> str:
    """Entry point OCR. Fallback otomatis jika engine pilihan kosong."""
    if engine == "tesseract":
        text = ocr_tesseract(image)
        if not text.strip() and EASYOCR_AVAILABLE:
            st.info("Tesseract tidak menghasilkan teks, mencoba EasyOCR...")
            text = ocr_easyocr(image)
    else:
        text = ocr_easyocr(image)
        if not text.strip() and TESSERACT_AVAILABLE:
            text = ocr_tesseract(image)
    return text


# =============================================================================
# MODUL 1b: CLAUDE VISION API
# =============================================================================

# System prompt khusus ekstraksi laporan keuangan Indonesia
_VISION_SYSTEM_PROMPT = """Kamu adalah asisten ekstraksi data laporan keuangan Indonesia.
Tugasmu: baca gambar laporan keuangan dan keluarkan HANYA data terstruktur dalam format JSON.

FORMAT OUTPUT (JSON saja, tanpa penjelasan apapun):
{
  "years": ["2022", "2021"],
  "rows": [
    {"tipe": "section", "keterangan": "ASET LANCAR", "catatan": "", "nilai": []},
    {"tipe": "item",    "keterangan": "Kas dan setara kas", "catatan": "4,33", "nilai": [1068980860803, 982621700996]},
    {"tipe": "total",   "keterangan": "TOTAL ASET LANCAR", "catatan": "", "nilai": [1730514103614, 1641297422821]}
  ]
}

ATURAN PENTING:
1. "tipe" hanya boleh: "section" (judul kelompok), "item" (baris data), "total" (jumlah/total)
2. "nilai" adalah array angka bulat dalam urutan kolom kiri ke kanan (tahun terbaru dulu)
3. Angka NEGATIF (dalam kurung di laporan) → tulis sebagai angka negatif, contoh: -2745694073
4. Baris yang HANYA judul (tanpa nilai) → "nilai": []
5. Gabungkan label akun yang terbagi beberapa baris menjadi satu keterangan utuh
6. Hilangkan nomor catatan dari kolom Keterangan, taruh di field "catatan"
7. Baca dengan teliti kata-kata italic/bold seperti "bruto", "neto", "reasuransi"
8. Jangan tambahkan komentar, penjelasan, atau markdown — JSON murni saja"""


def ocr_via_claude_vision(
    image: Image.Image,
    api_key: str,
    model: str = "claude-haiku-4-5-20251001",
) -> str:
    """
    Gunakan Claude Vision API untuk membaca laporan keuangan dari gambar.

    Cara kerja:
    1. Konversi PIL Image → base64 PNG
    2. Kirim ke Claude API dengan system prompt ekstraksi terstruktur
    3. Claude mengembalikan JSON → konversi ke teks baris demi baris
       agar compatible dengan parser parse_ocr_text yang sudah ada

    Args:
        image:   PIL.Image gambar laporan keuangan
        api_key: Anthropic API key (dari st.session_state)
        model:   Model Claude yang dipakai (default Haiku = paling murah)

    Returns:
        str: Teks terformat yang bisa diproses parse_ocr_text, ATAU
             teks JSON mentah jika parsing berhasil (lebih akurat)
    """
    try:
        import anthropic
    except ImportError:
        st.error("Package 'anthropic' belum terinstall. Jalankan: pip install anthropic")
        return ""

    # Konversi gambar ke base64
    buf = io.BytesIO()
    image.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("utf-8")

    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model=model,
            max_tokens=4096,
            system=_VISION_SYSTEM_PROMPT,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": b64,
                        },
                    },
                    {
                        "type": "text",
                        "text": "Ekstrak semua baris laporan keuangan dari gambar ini. "
                                "Output JSON saja sesuai format yang diberikan.",
                    },
                ],
            }],
        )
        return response.content[0].text

    except Exception as e:
        st.error(f"Claude Vision API gagal: {e}")
        return ""


def parse_vision_json(json_text: str, num_value_cols: int = 2) -> pd.DataFrame:
    """
    Parse JSON dari Claude Vision API menjadi DataFrame.

    PENTING: num_value_cols hanya sebagai MINIMUM.
    Jika Claude mengembalikan 4 kolom nilai, DataFrame akan punya 4 kolom.
    Jangan potong dengan [:num_value_cols] — itu yang menyebabkan
    Nilai_3 dan Nilai_4 selalu kosong meskipun data ada.
    """
    import json

    clean = json_text.strip()
    if clean.startswith("```"):
        clean = re.sub(r"^```(?:json)?\s*", "", clean)
        clean = re.sub(r"\s*```$", "", clean)

    try:
        data = json.loads(clean)
    except json.JSONDecodeError as e:
        st.warning(f"JSON dari Claude tidak valid: {e}. Mencoba parse sebagai teks biasa...")
        return parse_ocr_text(json_text, num_value_cols=num_value_cols)

    rows_raw = data.get("rows", [])
    if not rows_raw:
        return pd.DataFrame(
            columns=["Tipe", "Keterangan", "Catatan"] +
                    [f"Nilai_{k+1}" for k in range(num_value_cols)]
        )

    # Tentukan jumlah kolom aktual: max panjang nilai di semua baris,
    # minimal num_value_cols
    max_nilai = max((len(r.get("nilai", [])) for r in rows_raw), default=0)
    actual_cols = max(num_value_cols, max_nilai)

    rows_out = []
    for row in rows_raw:
        nilai = list(row.get("nilai", []))
        # Padding ke actual_cols jika kurang (jangan dipotong)
        while len(nilai) < actual_cols:
            nilai.append(None)

        rows_out.append({
            "Tipe":       row.get("tipe", "item"),
            "Keterangan": row.get("keterangan", ""),
            "Catatan":    row.get("catatan", ""),
            **{f"Nilai_{k+1}": v for k, v in enumerate(nilai)},
        })

    return pd.DataFrame(rows_out)


def ocr_image_smart(
    image: Image.Image,
    engine: str,
    api_key: str = "",
    vision_model: str = "claude-haiku-4-5-20251001",
) -> tuple[str, str]:
    """
    Entry point OCR yang mendukung semua engine termasuk Claude Vision.

    Returns:
        (raw_text_or_json, mode)
        mode: 'ocr' | 'vision'  — untuk menentukan parser yang dipakai
    """
    if engine == "claude_vision":
        if not api_key:
            st.error("API Key Claude belum diisi di sidebar!")
            return "", "vision"
        result = ocr_via_claude_vision(image, api_key, model=vision_model)
        return result, "vision"
    else:
        return ocr_image(image, engine=engine), "ocr"


# =============================================================================
# MODUL 2: DETEKSI JENIS LAPORAN
# =============================================================================

def detect_report_type(text: str) -> str:
    """
    Deteksi jenis laporan berdasarkan frekuensi kata kunci.
    Returns: 'neraca' | 'laba_rugi' | 'arus_kas' | 'unknown'
    """
    t = text.lower()
    scores = {
        "neraca": sum(1 for kw in KEYWORDS_NERACA if kw in t),
        "laba_rugi": sum(1 for kw in KEYWORDS_LABA_RUGI if kw in t),
        "arus_kas": sum(1 for kw in KEYWORDS_ARUS_KAS if kw in t),
    }
    best = max(scores, key=scores.get)
    return best if scores[best] >= 2 else "unknown"


def detect_years(text: str, max_years: int = 5) -> list[str]:
    """
    Ekstrak tahun-tahun yang muncul dalam teks.
    Mendukung hingga max_years tahun (default 5) — bukan 2 saja.
    """
    years = re.findall(r"\b(20[0-3]\d)\b", text)
    unique_sorted = sorted(set(years), reverse=True)
    return unique_sorted[:max_years] if unique_sorted else ["Tahun 1", "Tahun 2"]


# =============================================================================
# MODUL 3: PARSER ANGKA & BARIS
# =============================================================================

def parse_indonesian_number(s: str) -> float | None:
    """
    Parse angka keuangan dengan format pemisah ribuan APAPUN:
      - Titik sebagai ribuan:  1.068.980.860.803  → 1068980860803
      - Koma sebagai ribuan:   375,272,610        → 375272610      (format OCR barat)
      - Negatif dalam kurung:  (96,508,401)       → -96508401
      - Negatif dengan titik:  (9.531.540.607)    → -9531540607
    Returns None jika bukan angka keuangan valid.
    """
    if not s:
        return None
    s = s.strip()

    # Deteksi negatif (dalam kurung)
    is_negative = s.startswith("(") and s.endswith(")")
    if is_negative:
        s = s[1:-1]

    # Bersihkan simbol mata uang & spasi
    s = re.sub(r"[Rp\s$]", "", s).strip()
    if not s:
        return None

    # Strategi: hapus SEMUA pemisah ribuan (titik dan koma), sisakan hanya digit
    # Ini aman karena angka keuangan di laporan keuangan tidak punya desimal
    # (nilai dalam Rupiah penuh, bukan sen)
    # Kecuali: format "1.234,56" atau "1,234.56" (ada desimal) → tangani khusus

    # Hitung jumlah titik dan koma
    dot_count   = s.count(".")
    comma_count = s.count(",")

    if dot_count == 1 and comma_count == 1:
        # Mungkin ada desimal: cek mana yang lebih kanan (= desimal separator)
        dot_pos   = s.rfind(".")
        comma_pos = s.rfind(",")
        if comma_pos > dot_pos:
            # Format: 1.234,56 → titik=ribuan, koma=desimal
            s = s.replace(".", "").replace(",", ".")
        else:
            # Format: 1,234.56 → koma=ribuan, titik=desimal
            s = s.replace(",", "")
    else:
        # Hapus semua pemisah ribuan (titik dan koma)
        s = s.replace(".", "").replace(",", "")

    try:
        val = float(s)
        # Sanity check: angka keuangan biasanya ≥ 1 atau 0
        return (-val if is_negative else val)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Pola angka keuangan: mengenali KEDUA format pemisah ribuan
#
# Format titik  : 1.068.980.860.803 | (9.531.540.607)
# Format koma   : 375,272,610       | (96,508,401)    ← OCR barat/campuran
#
# Aturan: minimal 2 grup pemisah (≥7 digit total) supaya tidak salah tangkap
# nomor catatan seperti "4,33" atau "19" atau "2017"
# ---------------------------------------------------------------------------
_PAT_FINANCIAL_NUM = re.compile(
    r"\(?"                           # kurung buka opsional (angka negatif)
    r"\d{1,3}"                       # 1-3 digit pertama
    r"(?:"
        r"(?:\.\d{3})+"              # format titik: 1+ grup → ≥4 digit
        r"|"
        r"(?:,\d{3})+"               # format koma:  1+ grup → ≥4 digit
    r")"
    r"\)?"                           # kurung tutup opsional
)
_PAT_YEAR  = re.compile(r"\b20[0-3]\d\b")
_PAT_NOTES = re.compile(r"^\s*[\d,a-z]{1,10}\s*$", re.IGNORECASE)


def _is_section_header(line: str) -> bool:
    """
    True jika baris adalah judul section murni:
    ALL CAPS, tidak mengandung angka keuangan (≥7 digit).

    Catatan: baris seperti 'LABA BERSIH 360,724,480 366,412,599'
    BUKAN section header — mengandung angka → False.
    """
    t = line.strip()
    if not t:
        return False
    # Jika ada angka keuangan → bukan pure section header
    if _PAT_FINANCIAL_NUM.search(t):
        return False
    # Harus ALL CAPS (huruf semua kapital), minimal 3 karakter
    # isupper() True hanya jika ada setidaknya satu huruf dan semua huruf kapital
    return t.isupper() and len(t) >= 3 and not t.isdigit()


def _extract_financial_numbers(line: str) -> list[float]:
    """
    Ekstrak semua angka keuangan dari satu baris.
    Filter: nilai absolut >= 1000 (menghindari nomor catatan, tahun, dll.)
    """
    results = []
    for m in _PAT_FINANCIAL_NUM.finditer(line):
        raw = m.group()
        val = parse_indonesian_number(raw)
        if val is not None and abs(val) >= 1000:
            results.append(val)
    return results


def _label_before_numbers(line: str) -> str:
    """Ambil teks sebelum angka keuangan pertama di baris."""
    m = _PAT_FINANCIAL_NUM.search(line)
    if m:
        return line[:m.start()].strip()
    return line.strip()


def _clean_label(raw: str) -> str:
    """
    Bersihkan label dari artefak OCR:
    - Hapus nomor catatan yang tertinggal di akhir (pola: spasi + angka/huruf pendek)
    - Tidak memotong label ALL CAPS (section/total header)
    - Normalisasi spasi
    """
    t = raw.strip()
    # Jika ALL CAPS → kemungkinan section/total header, jangan dipotong
    if t.isupper():
        return re.sub(r"\s+", " ", t)
    # Hapus trailing nomor catatan: pola spasi + angka/huruf pendek di akhir
    # Contoh: '  4,33' | '  19a' | '  32' | '  12,14,33'
    cleaned = re.sub(r"\s+[\d,;a-z]{1,12}\s*$", "", t, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned if cleaned else t


def _classify_row(label: str) -> str:
    """
    Klasifikasi tipe baris: 'section' | 'total' | 'item'.

    Aturan:
    - ALL CAPS + mengandung kata total/jumlah/laba/rugi = 'total'
    - ALL CAPS tanpa kata kunci = 'section'
    - Diawali 'Total'/'Jumlah' = 'total'
    - Lainnya = 'item'
    """
    l = label.lower().strip()
    if not l:
        return "empty"

    upper = label.strip().isupper()

    if upper and len(label.strip()) >= 3:
        if any(kw in l for kw in TOTAL_KEYWORDS):
            return "total"
        # ALL CAPS yang diikuti nilai tapi bukan keyword total → tetap 'total'
        # (misal "LABA USAHA OPERASIONAL", "BEBAN PAJAK PENGHASILAN")
        # Jadikan 'total' agar dapat styling bold dan formula di Excel
        return "total"

    if any(l.startswith(kw) for kw in ["total ", "jumlah "]):
        return "total"

    return "item"


# ---------------------------------------------------------------------------
# LANGKAH 1 — Deteksi header kolom (tahun pelaporan)
# ---------------------------------------------------------------------------

def detect_header_row(lines: list[str], scan_rows: int = 8) -> dict:
    """
    LANGKAH 1: Scan baris pertama (default 8) untuk menemukan baris header
    yang memuat kolom tahun.

    Returns dict:
        {
          'header_line': str,          # teks baris header
          'header_idx':  int,          # indeks baris header
          'years':       ['2022','2021'],
          'col_positions': [pos1, pos2] # posisi karakter kanan kolom nilai
        }
    """
    result = {
        'header_line': '',
        'header_idx': -1,
        'years': [],
        'col_positions': [],
    }

    for i, line in enumerate(lines[:scan_rows]):
        years_found = _PAT_YEAR.findall(line)
        if len(years_found) >= 1:
            positions = [m.end() + 10 for m in _PAT_YEAR.finditer(line)]
            result['header_line']   = line
            result['header_idx']    = i
            # Ambil hingga 5 tahun — jangan batasi ke 2
            result['years']         = sorted(set(years_found), reverse=True)[:5]
            result['col_positions'] = positions
            # Jika kurang dari yang diharapkan, cek baris berikutnya
            if len(result['years']) < 2 and i + 1 < len(lines):
                next_years = _PAT_YEAR.findall(lines[i + 1])
                all_years = sorted(set(years_found + next_years), reverse=True)[:5]
                result['years'] = all_years
            break

    # Fallback: tidak ketemu header → default 2 kolom
    if not result['years']:
        result['years'] = ['Tahun 1', 'Tahun 2']

    return result


# ---------------------------------------------------------------------------
# LANGKAH 2 — Filter baris berAngka
# ---------------------------------------------------------------------------

def filter_numeric_lines(lines: list[str], header_idx: int) -> list[tuple[int, str, list[float]]]:
    """
    LANGKAH 2: Skip baris header dan semua baris di atasnya, lalu
    ambil hanya baris yang memiliki angka keuangan.

    FIX: header_idx dan semua baris sebelumnya di-skip (bukan ikut diproses).
    """
    numeric_lines = []
    # +1: mulai SETELAH baris header, bukan dari header
    start = header_idx + 1 if header_idx >= 0 else 0
    for i, line in enumerate(lines[start:], start=start):
        nums = _extract_financial_numbers(line)
        if nums:
            numeric_lines.append((i, line, nums))
    return numeric_lines


# ---------------------------------------------------------------------------
# LANGKAH 3 — Rekonstruksi label akun multi-baris
# ---------------------------------------------------------------------------

def reconstruct_labels(
    lines: list[str],
    numeric_lines: list[tuple[int, str, list[float]]],
    header_idx: int,
    num_value_cols: int = 2,
) -> pd.DataFrame:
    """
    LANGKAH 3: Untuk setiap baris berAngka, rekonstruksi label akun yang
    mungkin tersebar di beberapa baris sebelumnya.

    FIX 1 — Skip header zone:
      Tidak memproses baris di atas atau sama dengan header_idx.

    FIX 2 — Continuation lebih agresif:
      Baris yang diawali huruf kecil atau kata sambung ('dan','yang','atau',
      'belum','dengan') SELALU dianggap continuation, bukan baris mandiri.

    FIX 3 — Label pendek (≤2 kata) → selalu gabung ke atas:
      Jika label di baris berAngka hanya 1-2 kata (misal "Premi", "Klaim"),
      gabungkan dengan baris-baris di atasnya untuk mendapat label lengkap.

    FIX 4 — Section header dengan angka (ALL CAPS + angka):
      Baris seperti "LABA USAHA OPERASIONAL  375,272,610  435,531,332"
      seharusnya jadi baris 'total', bukan section tanpa nilai.
      _is_section_header sudah return False untuk kasus ini,
      tapi _classify_row perlu menangkapnya sebagai 'total'.
    """
    rows = []
    numeric_idx_set = {idx for idx, _, _ in numeric_lines}
    # Batas bawah yang aman untuk melihat ke atas
    safe_start = header_idx + 1 if header_idx >= 0 else 0

    # Kata-kata yang menandai baris ini adalah continuation dari baris atas
    CONTINUATION_STARTERS = {
        'dan', 'yang', 'atau', 'dengan', 'belum', 'setelah', 'untuk',
        'atas', 'dari', 'dalam', 'pada', 'oleh', 'serta', 'maupun',
        'telah', 'akan', 'tidak', 'bukan', 'sebagai', 'terhadap',
    }

    def _is_continuation(text: str) -> bool:
        """True jika teks kemungkinan lanjutan dari baris sebelumnya."""
        t = text.strip()
        if not t:
            return False
        # Dimulai huruf kecil → hampir pasti continuation
        if t[0].islower():
            return True
        # Dimulai kata sambung (meski kapital)
        first_word = t.split()[0].lower().rstrip('/()')
        if first_word in CONTINUATION_STARTERS:
            return True
        return False

    def _is_short_label(label: str) -> bool:
        """True jika label terlalu pendek (≤2 kata) → perlu gabung ke atas."""
        words = label.strip().split()
        return len(words) <= 2

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Skip baris di zona header
        if i <= header_idx:
            i += 1
            continue

        # Tandai section header MURNI (ALL CAPS tanpa angka)
        if _is_section_header(stripped):
            rows.append({
                "Tipe":       "section",
                "Keterangan": stripped,
                "Catatan":    "",
                **{f"Nilai_{k+1}": None for k in range(num_value_cols)},
            })
            i += 1
            continue

        nums = _extract_financial_numbers(line)
        if not nums:
            i += 1
            continue

        # ── Baris berAngka ditemukan ──────────────────────────────────────
        label_from_line = _label_before_numbers(line)
        indent_this = len(line) - len(line.lstrip())

        # ─── Tentukan apakah perlu cari continuation ke atas ───────────
        # Label sudah cukup informatif jika ≥ 3 kata DAN bukan pendek
        label_words = label_from_line.strip().split()
        label_is_sufficient = len(label_words) >= 3

        continuation_parts = []

        if not label_is_sufficient:
            # Label pendek → perlu melihat ke atas untuk melengkapi
            j = i - 1
            while j >= safe_start:
                prev_raw    = lines[j]
                prev        = prev_raw.strip()
                indent_prev = len(prev_raw) - len(prev_raw.lstrip())

                if not prev:
                    break                               # baris kosong → stop
                if i - j > 6:
                    break                               # terlalu jauh
                if _is_section_header(prev):
                    break                               # ALL CAPS tanpa angka → stop
                if _extract_financial_numbers(prev_raw):
                    break                               # baris berAngka lain → stop
                if _PAT_NOTES.match(prev):
                    j -= 1
                    continue                            # nomor catatan → skip

                # Jika baris atas jelas lanjutan kalimat → gabung
                if _is_continuation(prev):
                    continuation_parts.insert(0, prev)
                    j -= 1
                    continue

                # Jika baris atas sudah punya ≥ 3 kata → itu label mandiri, stop
                if len(prev.split()) >= 3:
                    # Jangan ambil, tapi juga jangan stop — baris atas ini
                    # mungkin context induk. Ambil hanya jika label hasil
                    # gabungan masih pendek.
                    continuation_parts.insert(0, prev)
                    break
                else:
                    continuation_parts.insert(0, prev)

                j -= 1

        elif _is_continuation(label_from_line):
            # Label di baris ini dimulai dengan kata sambung → pasti butuh gabung
            j = i - 1
            while j >= safe_start:
                prev_raw = lines[j]
                prev     = prev_raw.strip()
                if not prev or _is_section_header(prev):
                    break
                if _extract_financial_numbers(prev_raw):
                    break
                if _PAT_NOTES.match(prev):
                    j -= 1
                    continue
                continuation_parts.insert(0, prev)
                j -= 1

        # ── Gabungkan label ───────────────────────────────────────────────
        # Strategi: prefix (continuation) + label dari baris ini
        if continuation_parts and label_from_line:
            raw_label = " ".join(continuation_parts) + " " + label_from_line
        elif continuation_parts:
            raw_label = " ".join(continuation_parts)
        else:
            raw_label = label_from_line

        label = _clean_label(raw_label)
        if not label:
            label = stripped

        # Baris tanpa label bermakna (hanya angka/tanda baca) — biasanya
        # angka yatim dari layout dwibahasa/dua kolom hasil OCR.
        if re.fullmatch(r"[\d.,()\-\s%]*", label):
            # Pola umum laporan dwibahasa: label total di satu baris
            # ("JUMLAH ASET"), nilainya di baris berikutnya. Tempelkan
            # angka yatim ke baris total/jumlah di atasnya yang masih
            # punya slot nilai kosong.
            if rows:
                prev_row = rows[-1]
                prev_label = str(prev_row.get("Keterangan", "")).lower()
                is_total_like = any(kw in prev_label for kw in ("total", "jumlah"))
                if is_total_like:
                    empty_slots = [k for k in range(num_value_cols)
                                   if prev_row.get(f"Nilai_{k+1}") is None]
                    for v in nums:
                        if not empty_slots:
                            break
                        prev_row[f"Nilai_{empty_slots.pop(0)+1}"] = v
                    if prev_row["Tipe"] == "section":
                        prev_row["Tipe"] = "total"
            # Apapun hasilnya, jangan jadikan "akun" dengan nama berupa angka
            i += 1
            continue

        # ── Pisahkan nomor catatan dari nilai ─────────────────────────────
        # Layout: [Nilai_t1] [Catatan kecil] [Nilai_t0]
        # Catatan ada di TENGAH (bukan awal), tapi OCR kadang
        # meletakkannya setelah label atau di antara dua nilai.
        # Heuristik: jika ada angka kecil (abs ≤ 100) di antara dua angka besar
        values = nums.copy()
        catatan = ""

        if len(values) > num_value_cols:
            # Cari angka kecil yang mungkin adalah nomor catatan
            # Posisi tengah (antara nilai pertama dan terakhir)
            filtered = []
            cat_candidates = []
            for v in values:
                if abs(v) <= 100:
                    cat_candidates.append(str(int(v)) if v == int(v) else str(v))
                else:
                    filtered.append(v)
            if cat_candidates:
                catatan = ",".join(cat_candidates)
                values = filtered
            elif values[0] <= 1000 and abs(values[0]) < abs(values[-1]):
                # Angka pertama jauh lebih kecil → nomor catatan
                catatan = str(int(values[0])) if values[0] == int(values[0]) else str(values[0])
                values = values[1:]

        # Padding / truncate ke num_value_cols
        while len(values) < num_value_cols:
            values.append(None)
        values = values[-num_value_cols:]

        # ── Klasifikasi tipe ──────────────────────────────────────────────
        row_type = _classify_row(label)

        rows.append({
            "Tipe":       row_type,
            "Keterangan": label,
            "Catatan":    catatan,
            **{f"Nilai_{k+1}": v for k, v in enumerate(values)},
        })
        i += 1

    return pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["Tipe", "Keterangan", "Catatan"] + [f"Nilai_{k+1}" for k in range(num_value_cols)]
    )


# ---------------------------------------------------------------------------
# Pre-cleaning teks OCR: perbaiki typo umum pada angka
# ---------------------------------------------------------------------------

# Huruf yang sering salah baca oleh OCR di dalam angka: O→0, Z→7, l/I→1, S→5, B→8
_OCR_DIGIT_MAP = str.maketrans("OoZzlISB", "00771158")


def _preclean_ocr_line(line: str) -> str:
    """
    Perbaiki typo OCR umum pada angka SEBELUM parsing:
    - Pemisah ganda:  '6,326,.921' / '381.,453'  → '6,326,921' / '381,453'
    - Spasi di dalam angka: '19. 063'            → '19.063'
    - Huruf mirip digit di token angka: '19,063,10Z' → '19,063,107'
      (hanya jika token didominasi digit DAN punya pemisah ribuan,
       supaya kata biasa tidak ikut berubah)
    """
    # Pemisah ganda ".," / ",." / ".." di antara digit → satu pemisah
    line = re.sub(r"(?<=\d)[.,]{2,}(?=\d)", ",", line)
    # Spasi setelah pemisah ribuan di dalam angka
    line = re.sub(r"(?<=\d)([.,])\s+(?=\d{3}(?:\D|$))", r"\1", line)

    def _fix_token(m: re.Match) -> str:
        tok = m.group(0)
        n_digits = sum(ch.isdigit() for ch in tok)
        n_alpha  = sum(ch.isalpha() for ch in tok)
        if n_digits >= 4 and n_alpha <= 2 and ("." in tok or "," in tok):
            return tok.translate(_OCR_DIGIT_MAP)
        return tok

    return re.sub(r"[\d.,()OoZzlISB]+", _fix_token, line)


# ---------------------------------------------------------------------------
# Entry point: parse_ocr_text (menggabungkan 3 langkah)
# ---------------------------------------------------------------------------

def parse_ocr_text(text: str, num_value_cols: int = 2) -> pd.DataFrame:
    """
    Algoritma 3-langkah:
    1. Scan baris awal → deteksi header kolom tahun
    2. Filter baris berAngka
    3. Rekonstruksi label multi-baris

    PENTING: num_value_cols hanya dipakai sebagai MINIMUM fallback.
    Jika header berhasil mendeteksi lebih banyak tahun (misal 4 tahun),
    maka otomatis pakai jumlah tahun tersebut. Ini memastikan laporan
    dengan 4 kolom tahun menghasilkan 4 Nilai_ kolom, bukan 2.
    """
    lines = [_preclean_ocr_line(l) for l in text.split("\n")]

    # Langkah 1 — deteksi header
    header_info = detect_header_row(lines, scan_rows=8)

    # Gunakan jumlah tahun aktual jika lebih besar dari num_value_cols
    actual_cols = max(num_value_cols, len(header_info['years']))

    # Langkah 2
    numeric_lines = filter_numeric_lines(lines, header_info['header_idx'])

    # Langkah 3
    df = reconstruct_labels(lines, numeric_lines,
                            header_idx=header_info['header_idx'],
                            num_value_cols=actual_cols)

    return df


# Backward compat — fungsi lama yang masih dipakai test
def extract_numbers_from_line(line: str) -> tuple[str, list[float]]:
    label = _label_before_numbers(line)
    nums  = _extract_financial_numbers(line)
    return label, nums


def classify_row(label: str) -> str:
    return _classify_row(label)


# =============================================================================
# MODUL 4: VALIDASI SILANG
# =============================================================================

def validate_neraca(df: pd.DataFrame, value_cols: list[str]) -> list[dict]:
    """
    Validasi: Total Aset == Total Liabilitas + Total Ekuitas
    Returns list of dicts dengan status per kolom nilai.
    """
    results = []
    for col in value_cols:
        total_aset = None
        total_liab = None
        total_ekuitas = None
        total_liab_ekuitas = None

        for _, row in df.iterrows():
            label = str(row["Keterangan"]).lower().strip()
            val = row[col]
            if pd.isna(val):
                continue
            # Cari baris total utama (bukan sub-total)
            if label.startswith("total aset") and "lancar" not in label and "tidak" not in label:
                total_aset = val
            elif label.startswith("total liabilitas") and "jangka" not in label:
                total_liab = val
            elif label.startswith("total ekuitas") or label.startswith("jumlah ekuitas"):
                total_ekuitas = val
            elif "total liabilitas dan ekuitas" in label or "jumlah liabilitas dan ekuitas" in label:
                total_liab_ekuitas = val

        # Tentukan total right-side
        right_side = total_liab_ekuitas
        if right_side is None and total_liab is not None and total_ekuitas is not None:
            right_side = total_liab + total_ekuitas

        if total_aset is not None and right_side is not None:
            selisih = total_aset - right_side
            status = "✅ Seimbang" if abs(selisih) < 1 else f"⚠ Selisih: {selisih:,.0f}"
            results.append({
                "Kolom": col,
                "Total Aset": total_aset,
                "Total Liab+Ekuitas": right_side,
                "Status": status,
            })
        else:
            results.append({
                "Kolom": col,
                "Total Aset": total_aset,
                "Total Liab+Ekuitas": right_side,
                "Status": "ℹ Data tidak lengkap untuk validasi",
            })
    return results


def validate_subtotal_sums(df: pd.DataFrame, value_cols: list[str]) -> list[dict]:
    """
    Validasi: apakah jumlah item di bawah section = nilai total section-nya.
    Deteksi sederhana: iterasi dari total ke atas, sum semua item setelah section terakhir.
    """
    results = []
    for col in value_cols:
        items_buffer = []
        for idx, row in df.iterrows():
            label = str(row["Keterangan"])
            tipe = row["Tipe"]
            val = row[col]

            if tipe == "section":
                items_buffer = []
            elif tipe == "item" and pd.notna(val):
                items_buffer.append(val)
            elif tipe == "total" and pd.notna(val):
                if items_buffer:
                    sum_items = sum(items_buffer)
                    selisih = val - sum_items
                    if abs(selisih) > 1:
                        results.append({
                            "Kolom": col,
                            "Total": label,
                            "Nilai Total": val,
                            "Sum Item": sum_items,
                            "Selisih": selisih,
                        })
                items_buffer = []
    return results


# =============================================================================
# MODUL 5: EXPORT EXCEL
# =============================================================================

def build_excel(
    reports: dict,
    entity_name: str = "PT [Nama Perusahaan]",
) -> bytes:
    """
    Build Excel workbook dari dict of reports.
    reports = {
        'neraca': {'df': df, 'years': [...], 'title': 'Laporan Posisi Keuangan'},
        'laba_rugi': {...},
        'arus_kas': {...},
    }
    Returns: bytes of xlsx file
    """
    wb = Workbook()
    wb.remove(wb.active)

    FONT_NAME = "Arial"
    title_font = Font(name=FONT_NAME, size=12, bold=True)
    subtitle_font = Font(name=FONT_NAME, size=10, italic=True)
    header_font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    section_font = Font(name=FONT_NAME, size=10, bold=True)
    total_font = Font(name=FONT_NAME, size=10, bold=True)
    normal_font = Font(name=FONT_NAME, size=10)

    header_fill = PatternFill("solid", start_color="1F4E79")
    section_fill = PatternFill("solid", start_color="D9E1F2")
    total_fill = PatternFill("solid", start_color="B4C7E7")

    medium = Side(border_style="medium", color="000000")
    thin = Side(border_style="thin", color="999999")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_total = Border(top=medium, bottom=medium)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    left_indent = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)

    # Format angka: ribuan dengan koma, negatif dalam kurung, nol tampil "0"
    # (akun yang tidak ada di suatu tahun diisi 0 dan harus terlihat sebagai 0)
    NUM_FMT = '#,##0;(#,##0);0'

    # Sheet Info
    ws_info = wb.create_sheet("Info")
    ws_info.column_dimensions["A"].width = 35
    ws_info.column_dimensions["B"].width = 60
    ws_info["A1"] = "LAPORAN KEUANGAN KONSOLIDASIAN"
    ws_info["A1"].font = title_font
    ws_info["A2"] = entity_name
    ws_info["A2"].font = title_font
    ws_info.merge_cells("A1:B1")
    ws_info.merge_cells("A2:B2")

    info_rows = [
        ("", ""),
        ("Tanggal Generate", datetime.now().strftime("%d %B %Y, %H:%M")),
        ("Dibuat dengan", "Konverter Laporan Keuangan (OCR + Streamlit)"),
        ("", ""),
        ("STATUS KELENGKAPAN", ""),
    ]
    for key, jenis_label in [("neraca", "Laporan Posisi Keuangan"),
                              ("laba_rugi", "Laporan Laba Rugi"),
                              ("arus_kas", "Laporan Arus Kas")]:
        status = "✅ Tersedia" if key in reports else "❌ Belum diupload"
        info_rows.append((jenis_label, status))

    info_rows.append(("", ""))
    info_rows.append(("PENTING", ""))
    info_rows.append(("Review manual",
                      "Hasil OCR TIDAK 100% akurat. Pastikan sudah di-review & edit "
                      "sebelum digunakan untuk keperluan profesional."))

    for i, (k, v) in enumerate(info_rows, start=3):
        ws_info.cell(row=i, column=1, value=k).font = section_font if k.isupper() else normal_font
        c = ws_info.cell(row=i, column=2, value=v)
        c.font = normal_font
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws_info.sheet_view.showGridLines = False

    # Sheets untuk setiap laporan
    sheet_name_map = {
        "neraca": "Laporan Posisi Keuangan",
        "laba_rugi": "Laporan Laba Rugi",
        "arus_kas": "Laporan Arus Kas",
    }
    title_id_map = {
        "neraca": "LAPORAN POSISI KEUANGAN KONSOLIDASIAN",
        "laba_rugi": "LAPORAN LABA RUGI DAN PENGHASILAN KOMPREHENSIF LAIN KONSOLIDASIAN",
        "arus_kas": "LAPORAN ARUS KAS KONSOLIDASIAN",
    }

    for key, data in reports.items():
        df = data["df"]
        years = data["years"]
        sheet_name = sheet_name_map.get(key, key.title())

        ws = wb.create_sheet(sheet_name)
        ws.column_dimensions["A"].width = 50
        ws.column_dimensions["B"].width = 12
        for i in range(len(years)):
            ws.column_dimensions[get_column_letter(3 + i)].width = 22

        # Title
        ws.cell(row=1, column=1, value=entity_name).font = title_font
        ws.cell(row=2, column=1, value=title_id_map.get(key, sheet_name)).font = title_font
        ws.cell(row=3, column=1, value="(Disajikan dalam Rupiah, kecuali dinyatakan lain)").font = subtitle_font
        last_col = 2 + len(years)
        for r in range(1, 4):
            ws.cell(row=r, column=1).alignment = center
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)

        # Header
        headers = ["Keterangan", "Catatan"] + [str(y) for y in years]
        for i, h in enumerate(headers, start=1):
            c = ws.cell(row=5, column=i, value=h)
            c.font = header_font
            c.fill = header_fill
            c.alignment = center
            c.border = border_all
        ws.row_dimensions[5].height = 28

        # Data rows
        current_row = 6
        section_start = None
        section_total_row = None

        for _, row in df.iterrows():
            label = row["Keterangan"]
            catatan = row.get("Catatan", "")
            tipe = row["Tipe"]

            cell_label = ws.cell(row=current_row, column=1, value=label)
            cell_cat = ws.cell(row=current_row, column=2, value=str(catatan) if catatan else "")
            cell_cat.alignment = center
            cell_cat.font = normal_font

            value_cols_in_df = [f"Nilai_{i+1}" for i in range(len(years))]

            if tipe == "section":
                cell_label.font = section_font
                cell_label.fill = section_fill
                cell_label.alignment = left
                for col in range(2, last_col + 1):
                    ws.cell(row=current_row, column=col).fill = section_fill
                section_start = current_row + 1

            elif tipe == "total":
                cell_label.font = total_font
                cell_label.fill = total_fill
                cell_label.alignment = left
                # Formula SUM jika ada section_start, else hardcode
                for i, col_df in enumerate(value_cols_in_df):
                    col_idx = 3 + i
                    col_letter = get_column_letter(col_idx)
                    val = row.get(col_df)
                    if section_start and section_start < current_row:
                        formula = f"=SUM({col_letter}{section_start}:{col_letter}{current_row - 1})"
                        c = ws.cell(row=current_row, column=col_idx, value=formula)
                    else:
                        c = ws.cell(row=current_row, column=col_idx, value=val if pd.notna(val) else None)
                    c.font = total_font
                    c.fill = total_fill
                    c.number_format = NUM_FMT
                    c.alignment = right
                    c.border = border_total
                ws.cell(row=current_row, column=2).fill = total_fill
                section_start = None

            else:  # item
                cell_label.font = normal_font
                cell_label.alignment = left_indent
                for i, col_df in enumerate(value_cols_in_df):
                    val = row.get(col_df)
                    c = ws.cell(row=current_row, column=3 + i,
                                value=val if pd.notna(val) else None)
                    c.font = normal_font
                    c.number_format = NUM_FMT
                    c.alignment = right

            current_row += 1

        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A6"

    # Save to bytes
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# =============================================================================
# UI STREAMLIT
# =============================================================================

def init_session_state():
    defaults = {
        "ocr_results":           {},
        "pdf_results":           {},
        "pdf_page_selections":   {},
        "entity_name":           "PT [Nama Perusahaan]",
        "ocr_engine":            "tesseract",
        "num_value_cols":        2,
        "pasted_images":         [],
        "paste_counter":         0,
        "claude_api_key":        "",
        "claude_vision_model":   "claude-haiku-4-5-20251001",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def render_sidebar():
    with st.sidebar:
        st.header("⚙️ Pengaturan")

        st.session_state.entity_name = st.text_input(
            "Nama Entitas",
            value=st.session_state.entity_name,
        )
        st.session_state.num_value_cols = st.number_input(
            "Jumlah Kolom Nilai (Tahun)",
            min_value=1, max_value=5,
            value=st.session_state.num_value_cols,
        )

        # ── OCR Engine ───────────────────────────────────────────────────
        st.divider()
        st.subheader("🔍 Engine Ekstraksi")

        engines_available = []
        if TESSERACT_AVAILABLE:
            engines_available.append("tesseract")
        if EASYOCR_AVAILABLE:
            engines_available.append("easyocr")
        # Claude Vision selalu tersedia (cek API key saat run)
        engines_available.append("claude_vision")

        engine_labels = {
            "tesseract":     "Tesseract OCR (gratis, lokal)",
            "easyocr":       "EasyOCR (gratis, lokal)",
            "claude_vision": "✨ Claude Vision API (berbayar, akurasi tinggi)",
        }

        st.session_state.ocr_engine = st.radio(
            "Pilih engine:",
            options=engines_available,
            format_func=lambda x: engine_labels.get(x, x),
            index=engines_available.index(st.session_state.ocr_engine)
                  if st.session_state.ocr_engine in engines_available else 0,
        )

        # ── Khusus Claude Vision ─────────────────────────────────────────
        if st.session_state.ocr_engine == "claude_vision":
            st.info("Claude Vision membaca gambar secara langsung — lebih akurat untuk teks italic, bold, dan multi-baris.")

            # API Key
            api_key_input = st.text_input(
                "Anthropic API Key",
                value=st.session_state.claude_api_key,
                type="password",
                help="Dapatkan di: https://console.anthropic.com/settings/keys",
                placeholder="sk-ant-api03-...",
            )
            if api_key_input:
                st.session_state.claude_api_key = api_key_input

            # Model
            vision_models = {
                "claude-haiku-4-5-20251001": "Haiku 4.5 — Rp 65/hal (cepat, hemat)",
                "claude-sonnet-4-6":         "Sonnet 4.6 — Rp 194/hal (akurasi terbaik)",
            }
            st.session_state.claude_vision_model = st.selectbox(
                "Model",
                options=list(vision_models.keys()),
                format_func=lambda x: vision_models.get(x, x),
                index=list(vision_models.keys()).index(
                    st.session_state.claude_vision_model
                ) if st.session_state.claude_vision_model in vision_models else 0,
            )

            # Status API key
            if st.session_state.claude_api_key:
                if st.session_state.claude_api_key.startswith("sk-ant"):
                    st.success("✅ API Key terisi")
                else:
                    st.warning("⚠ Format API Key tidak biasa (harus diawali 'sk-ant')")
            else:
                st.warning("⚠ Isi API Key untuk menggunakan Claude Vision")

            # Cara mendapatkan API key
            with st.expander("📖 Cara mendapatkan API Key"):
                st.markdown("""
1. Buka [console.anthropic.com](https://console.anthropic.com)
2. Daftar / login dengan email
3. Klik **Settings → API Keys**
4. Klik **Create Key**
5. Copy key (hanya tampil sekali) → paste di sini
6. Top-up kredit minimal $5 di **Billing**

> **Kredit awal:** Akun baru biasanya mendapat $5 kredit gratis.
> Dengan Haiku 4.5, $5 cukup untuk **~76.000 halaman** laporan keuangan.
                """)
        else:
            st.write(f"Tesseract: {'✅' if TESSERACT_AVAILABLE else '❌ (install)'}")
            st.write(f"EasyOCR:   {'✅' if EASYOCR_AVAILABLE else '❌ (install)'}")

        # ── Status Antrian ────────────────────────────────────────────────
        st.divider()
        st.subheader("📋 Status")
        n_paste   = len(st.session_state.get("pasted_images", []))
        n_img     = len(st.session_state.get("ocr_results", {}))
        n_pdf     = len(st.session_state.get("pdf_results", {}))
        pypdf_ok  = "✅" if PYPDF_AVAILABLE else "❌ pip install pypdf"
        pdf2img_ok = "✅" if PDF2IMAGE_AVAILABLE else "⚠ opsional (pip install pdf2image)"
        st.write(f"pypdf: {pypdf_ok}")
        st.write(f"pdf2image: {pdf2img_ok}")
        st.write(f"PDF diproses: **{n_pdf}** file")
        st.write(f"Gambar di antrian: **{n_paste}**")
        st.write(f"Hasil OCR: **{n_img}**")

        st.divider()
        if st.button("🗑️ Reset Semua", use_container_width=True):
            st.session_state.ocr_results         = {}
            st.session_state.pdf_results         = {}
            st.session_state.pdf_page_selections = {}
            st.session_state.pasted_images       = []
            st.session_state.paste_counter       = 0
            st.rerun()


# =============================================================================
# MODUL 6: CLIPBOARD PASTE  (v2 — fixed)
# =============================================================================

# Zona paste visual (hanya untuk tampilan & drag-drop).
# Komunikasi ke Python lewat HIDDEN FILE INPUT yang di-trigger JS,
# bukan lewat postMessage yang tidak bisa diterima Python.
CLIPBOARD_JS = """
<style>
  #paste-zone {
    border: 2.5px dashed #1F4E79;
    border-radius: 10px;
    padding: 28px 20px;
    text-align: center;
    background: #f0f4f9;
    cursor: pointer;
    transition: all 0.2s ease;
    user-select: none;
    font-family: Arial, sans-serif;
  }
  #paste-zone:hover, #paste-zone.dragover {
    background: #d9e8f7;
    border-color: #0d6efd;
  }
  #paste-zone.active {
    background: #d4edda;
    border-color: #198754;
  }
  #paste-zone .icon { font-size: 2.2rem; margin-bottom: 6px; }
  #paste-zone .main-text { font-size: 1.05rem; font-weight: bold; color: #1F4E79; }
  #paste-zone .sub-text  { font-size: 0.82rem; color: #555; margin-top: 4px; }
  #status-msg { margin-top: 10px; font-size: 0.85rem; color: #333; min-height: 20px; font-family: Arial, sans-serif; }
  #thumb-wrap { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px; justify-content: center; }
  .thumb { width: 90px; height: 70px; object-fit: contain; border: 1px solid #ccc;
           border-radius: 4px; background: #fff; }
</style>

<div id="paste-zone" tabindex="0" title="Klik lalu Ctrl+V / Cmd+V untuk paste gambar">
  <div class="icon">📋</div>
  <div class="main-text">Klik di sini, lalu tekan Ctrl+V / ⌘V</div>
  <div class="sub-text">atau Drag &amp; Drop gambar ke sini</div>
</div>
<div id="status-msg"></div>
<div id="thumb-wrap"></div>

<script>
  const zone     = document.getElementById('paste-zone');
  const statusEl = document.getElementById('status-msg');
  const thumbWrap= document.getElementById('thumb-wrap');

  function setStatus(msg, color) {
    statusEl.textContent = msg;
    statusEl.style.color = color || '#333';
  }

  function processFile(file) {
    if (!file || !file.type.startsWith('image/')) {
      setStatus('⚠ Bukan file gambar. Coba lagi.', 'red');
      return;
    }
    const reader = new FileReader();
    reader.onload = function(e) {
      const dataUrl = e.target.result;
      const base64  = dataUrl.split(',')[1];
      const mime    = file.type;
      const fname   = file.name || ('paste_' + Date.now() + '.png');

      // Tambahkan thumbnail
      const img = document.createElement('img');
      img.src   = dataUrl;
      img.className = 'thumb';
      img.title = fname;
      thumbWrap.appendChild(img);

      zone.classList.add('active');
      setTimeout(() => zone.classList.remove('active'), 800);
      setStatus('✅ Gambar berhasil ditangkap: ' + fname, '#198754');

      // Kirim ke Streamlit via query param trick + hidden anchor
      // Kita pakai window.parent.postMessage karena st.components iframe
      window.parent.postMessage(
        JSON.stringify({ type: 'clipboard_image', base64: base64, mime: mime, filename: fname }),
        '*'
      );
    };
    reader.readAsDataURL(file);
  }

  // Paste event (Ctrl+V)
  zone.addEventListener('click', () => zone.focus());
  document.addEventListener('paste', function(e) {
    const items = (e.clipboardData || e.originalEvent.clipboardData).items;
    let found = false;
    for (let i = 0; i < items.length; i++) {
      if (items[i].type.indexOf('image') !== -1) {
        found = true;
        processFile(items[i].getAsFile());
        e.preventDefault();
        break;
      }
    }
    if (!found) {
      setStatus('⚠ Tidak ada gambar di clipboard. Pastikan Anda sudah screenshot/copy gambar.', 'orange');
    }
  });

  // Drag & Drop
  zone.addEventListener('dragover', function(e) {
    e.preventDefault();
    zone.classList.add('dragover');
  });
  zone.addEventListener('dragleave', function() {
    zone.classList.remove('dragover');
  });
  zone.addEventListener('drop', function(e) {
    e.preventDefault();
    zone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      for (let i = 0; i < files.length; i++) processFile(files[i]);
    } else {
      // Coba ambil gambar dari dataTransfer.items (saat drag dari browser)
      const items = e.dataTransfer.items;
      for (let i = 0; i < items.length; i++) {
        if (items[i].type.indexOf('image') !== -1) {
          processFile(items[i].getAsFile());
          break;
        }
      }
    }
  });
</script>
"""


def render_upload_tab():
    st.subheader("1️⃣ Input Laporan Keuangan")

    method_tab1, method_tab2, method_tab3 = st.tabs([
        "📄 Upload PDF (Cerdas)",
        "🖼️ Upload Gambar",
        "📋 Paste Clipboard",
    ])

    # ─────────────────────────────────────────────────────────────────────────
    # TAB A: PDF — FITUR UTAMA BARU
    # ─────────────────────────────────────────────────────────────────────────
    with method_tab1:
        _render_pdf_tab()

    # ─────────────────────────────────────────────────────────────────────────
    # TAB B: Gambar
    # ─────────────────────────────────────────────────────────────────────────
    with method_tab2:
        st.caption("Format: PNG, JPG, JPEG. Bisa upload lebih dari satu sekaligus.")
        uploaded_files = st.file_uploader(
            "Pilih atau seret gambar ke sini",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key="uploader",
        )
        if uploaded_files:
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔍 Jalankan OCR untuk Semua", type="primary",
                             use_container_width=True, key="ocr_upload"):
                    run_ocr_on_files(uploaded_files)
            with col2:
                if st.button("🔄 Ulangi OCR (bersihkan hasil)", use_container_width=True,
                             key="reocr_upload"):
                    st.session_state.ocr_results = {}
                    run_ocr_on_files(uploaded_files)
            with st.expander("🖼️ Preview", expanded=False):
                cols = st.columns(min(3, len(uploaded_files)))
                for i, f in enumerate(uploaded_files):
                    with cols[i % len(cols)]:
                        st.image(f, caption=f.name, use_container_width=True)
        else:
            st.info("👆 Klik atau seret file gambar ke sini")

    # ─────────────────────────────────────────────────────────────────────────
    # TAB C: Paste clipboard
    # ─────────────────────────────────────────────────────────────────────────
    with method_tab3:
        _render_clipboard_section()


def _render_pdf_tab():
    """
    Tab PDF — pendekatan sederhana & reliable:
    1. Upload PDF → konversi SEMUA halaman ke gambar
    2. Tampilkan thumbnail semua halaman
    3. User pilih halaman mana yang mau diproses (centang)
    4. Tentukan jenis laporan & tahun per halaman
    5. Proses via Claude Vision atau OCR
    6. Support lebih dari 2 kolom tahun
    """
    st.markdown(
        "Upload file PDF laporan keuangan. Semua halaman akan ditampilkan "
        "sebagai thumbnail — pilih halaman yang berisi laporan yang ingin dikonversi."
    )

    if not PYPDF_AVAILABLE:
        st.error("pypdf belum terinstall: `pip install pypdf`")
        return

    # ── Upload ────────────────────────────────────────────────────────────
    pdf_files = st.file_uploader(
        "Pilih file PDF (bisa lebih dari satu)",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_uploader",
    )

    if not pdf_files:
        st.info("👆 Upload file PDF laporan keuangan")
        with st.expander("ℹ️ Cara penggunaan"):
            st.markdown(
                "1. Upload satu atau lebih file PDF\n"
                "2. Semua halaman ditampilkan sebagai thumbnail\n"
                "3. Centang halaman yang berisi laporan keuangan\n"
                "4. Pilih jenis laporan dan tahun pelaporan\n"
                "5. Klik **Proses Halaman Terpilih**\n\n"
                "**Fitur multi-tahun:** Jika satu halaman berisi data 3+ tahun "
                "(misal 2024, 2023, 2022), atur **Jumlah Kolom Nilai** di sidebar "
                "sesuai jumlah tahun yang ada."
            )
        return

    # ── Scan semua PDF → kumpulkan pages ─────────────────────────────────
    all_pages = []  # list of {fname, page_idx, image, page_text, years_found}

    for f in pdf_files:
        f.seek(0)
        pdf_bytes = f.read()
        f.seek(0)

        # Ekstrak teks untuk deteksi otomatis
        page_texts = pdf_extract_page_text(pdf_bytes)
        n_pages    = len(page_texts)

        st.caption(f"📄 {f.name} — {n_pages} halaman")

        # Konversi ke gambar
        with st.spinner(f"Mengkonversi {f.name} ke gambar..."):
            images = pdf_to_images(pdf_bytes, dpi=150)

        # Deteksi laporan per halaman
        report_pages = detect_report_pages(page_texts)
        # Buat reverse map: page_idx → rtype
        page_to_rtype = {}
        for rtype, idxs in report_pages.items():
            for idx in idxs:
                page_to_rtype[idx] = rtype

        for pg_idx, img in enumerate(images):
            pg_text = page_texts[pg_idx] if pg_idx < len(page_texts) else ""
            years   = detect_years_in_pages(page_texts, [pg_idx])
            auto_rtype = page_to_rtype.get(pg_idx, "unknown")

            all_pages.append({
                "fname":      f.name,
                "pdf_bytes":  pdf_bytes,
                "page_idx":   pg_idx,
                "image":      img,
                "page_text":  pg_text,
                "years_found": years,
                "auto_rtype":  auto_rtype,
            })

    if not all_pages:
        st.warning("Tidak ada halaman yang bisa dibaca dari PDF")
        return

    # ── Tampilkan thumbnail + kontrol per halaman ─────────────────────────
    st.divider()
    st.markdown(f"**Total: {len(all_pages)} halaman dari {len(pdf_files)} file PDF**")
    st.caption("Centang halaman yang berisi laporan keuangan yang ingin dikonversi:")

    # Simpan state pilihan halaman
    if "pdf_page_selections" not in st.session_state:
        st.session_state.pdf_page_selections = {}

    # Tombol select all / deselect untuk halaman yang terdeteksi otomatis
    col_auto, col_clear = st.columns(2)
    with col_auto:
        if st.button("✅ Pilih Semua Halaman Terdeteksi", use_container_width=True,
                     key="btn_select_detected"):
            for pg in all_pages:
                if pg["auto_rtype"] != "unknown":
                    key = f"{pg['fname']}_{pg['page_idx']}"
                    st.session_state.pdf_page_selections[key] = True
            st.rerun()
    with col_clear:
        if st.button("☐ Batal Semua Pilihan", use_container_width=True,
                     key="btn_deselect_all"):
            st.session_state.pdf_page_selections = {}
            st.rerun()

    st.divider()

    # Tampilkan thumbnail dalam grid 4 kolom
    N_COLS = 4
    page_configs = {}  # {key: {selected, rtype, years, target_years}}

    # Group per file
    files_seen = []
    for pg in all_pages:
        if pg["fname"] not in files_seen:
            files_seen.append(pg["fname"])

    for fname in files_seen:
        file_pages = [p for p in all_pages if p["fname"] == fname]
        st.markdown(f"##### 📄 {fname}")

        for row_start in range(0, len(file_pages), N_COLS):
            row_pages = file_pages[row_start:row_start + N_COLS]
            cols = st.columns(N_COLS)

            for col_idx, pg in enumerate(row_pages):
                pg_key = f"{pg['fname']}_{pg['page_idx']}"
                with cols[col_idx]:
                    # Thumbnail
                    st.image(pg["image"],
                             caption=f"Hal. {pg['page_idx']+1}",
                             use_container_width=True)

                    # Centang pilih
                    default_sel = (pg["auto_rtype"] != "unknown" or
                                   st.session_state.pdf_page_selections.get(pg_key, False))
                    selected = st.checkbox(
                        "Pilih",
                        value=st.session_state.pdf_page_selections.get(pg_key, default_sel),
                        key=f"chk_{pg_key}",
                    )
                    st.session_state.pdf_page_selections[pg_key] = selected

                    if selected:
                        # Jenis laporan
                        rtype_options = ["neraca", "laba_rugi", "arus_kas", "unknown"]
                        auto_idx = rtype_options.index(pg["auto_rtype"]) \
                                   if pg["auto_rtype"] in rtype_options else 3
                        rtype_sel = st.selectbox(
                            "Jenis",
                            options=rtype_options,
                            index=auto_idx,
                            key=f"rtype_{pg_key}",
                            label_visibility="collapsed",
                        )

                        # Tahun yang tersedia + pilih yang ingin diambil
                        years = pg["years_found"]
                        if years:
                            st.caption(f"Tahun: {', '.join(years)}")
                            # Multi-select tahun yang diambil
                            # Default: semua tahun (untuk mendukung >2 kolom)
                            target_years = st.multiselect(
                                "Ambil tahun",
                                options=years,
                                default=years,  # default: ambil semua
                                key=f"tyears_{pg_key}",
                                label_visibility="collapsed",
                            )
                        else:
                            st.caption("Tahun: belum terdeteksi")
                            target_years = []

                        page_configs[pg_key] = {
                            "page":         pg,
                            "rtype":        rtype_sel,
                            "target_years": target_years,
                        }

        st.divider()

    # ── Tombol proses ─────────────────────────────────────────────────────
    selected_keys = [k for k, v in st.session_state.pdf_page_selections.items() if v]
    n_selected    = len(selected_keys)

    if n_selected == 0:
        st.info("Belum ada halaman yang dipilih")
        return

    col_proc, col_reset = st.columns(2)
    with col_proc:
        need_api = (st.session_state.ocr_engine == "claude_vision"
                    and not st.session_state.claude_api_key)
        proc_btn = st.button(
            f"🚀 Proses {n_selected} Halaman Terpilih",
            type="primary",
            use_container_width=True,
            key="btn_proc_pages",
            disabled=need_api,
        )
        if need_api:
            st.caption("⚠ Isi API Key Claude di sidebar")

    with col_reset:
        if st.button("🗑️ Hapus Semua Hasil", use_container_width=True, key="btn_clr_pdf"):
            st.session_state.ocr_results  = {}
            st.session_state.pdf_results  = {}
            st.session_state.pdf_page_selections = {}
            st.rerun()

    if proc_btn and page_configs:
        _process_selected_pages(page_configs)


def _process_selected_pages(page_configs: dict):
    """
    Proses halaman-halaman PDF yang sudah dipilih user.

    Alur per halaman:
    1. Gambar sudah ada (dari thumbnail) → kirim ke Vision API / OCR
    2. Parse hasilnya → DataFrame
    3. Filter ke kolom tahun yang dipilih user (support >2 tahun)
    4. Simpan ke ocr_results dengan key {fname} › hal.{n} › {rtype}
    """
    engine    = st.session_state.ocr_engine
    api_key   = st.session_state.claude_api_key
    vis_model = st.session_state.claude_vision_model
    n_cols    = st.session_state.num_value_cols

    total    = len(page_configs)
    progress = st.progress(0, text="Memulai proses halaman PDF...")
    errors   = []

    for i, (pg_key, cfg) in enumerate(page_configs.items()):
        pg         = cfg["page"]
        rtype      = cfg["rtype"]
        t_years    = cfg["target_years"]
        fname      = pg["fname"]
        pg_idx     = pg["page_idx"]
        image      = pg["image"]
        all_years  = pg["years_found"]

        progress.progress(
            (i + 1) / total,
            text=f"Proses {fname} hal.{pg_idx+1} ({i+1}/{total})"
        )

        try:
            # ── Ekstraksi ────────────────────────────────────────────────
            # Coba teks langsung dulu jika bukan Claude Vision
            page_text = pg["page_text"]
            use_vision = (engine == "claude_vision" or len(page_text.strip()) < 100)

            if use_vision:
                raw, mode = ocr_image_smart(image, engine, api_key, vis_model)
            else:
                raw, mode = page_text, "ocr"

            if not raw or not raw.strip():
                errors.append(f"{fname} hal.{pg_idx+1}: ekstraksi kosong")
                continue

            # ── Parse ────────────────────────────────────────────────────
            if mode == "vision":
                df = parse_vision_json(raw, num_value_cols=n_cols)
                # Ambil years dari JSON — ini adalah sumber paling akurat
                try:
                    clean_json = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw.strip())
                    json_data  = json.loads(clean_json)
                    json_years = json_data.get("years", [])
                    if json_years:
                        all_years = json_years
                except Exception:
                    pass
                # Jika years dari JSON kosong, hitung dari kolom df
                if not all_years:
                    n_val = len([c for c in df.columns if c.startswith("Nilai_")])
                    all_years = [f"Tahun {i+1}" for i in range(n_val)]
            else:
                df = parse_ocr_text(raw, num_value_cols=n_cols)
                if not all_years:
                    all_years = detect_years(raw)
                # Sinkronkan all_years dengan jumlah kolom aktual di df
                n_val = len([c for c in df.columns if c.startswith("Nilai_")])
                if n_val > len(all_years):
                    # Ada lebih banyak kolom dari tahun yang terdeteksi
                    # Tambah placeholder
                    all_years = list(all_years) + [f"Tahun {i+1}" for i in range(len(all_years), n_val)]
                elif n_val < len(all_years):
                    all_years = all_years[:n_val]

            if df.empty:
                errors.append(f"{fname} hal.{pg_idx+1}: DataFrame kosong")
                continue

            # ── Filter tahun (HANYA jika user memilih subset) ────────────
            # Jika t_years = all_years (user pilih semua) → JANGAN filter,
            # langsung pakai df penuh untuk menghindari kehilangan data
            if t_years and set(t_years) != set(all_years):
                df_out, final_years = _filter_df_to_years(df, all_years, t_years)
            else:
                df_out, final_years = df, list(all_years)

            # ── Simpan ke ocr_results ────────────────────────────────────
            result_key = f"{fname} › hal.{pg_idx+1} › {rtype}"
            st.session_state.ocr_results[result_key] = {
                "text":  raw,
                "mode":  mode,
                "type":  rtype,
                "years": final_years,
                "df":    df_out,
            }

        except Exception as e:
            errors.append(f"{fname} hal.{pg_idx+1}: {e}")

    progress.empty()

    n_ok = total - len(errors)
    if n_ok > 0:
        st.success(f"✅ {n_ok}/{total} halaman berhasil diproses. Lihat di tab **Review & Edit**.")
    if errors:
        with st.expander(f"⚠ {len(errors)} halaman gagal"):
            for e in errors:
                st.caption(e)
    st.rerun()


def _filter_df_to_years(
    df: pd.DataFrame,
    all_years: list[str],
    target_years: list[str],
) -> tuple[pd.DataFrame, list[str]]:
    """
    Filter DataFrame ke kolom nilai yang sesuai dengan target_years.
    Mendukung lebih dari 2 kolom tahun.

    Logika:
    - all_years  = ['2024','2023','2022'] (urutan kolom di df: Nilai_1, Nilai_2, Nilai_3)
    - target_years = ['2023','2022'] → pertahankan Nilai_2 dan Nilai_3
    - Rename ke Nilai_1, Nilai_2, dst

    Jika target_years kosong atau tidak match → kembalikan df asli.
    """
    if not target_years or not all_years:
        return df, all_years

    value_cols = [c for c in df.columns if c.startswith("Nilai_")]
    if not value_cols:
        return df, all_years

    # Buat mapping: tahun → nama kolom
    # Asumsi: all_years[0] → Nilai_1, all_years[1] → Nilai_2, dst
    year_to_col = {}
    for i, yr in enumerate(all_years):
        col = f"Nilai_{i+1}"
        if col in value_cols:
            year_to_col[yr] = col

    # Kolom yang ingin dipertahankan (dalam urutan target_years)
    cols_to_keep = []
    final_years  = []
    for yr in target_years:
        if yr in year_to_col:
            cols_to_keep.append(year_to_col[yr])
            final_years.append(yr)

    if not cols_to_keep:
        return df, all_years

    # Buat df baru
    base_cols = [c for c in ["Tipe", "Keterangan", "Catatan"] if c in df.columns]
    df2 = df[base_cols + cols_to_keep].copy()

    # Rename ke Nilai_1, Nilai_2, ...
    rename_map = {old: f"Nilai_{i+1}" for i, old in enumerate(cols_to_keep)}
    df2 = df2.rename(columns=rename_map)

    return df2, final_years


# ---------------------------------------------------------------------------
# Tampilan visual zona paste (hanya dekoratif — komunikasi tetap via widget)
# ---------------------------------------------------------------------------
_PASTE_ZONE_HTML = """
<style>
* { box-sizing: border-box; font-family: Arial, sans-serif; }
#pz {
  border: 2.5px dashed #1F4E79;
  border-radius: 12px;
  padding: 28px 16px;
  text-align: center;
  background: #f0f4f9;
  transition: background .2s, border-color .2s;
  margin-bottom: 6px;
}
#pz.ready { background: #e8f0fe; border-color: #1a73e8; }
#pz.ok    { background: #d4edda; border-color: #198754; }
#pz.bad   { background: #f8d7da; border-color: #dc3545; }
.ico  { font-size: 2.2rem; }
.ttl  { font-size: 1rem; font-weight: bold; color: #1F4E79; margin: 4px 0; }
.sub  { font-size: 0.8rem; color: #666; }
#msg  { font-size: 0.84rem; margin-top: 6px; min-height: 18px; color: #333; }
#tbs  { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; justify-content: center; }
.tb   { max-width: 90px; max-height: 70px; object-fit: contain;
        border: 1px solid #ccc; border-radius: 4px; }
</style>

<div id="pz" tabindex="0"
     title="Klik di sini lalu Ctrl+V / ⌘V — atau drag-drop file PNG/JPG">
  <div class="ico">📋</div>
  <div class="ttl">Klik di sini → Ctrl+V / ⌘V</div>
  <div class="sub">Atau drag-drop file PNG/JPG langsung ke sini</div>
</div>
<div id="msg"></div>
<div id="tbs"></div>

<script>
(function(){
  const pz   = document.getElementById('pz');
  const msg  = document.getElementById('msg');
  const tbs  = document.getElementById('tbs');

  /* cari file_uploader Streamlit di parent frame berdasarkan data-testid */
  function getStreamlitUploader() {
    try {
      const frames = window.parent.document.querySelectorAll('input[type="file"]');
      /* ambil yang terakhir di-render (clipboard_uploader) */
      return frames.length ? frames[frames.length - 1] : null;
    } catch(e) { return null; }
  }

  function flash(cls, text) {
    ['ready','ok','bad'].forEach(c => pz.classList.remove(c));
    pz.classList.add(cls);
    msg.textContent = text;
    msg.style.color = cls === 'ok' ? '#198754' : cls === 'bad' ? '#dc3545' : '#1F4E79';
    if (cls !== 'ready') setTimeout(() => pz.classList.remove(cls), 2000);
  }

  function addThumb(src, name) {
    const img = document.createElement('img');
    img.src = src; img.className = 'tb'; img.title = name;
    tbs.appendChild(img);
  }

  /* ① Kirim file ke Streamlit file_uploader via DataTransfer trick */
  function injectToStreamlit(file) {
    const uploader = getStreamlitUploader();
    if (!uploader) return false;
    const dt = new DataTransfer();
    dt.items.add(file);
    uploader.files = dt.files;
    uploader.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }

  function handleFile(file) {
    if (!file || !file.type.startsWith('image/')) {
      flash('bad', '⚠ Bukan file gambar. Coba lagi.'); return;
    }
    /* baca sebagai data URL untuk thumbnail */
    const reader = new FileReader();
    reader.onload = ev => {
      addThumb(ev.target.result, file.name || 'paste.png');
      /* coba inject ke Streamlit uploader */
      const ok = injectToStreamlit(file);
      flash('ok', ok
        ? '✅ Berhasil! Klik "Tambah ke Antrian" di bawah.'
        : '✅ Gambar tertangkap. Gunakan widget upload di bawah.'
      );
    };
    reader.readAsDataURL(file);
  }

  /* fokus saat klik */
  pz.addEventListener('click', () => {
    pz.focus();
    flash('ready', '🎯 Siap — tekan Ctrl+V / ⌘V sekarang');
  });

  /* paste event */
  window.addEventListener('paste', function(e) {
    const items = (e.clipboardData || window.clipboardData || {}).items || [];
    let hit = false;
    for (let i = 0; i < items.length; i++) {
      if (items[i].kind === 'file' && items[i].type.startsWith('image/')) {
        hit = true;
        handleFile(items[i].getAsFile());
        e.preventDefault(); break;
      }
    }
    if (!hit) flash('bad', '⚠ Tidak ada gambar di clipboard. Screenshot dulu lalu Ctrl+V.');
  });

  /* drag-drop */
  ['dragenter','dragover'].forEach(ev =>
    pz.addEventListener(ev, e => { e.preventDefault(); pz.style.background='#cce5ff'; })
  );
  pz.addEventListener('dragleave', () => pz.style.background = '');
  pz.addEventListener('drop', function(e) {
    e.preventDefault(); pz.style.background = '';
    Array.from(e.dataTransfer.files).forEach(handleFile);
  });
})();
</script>
"""


def _render_clipboard_section():
    """
    Section paste clipboard yang BENAR.

    Arsitektur yang fixed:
    - Widget `st.file_uploader` (key='paste_uploader') adalah SUMBER KEBENARAN Python.
    - Komponen HTML hanya visual — saat user paste/drop, JS mencoba inject file
      ke uploader Streamlit via DataTransfer trick (bekerja di Chrome/Edge).
    - Jika inject JS gagal, user bisa drag-drop langsung ke widget di bawah.
    - Gambar masuk ke antrian SAAT WIDGET BERUBAH (on_change), BUKAN saat tombol ditekan.
    - Tidak ada st.rerun() di dalam helper — cukup update session_state, Streamlit
      otomatis rerun dari on_change callback.
    """
    st.markdown(
        "Copy gambar laporan keuangan (dari PDF, screenshot, scan) "
        "lalu **paste langsung di sini**."
    )

    # Zona visual paste
    components.html(_PASTE_ZONE_HTML, height=240, scrolling=False)

    st.caption(
        "**Cara pakai:** "
        "① Screenshot (Win+Shift+S / ⌘+Shift+4) atau copy gambar dari PDF viewer  "
        "② Klik zona abu-abu di atas  "
        "③ Tekan **Ctrl+V** / **⌘V**  "
        "④ Gambar muncul sebagai thumbnail  "
        "⑤ Klik **Tambah ke Antrian** di bawah"
    )

    st.divider()

    # Widget uploader yang BENAR-BENAR terhubung ke Python
    # Ini adalah single source of truth — gambar paste masuk dari sini
    st.markdown("##### 📥 Konfirmasi Gambar yang Akan Ditambahkan")
    st.caption(
        "Di Chrome/Edge: setelah paste di zona atas, gambar otomatis masuk ke sini. "
        "Atau drag-drop file PNG/JPG langsung ke widget ini."
    )

    paste_file = st.file_uploader(
        "Drop atau paste gambar di sini",
        type=["png", "jpg", "jpeg"],
        key="paste_uploader",
        label_visibility="collapsed",
    )

    if paste_file is not None:
        # Buat nama unik berdasarkan nama file asli + timestamp
        base = Path(paste_file.name).stem
        ext  = Path(paste_file.name).suffix or ".png"
        ts   = datetime.now().strftime("%H%M%S")
        fname = f"{base}_{ts}{ext}"

        # Preview langsung
        col_prev, col_btn = st.columns([2, 1])
        with col_prev:
            st.image(paste_file, caption=f"Preview: {paste_file.name}", use_container_width=True)
        with col_btn:
            st.write("")
            st.write("")
            # Tombol Tambah
            if st.button("➕ Tambah ke Antrian OCR", type="primary",
                         use_container_width=True, key="btn_add_paste"):
                _add_file_to_queue(paste_file, fname)

            st.write("")
            # Tombol OCR Langsung
            if st.button("⚡ Langsung OCR", use_container_width=True,
                         key="btn_direct_ocr"):
                _add_file_to_queue(paste_file, fname)
                # Jalankan OCR pada item terakhir (langsung dari bytes, tanpa cari di queue)
                _ocr_from_file_obj(paste_file, fname)
    else:
        st.info("👆 Paste (Ctrl+V) atau drag-drop file gambar ke zona di atas / widget ini")

    # -----------------------------------------------------------------------
    # Tampilkan antrian
    # -----------------------------------------------------------------------
    _render_queue_section()


def _add_file_to_queue(file_obj, filename: str):
    """
    Tambahkan file ke antrian. Baca bytes sekarang sebelum Streamlit reset widget.
    Simpan sebagai bytes di session_state (bukan PIL.Image) supaya tidak hilang
    saat rerun.
    """
    try:
        file_obj.seek(0)
        raw_bytes = file_obj.read()
        img = Image.open(io.BytesIO(raw_bytes)).convert("RGB")

        # Cegah duplikat nama
        existing = [p["filename"] for p in st.session_state.pasted_images]
        if filename in existing:
            stem = Path(filename).stem
            ext  = Path(filename).suffix
            filename = f"{stem}_{datetime.now().strftime('%f')}{ext}"

        st.session_state.pasted_images.append({
            "filename": filename,
            "image":    img,
            "bytes":    raw_bytes,
        })
        st.session_state.paste_counter += 1
        st.success(f"✅ '{filename}' ditambahkan ke antrian "
                   f"({len(st.session_state.pasted_images)} gambar total)")
    except Exception as e:
        st.error(f"Gagal menambahkan gambar: {e}")


def _ocr_from_file_obj(file_obj, filename: str):
    """OCR langsung dari file object (sebelum rerun)."""
    try:
        file_obj.seek(0)
        img = Image.open(io.BytesIO(file_obj.read())).convert("RGB")
        with st.spinner(f"OCR: {filename}..."):
            text = ocr_image(img, engine=st.session_state.ocr_engine)
        if not text.strip():
            st.warning(f"⚠ OCR tidak menghasilkan teks untuk {filename}")
            return
        st.session_state.ocr_results[filename] = {
            "text":  text,
            "type":  detect_report_type(text),
            "years": detect_years(text),
            "df":    parse_ocr_text(text, num_value_cols=st.session_state.num_value_cols),
        }
        st.success(f"✅ OCR selesai: {filename} (jenis: {st.session_state.ocr_results[filename]['type']})")
    except Exception as e:
        st.error(f"Gagal OCR {filename}: {e}")


def _render_queue_section():
    """Tampilkan antrian gambar paste dan tombol aksi."""
    st.divider()
    queue = st.session_state.pasted_images
    n = len(queue)
    st.markdown(f"#### 🖼️ Antrian Gambar: **{n} gambar**")

    if not queue:
        st.info("Belum ada gambar di antrian. Paste atau drag-drop gambar di atas.")
        return

    # Grid preview
    cols = st.columns(min(4, n))
    for i, item in enumerate(queue):
        with cols[i % len(cols)]:
            st.image(item["image"], caption=item["filename"], use_container_width=True)
            # Tombol hapus per item
            if st.button("🗑", key=f"del_{i}_{item['filename']}", help=f"Hapus {item['filename']}"):
                st.session_state.pasted_images.pop(i)
                st.rerun()

    # Tombol aksi massal
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("🔍 OCR Semua Antrian", type="primary",
                     use_container_width=True, key="btn_ocr_all"):
            _run_ocr_on_all_pasted()
    with c2:
        if st.button("🗑️ Kosongkan Antrian", use_container_width=True,
                     key="btn_clear_queue"):
            st.session_state.pasted_images = []
            st.rerun()
    with c3:
        # Download ZIP langsung (bukan via tombol dua langkah)
        zip_bytes = _build_zip_from_queue()
        if zip_bytes:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                "💾 Download ZIP",
                data=zip_bytes,
                file_name=f"gambar_lapkeu_{ts}.zip",
                mime="application/zip",
                use_container_width=True,
                key="btn_dl_zip",
            )


def _run_ocr_on_all_pasted():
    """OCR semua gambar di antrian paste."""
    queue = st.session_state.pasted_images
    if not queue:
        st.warning("Antrian kosong.")
        return
    progress = st.progress(0, text="Memulai OCR antrian...")
    total = len(queue)
    for i, item in enumerate(queue):
        fname = item["filename"]
        progress.progress((i + 1) / total, text=f"OCR: {fname} ({i+1}/{total})")
        if fname in st.session_state.ocr_results:
            continue
        try:
            img  = item["image"]
            text = ocr_image(img, engine=st.session_state.ocr_engine)
            if not text.strip():
                st.warning(f"⚠ OCR kosong: {fname}")
                continue
            st.session_state.ocr_results[fname] = {
                "text":  text,
                "type":  detect_report_type(text),
                "years": detect_years(text),
                "df":    parse_ocr_text(text, num_value_cols=st.session_state.num_value_cols),
            }
        except Exception as e:
            st.error(f"Gagal OCR {fname}: {e}")
    progress.empty()
    st.success(f"✅ OCR selesai untuk {total} gambar dari antrian")
    st.rerun()


def _build_zip_from_queue() -> bytes | None:
    """Buat ZIP dari semua gambar di antrian. Return bytes atau None jika kosong."""
    import zipfile
    queue = st.session_state.pasted_images
    if not queue:
        return None
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in queue:
            # Gunakan bytes asli jika ada, fallback re-encode ke PNG
            if "bytes" in item:
                zf.writestr(item["filename"], item["bytes"])
            else:
                img_buf = io.BytesIO()
                item["image"].save(img_buf, format="PNG")
                zf.writestr(item["filename"], img_buf.getvalue())
    buf.seek(0)
    return buf.getvalue()


def _export_pasted_as_zip():
    """Wrapper lama — dialihkan ke _build_zip_from_queue."""
    data = _build_zip_from_queue()
    if data:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button("⬇️ Download ZIP", data=data,
                           file_name=f"paste_lapkeu_{ts}.zip",
                           mime="application/zip", use_container_width=True)
    else:
        st.warning("Antrian kosong.")




def run_ocr_on_files(uploaded_files):
    """Jalankan OCR pada semua file yang belum diproses."""
    engine    = st.session_state.ocr_engine
    api_key   = st.session_state.get("claude_api_key", "")
    vis_model = st.session_state.get("claude_vision_model", "claude-haiku-4-5-20251001")
    n_cols    = st.session_state.num_value_cols

    progress = st.progress(0, text="Memulai ekstraksi...")
    total    = len(uploaded_files)

    for i, f in enumerate(uploaded_files):
        if f.name in st.session_state.ocr_results:
            progress.progress((i + 1) / total, text=f"Sudah diproses: {f.name}")
            continue

        label = "Claude Vision" if engine == "claude_vision" else "OCR"
        progress.progress((i + 1) / total, text=f"{label}: {f.name} ({i+1}/{total})")
        try:
            image = Image.open(f)
            raw, mode = ocr_image_smart(image, engine, api_key, vis_model)

            if not raw.strip():
                st.warning(f"⚠ Tidak ada teks untuk {f.name}")
                continue

            years = []
            report_type_text = raw

            if mode == "vision":
                df = parse_vision_json(raw, num_value_cols=n_cols)
                # Tahun dari JSON — sumber paling akurat
                try:
                    import json as _json
                    clean = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw.strip())
                    jdata = _json.loads(clean)
                    years = jdata.get("years", [])
                    # Gunakan keterangan untuk deteksi jenis laporan
                    report_type_text = " ".join(
                        r.get("keterangan", "") for r in jdata.get("rows", [])
                    )
                except Exception:
                    pass
                if not years:
                    n_val = len([c for c in df.columns if c.startswith("Nilai_")])
                    years = [f"Tahun {i+1}" for i in range(n_val)]
            else:
                df    = parse_ocr_text(raw, num_value_cols=n_cols)
                years = detect_years(raw)
                # Sinkronkan years dengan kolom aktual di df
                n_val = len([c for c in df.columns if c.startswith("Nilai_")])
                if n_val > len(years):
                    years = list(years) + [f"Tahun {j+1}" for j in range(len(years), n_val)]
                elif n_val < len(years):
                    years = years[:n_val]

            st.session_state.ocr_results[f.name] = {
                "text":  raw,
                "mode":  mode,
                "type":  detect_report_type(report_type_text),
                "years": years,
                "df":    df,
            }
        except Exception as e:
            st.error(f"Gagal memproses {f.name}: {e}")

    progress.empty()
    engine_label = "Claude Vision API" if engine == "claude_vision" else "OCR"
    st.success(f"✅ {engine_label} selesai untuk {total} file")
    st.rerun()


def render_review_tab():
    st.subheader("2️⃣ Review & Edit Hasil OCR")

    if not st.session_state.ocr_results:
        st.info("Belum ada hasil OCR. Silakan upload gambar dan jalankan OCR di tab 'Upload'.")
        return

    st.warning(
        "⚠️ **PENTING**: OCR lokal tidak 100% akurat. "
        "Wajib review setiap baris sebelum export. "
        "Angka yang besar paling rawan typo (misal '1.068.980' terbaca '1.068.960')."
    )

    for fname, data in st.session_state.ocr_results.items():
        with st.expander(f"📄 {fname} — Jenis: **{data['type']}**, Tahun: **{data['years']}**",
                         expanded=True):

            col1, col2, col3 = st.columns([2, 2, 3])
            with col1:
                new_type = st.selectbox(
                    "Jenis Laporan",
                    ["neraca", "laba_rugi", "arus_kas", "unknown"],
                    index=["neraca", "laba_rugi", "arus_kas", "unknown"].index(data["type"]),
                    key=f"type_{fname}",
                )
                st.session_state.ocr_results[fname]["type"] = new_type

            with col2:
                years_str = st.text_input(
                    "Tahun (pisahkan dengan koma)",
                    value=", ".join(data["years"]),
                    key=f"years_{fname}",
                )
                st.session_state.ocr_results[fname]["years"] = [
                    y.strip() for y in years_str.split(",") if y.strip()
                ]

            with col3:
                st.metric("Jumlah Baris Terekstrak", len(data["df"]))

            # Tabs: edit table, raw OCR
            tab_edit, tab_raw = st.tabs(["✏️ Edit Tabel", "📝 Raw OCR Text"])

            with tab_edit:
                st.caption("Edit langsung di tabel. Kolom 'Tipe' menentukan styling di Excel: "
                           "**section** (judul), **item** (baris data), **total** (formula SUM).")

                df_edit = data["df"].copy()

                # Konfigurasi kolom
                col_config = {
                    "Tipe": st.column_config.SelectboxColumn(
                        "Tipe",
                        options=["section", "item", "total", "empty"],
                        width="small",
                    ),
                    "Keterangan": st.column_config.TextColumn("Keterangan", width="large"),
                    "Catatan": st.column_config.TextColumn("Catatan", width="small"),
                }
                for col in df_edit.columns:
                    if col.startswith("Nilai_"):
                        col_config[col] = st.column_config.NumberColumn(
                            col, format="%.0f", width="medium"
                        )

                edited = st.data_editor(
                    df_edit,
                    column_config=col_config,
                    use_container_width=True,
                    num_rows="dynamic",
                    key=f"editor_{fname}",
                    height=400,
                )
                st.session_state.ocr_results[fname]["df"] = edited

            with tab_raw:
                st.text_area("Teks OCR mentah",
                             value=data["text"], height=300,
                             key=f"raw_{fname}")


def render_validation_tab():
    st.subheader("3️⃣ Validasi Silang")

    if not st.session_state.ocr_results:
        st.info("Belum ada data untuk divalidasi.")
        return

    # Gabungkan semua neraca menjadi satu (jika ada multiple)
    neraca_dfs = [d["df"] for d in st.session_state.ocr_results.values() if d["type"] == "neraca"]

    if not neraca_dfs:
        st.info("Tidak ada Laporan Posisi Keuangan (Neraca) untuk divalidasi.")
        st.caption("Validasi Aset = Liabilitas + Ekuitas hanya berlaku untuk Neraca.")
    else:
        combined = pd.concat(neraca_dfs, ignore_index=True)
        value_cols = [c for c in combined.columns if c.startswith("Nilai_")]

        st.markdown("### 💰 Validasi: Total Aset = Total Liabilitas + Ekuitas")
        results = validate_neraca(combined, value_cols)
        if results:
            st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)
        else:
            st.info("Tidak dapat menemukan Total Aset / Liabilitas / Ekuitas.")

        st.markdown("### 🧮 Validasi: Sub-total = Sum(Item) di atasnya")
        subtotal_issues = validate_subtotal_sums(combined, value_cols)
        if subtotal_issues:
            st.warning(f"Ditemukan {len(subtotal_issues)} ketidaksesuaian:")
            st.dataframe(pd.DataFrame(subtotal_issues), use_container_width=True, hide_index=True)
        else:
            st.success("✅ Semua sub-total konsisten dengan sum item-nya")

    # Validasi umum: multi-tahun compare
    st.markdown("### 📊 Perbandingan Multi-Tahun (Growth Analysis)")
    for fname, data in st.session_state.ocr_results.items():
        df = data["df"]
        value_cols = [c for c in df.columns if c.startswith("Nilai_")]
        if len(value_cols) < 2:
            continue

        compare_df = df[df["Tipe"] == "total"].copy()
        if compare_df.empty:
            continue

        v1, v2 = value_cols[0], value_cols[1]

        # Pastikan kolom numerik dan buang baris yang kedua kolomnya NaN semua
        compare_df[v1] = pd.to_numeric(compare_df[v1], errors="coerce")
        compare_df[v2] = pd.to_numeric(compare_df[v2], errors="coerce")
        compare_df = compare_df.dropna(subset=[v1, v2], how="all")

        if compare_df.empty:
            st.info(f"**{fname}**: tidak ada data total yang cukup untuk growth analysis.")
            continue

        # Growth absolut — aman meski salah satu NaN (hasil tetap NaN, tidak error)
        compare_df["Growth (Abs)"] = compare_df[v1] - compare_df[v2]

        # Growth % — hindari abs(None) dan div-by-zero
        def safe_growth_pct(row):
            a, b = row[v1], row[v2]
            if pd.isna(a) or pd.isna(b):
                return None
            if b == 0:
                return None  # tidak bisa hitung %, tampil sebagai "-"
            return round((a - b) / abs(b) * 100, 2)

        compare_df["Growth (%)"] = compare_df.apply(safe_growth_pct, axis=1)

        st.markdown(f"**{fname}** ({data['type']})")
        display_cols = ["Keterangan", v1, v2, "Growth (Abs)", "Growth (%)"]
        st.dataframe(
            compare_df[display_cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                v1: st.column_config.NumberColumn(v1, format="%.0f"),
                v2: st.column_config.NumberColumn(v2, format="%.0f"),
                "Growth (Abs)": st.column_config.NumberColumn(format="%.0f"),
                "Growth (%)": st.column_config.NumberColumn(format="%.2f%%"),
            },
        )


# Sinonim istilah akuntansi lama ↔ baru, supaya akun yang sama dengan
# penamaan berbeda antar tahun tetap dikenali sebagai satu akun.
_ACCOUNT_SYNONYMS = {
    "aktiva":    "aset",
    "kewajiban": "liabilitas",
    "hutang":    "utang",
}


def _normalize_keterangan(s) -> str:
    """
    Normalisasi nama akun untuk pencocokan antar tahun/dokumen:
    - Lowercase & rapikan spasi
    - Semua tanda baca dianggap spasi → "Piutang lain-lain" == "Piutang lain lain"
    - Buang token nomor catatan berpola angka+huruf di akhir ("Kas 2h" → "kas")
    - Samakan sinonim istilah lama/baru (aktiva→aset, kewajiban→liabilitas, dst.)
    """
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    t = str(s).lower().strip()
    # Semua non-alfanumerik jadi spasi (tanda hubung, koma, kurung, dll.)
    t = re.sub(r"[^a-z0-9]+", " ", t)
    words = [_ACCOUNT_SYNONYMS.get(w, w) for w in t.split()]
    # Buang trailing token nomor catatan berpola digit+huruf (mis. '2h', '12a').
    # Token angka murni TIDAK dibuang agar akun seperti "PPh pasal 21" vs
    # "PPh pasal 25" tidak salah tergabung.
    while len(words) > 1 and re.fullmatch(r"\d{1,3}[a-z]", words[-1]):
        words.pop()
    return " ".join(words)


def _extract_year_number(year_str) -> str | None:
    """
    Ekstrak HANYA angka tahun dari string yang bisa berbentuk:
      "2022"               → "2022"
      "30 Juni 2022"       → "2022"
      "31 Desember 2021"   → "2021"
      "Tahun 1"            → None (placeholder)

    Returns: string 4-digit tahun, atau None jika tidak ada.
    """
    if year_str is None:
        return None
    s = str(year_str).strip()
    # Cari pola 4-digit tahun (2000-2099)
    match = re.search(r"\b(20[0-3]\d)\b", s)
    if match:
        return match.group(1)
    return None


def _canonical_year(year_str) -> str:
    """
    Bentuk kanonik dari string tahun untuk pencocokan dan urutan.
    "30 Juni 2022" dan "31 Desember 2022" keduanya jadi "2022".
    Kalau tidak ada tahun terdeteksi, kembalikan string asli (lowercase).
    """
    yr = _extract_year_number(year_str)
    return yr if yr else str(year_str).strip().lower()


def _merge_reports_by_account(reports_list: list[dict]) -> dict:
    """
    Gabungkan beberapa laporan dengan pendekatan UNION BY ACCOUNT NAME.

    Algoritma (sesuai ide Pak Hendro):
    1. Urutkan dokumen dari tahun TERBARU → struktur/urutan akun mengikuti
       laporan terbaru (mis. neraca 2025 jadi kerangka utama).
    2. Kumpulkan SEMUA tahun unik dari semua laporan, normalisasi ke 4-digit
       (misal "30 Juni 2022" dan "31 Desember 2022" digabung jadi "2022")
       lalu urutkan menurun (terbaru → terlama).
    3. Untuk setiap laporan, untuk setiap baris:
       a. Cari akun dengan keterangan sama (dinormalisasi, per-section)
          di tabel master
       b. Jika ada → isi nilai di kolom tahun yang sesuai
       c. Jika tidak ada → SISIPKAN baris baru tepat setelah baris terakhir
          yang cocok, sehingga akun tetap berada di dalam section-nya
          (bukan ditumpuk di akhir tabel).

    Akun yang tidak ada di suatu tahun bernilai None di kolom tahun itu —
    gunakan _fill_missing_values_with_zero() untuk mengisinya dengan 0.

    Args:
        reports_list: list of dict dengan keys 'df' dan 'years'

    Returns:
        {'df': DataFrame gabungan, 'years': list tahun urut menurun}
    """
    reports_list = [r for r in reports_list if r and r.get("df") is not None]
    if not reports_list:
        return {"df": pd.DataFrame(), "years": []}

    if len(reports_list) == 1:
        return reports_list[0]

    # ── Langkah 1: dokumen tahun terbaru diproses lebih dulu ─────────────
    def _newest_year(r) -> int:
        yrs = []
        for y in r.get("years", []):
            n = _extract_year_number(y)
            if n:
                yrs.append(int(n))
        return max(yrs) if yrs else -1

    reports_sorted = sorted(reports_list, key=_newest_year, reverse=True)

    # ── Langkah 2: kumpulkan semua tahun unik (canonical) ───────────────
    # canonical_to_display: tahun_kanonik → label tampilan terbaik
    # contoh: "2022" → "30 Juni 2022" jika ditemukan, atau "2022" jika hanya itu
    canonical_to_display = {}
    for r in reports_sorted:
        for y in r.get("years", []):
            canon = _canonical_year(y)
            display = str(y)
            if canon not in canonical_to_display:
                canonical_to_display[canon] = display
            else:
                # Pilih label yang lebih lengkap (lebih panjang)
                if len(display) > len(canonical_to_display[canon]):
                    canonical_to_display[canon] = display

    # Urutkan tahun: numerik dulu (menurun), placeholder di belakang
    numeric_canons = sorted(
        [c for c in canonical_to_display.keys() if c.isdigit()],
        key=lambda x: int(x), reverse=True,
    )
    other_canons = sorted([c for c in canonical_to_display.keys() if not c.isdigit()])
    master_canons = numeric_canons + other_canons
    master_years_display = [canonical_to_display[c] for c in master_canons]

    if not master_canons:
        return reports_sorted[0]

    n_cols = len(master_canons)

    # ── Struktur master + dua indeks pencarian ───────────────────────────
    # key_index : (section_norm, akun_norm) → list of (master_idx, tipe)
    #             Pencocokan per-section supaya akun bernama sama di section
    #             berbeda (mis. "Lainnya" di Aset dan di Liabilitas) tidak
    #             salah tergabung.
    # norm_index: akun_norm → list of (master_idx, tipe)
    #             Fallback global — hanya dipakai jika nama akun UNIK,
    #             untuk kasus section header gagal terdeteksi di salah satu
    #             dokumen (perbedaan hasil OCR).
    master_rows = []
    key_index: dict = {}
    norm_index: dict = {}

    def _shift_index(index: dict, insert_pos: int):
        """Geser semua master_idx >= insert_pos setelah penyisipan baris."""
        for k in index:
            index[k] = [(i + 1 if i >= insert_pos else i, t) for i, t in index[k]]

    def find_master_idx(key, norm: str, tipe: str, used_this_doc: set) -> int | None:
        # Baris master yang sudah dipakai dokumen ini tidak boleh dipakai lagi:
        # dua baris berbeda dalam SATU dokumen harus tetap jadi dua baris.
        entries = [(i, t) for i, t in key_index.get(key, []) if i not in used_this_doc]
        # Prioritas 1: tipe sama persis
        for idx, t in entries:
            if t == tipe:
                return idx
        if tipe == "section":
            return None  # section hanya boleh match dengan section
        # Prioritas 2: tipe beda tapi bukan section (item ↔ total)
        for idx, t in entries:
            if t != "section":
                return idx
        # Prioritas 3 (fallback global): nama akun unik di seluruh master —
        # untuk kasus section header gagal terdeteksi di salah satu dokumen
        global_entries = [(i, t) for i, t in norm_index.get(norm, [])
                          if t != "section" and i not in used_this_doc]
        unique_idxs = {i for i, _ in global_entries}
        if len(unique_idxs) == 1:
            return next(iter(unique_idxs))
        return None

    # ── Langkah 3: proses setiap laporan ─────────────────────────────────
    for r in reports_sorted:
        df = r["df"]
        years_in_df = list(r.get("years", []))
        if df.empty or not years_in_df:
            continue

        val_cols = [c for c in df.columns if c.startswith("Nilai_")]

        # Mapping: kolom Nilai_X di df → index kolom di master (canonical)
        col_to_master_idx = {}
        for col_idx, col_name in enumerate(val_cols):
            if col_idx < len(years_in_df):
                yr_canon = _canonical_year(years_in_df[col_idx])
                if yr_canon in master_canons:
                    col_to_master_idx[col_name] = master_canons.index(yr_canon)

        cur_section = ""     # section aktif saat menyusuri baris dokumen ini
        last_pos = -1        # posisi baris master terakhir yang cocok (anchor)
        used_this_doc = set()  # baris master yang sudah di-klaim dokumen ini

        # Iterasi setiap baris di df
        for _, row in df.iterrows():
            ket   = row.get("Keterangan", "")
            cat   = row.get("Catatan", "")
            tipe  = row.get("Tipe", "item") or "item"
            norm  = _normalize_keterangan(ket)
            if not norm:
                continue

            if tipe == "section":
                key = ("__section__", norm)
                cur_section = norm
            else:
                key = (cur_section, norm)

            existing_idx = find_master_idx(key, norm, tipe, used_this_doc)

            if existing_idx is not None:
                # Akun sudah ada di master → isi/update kolom yang kosong
                master_row = master_rows[existing_idx]
                # Update catatan jika master masih kosong
                if not master_row.get("Catatan") and cat:
                    master_row["Catatan"] = cat
                # Update keterangan jika label di master pendek/UPPERCASE saja
                # (prefer Title Case dari dokumen lain)
                cur_ket = master_row.get("Keterangan", "")
                if cur_ket and cur_ket.isupper() and ket and not str(ket).isupper():
                    master_row["Keterangan"] = ket

                for src_col, master_col_idx in col_to_master_idx.items():
                    new_val = row.get(src_col)
                    if new_val is None or (isinstance(new_val, float) and pd.isna(new_val)):
                        continue
                    target_key = f"Nilai_{master_col_idx + 1}"
                    existing_val = master_row.get(target_key)
                    if existing_val is None or (isinstance(existing_val, float) and pd.isna(existing_val)):
                        master_row[target_key] = new_val
                    # Jika sudah terisi, biarkan first-wins (dokumen terbaru menang)
                last_pos = existing_idx
                used_this_doc.add(existing_idx)
            else:
                # Akun baru → SISIPKAN setelah anchor terakhir supaya tetap
                # berada di posisi section yang benar (bukan di akhir tabel)
                new_row = {
                    "Tipe":       tipe,
                    "Keterangan": ket,
                    "Catatan":    cat or "",
                }
                for k in range(n_cols):
                    new_row[f"Nilai_{k+1}"] = None
                for src_col, master_col_idx in col_to_master_idx.items():
                    new_val = row.get(src_col)
                    if new_val is not None and not (isinstance(new_val, float) and pd.isna(new_val)):
                        new_row[f"Nilai_{master_col_idx + 1}"] = new_val

                insert_pos = last_pos + 1
                master_rows.insert(insert_pos, new_row)
                _shift_index(key_index, insert_pos)
                _shift_index(norm_index, insert_pos)
                used_this_doc = {i + 1 if i >= insert_pos else i for i in used_this_doc}
                key_index.setdefault(key, []).append((insert_pos, tipe))
                norm_index.setdefault(norm, []).append((insert_pos, tipe))
                last_pos = insert_pos
                used_this_doc.add(insert_pos)

    # ── Langkah 4: bangun DataFrame final ────────────────────────────────
    if not master_rows:
        return {"df": pd.DataFrame(), "years": master_years_display}

    cols = ["Tipe", "Keterangan", "Catatan"] + [f"Nilai_{k+1}" for k in range(n_cols)]
    df_master = pd.DataFrame(master_rows, columns=cols)

    return {"df": df_master, "years": master_years_display}


def _fill_missing_values_with_zero(df: pd.DataFrame) -> pd.DataFrame:
    """
    Isi nilai kosong (None/NaN) dengan 0 pada baris 'item' dan 'total'.

    Ini mewujudkan aturan: akun yang tidak ada di suatu tahun tetap
    ditampilkan di laporan gabungan dengan nilai 0 untuk tahun tersebut
    (mis. "Piutang lain-lain" hanya ada di 2025 → kolom 2024 diisi 0).

    Baris 'section' (judul kelompok) dibiarkan kosong.
    """
    if df is None or df.empty:
        return df
    df = df.copy()
    val_cols = [c for c in df.columns if c.startswith("Nilai_")]
    if not val_cols:
        return df
    if "Tipe" in df.columns:
        mask = df["Tipe"].isin(["item", "total"])
    else:
        mask = pd.Series(True, index=df.index)
    for c in val_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")
        df.loc[mask & df[c].isna(), c] = 0
    return df


def _merge_reports_horizontal(existing: dict, new_data: dict) -> dict:
    """
    Wrapper untuk backward compatibility — sekarang menggunakan
    pendekatan union by account name yang lebih robust.
    """
    return _merge_reports_by_account([existing, new_data])


def render_export_tab():
    st.subheader("4️⃣ Export ke Excel")

    if not st.session_state.ocr_results:
        st.info("Belum ada data untuk di-export.")
        return

    # ── Kumpulkan semua laporan per jenis ─────────────────────────────────
    # reports_by_type[rtype] = list of {df, years} dari semua sumber
    # CATATAN: Laporan dengan jenis 'unknown' tetap diproses dan
    # digabung di sheet 'unknown' supaya Pak Hendro tetap bisa lihat
    # hasilnya (tidak hilang). Jika ingin disatukan ke neraca/laba_rugi/
    # arus_kas, edit jenisnya di tab Review.
    reports_by_type: dict[str, list[dict]] = {}
    for fname, data in st.session_state.ocr_results.items():
        rtype = data["type"] or "unknown"
        entry = {
            "df":    data["df"].copy(),
            "years": list(data["years"]),
            "source": fname,
        }
        reports_by_type.setdefault(rtype, []).append(entry)

    if not reports_by_type:
        st.warning("Belum ada hasil OCR yang bisa di-export.")
        return

    # ── Merge semua laporan per jenis sekaligus (union by account name) ──
    reports_raw = {}
    for rtype, entries in reports_by_type.items():
        merged = _merge_reports_by_account(entries)
        reports_raw[rtype] = merged

    # ── Opsi: isi akun yang tidak ada di suatu tahun dengan 0 ────────────
    fill_zero = st.checkbox(
        "Isi akun yang tidak ada di suatu tahun dengan angka 0",
        value=True,
        help="Contoh: 'Piutang lain-lain' hanya ada di neraca 2025 → "
             "pada kolom 2024 diisi 0 (bukan dikosongkan). "
             "Baris section (judul kelompok) tetap kosong.",
    )

    reports = {}
    for rtype, d in reports_raw.items():
        df_out = _fill_missing_values_with_zero(d["df"]) if fill_zero else d["df"]
        reports[rtype] = {"df": df_out, "years": d["years"]}

    # ── Ringkasan (dihitung dari data SEBELUM diisi 0, agar cakupan
    #    per tahun tetap menggambarkan data asli hasil ekstraksi) ─────────
    st.markdown("**Ringkasan yang akan di-export:**")
    summary = []
    for rtype, d in reports_raw.items():
        val_cols = [c for c in d["df"].columns if c.startswith("Nilai_")]
        filled_counts = []
        for vc in val_cols:
            n = int(d["df"][vc].notna().sum())
            filled_counts.append(str(n))
        n_sources = len(reports_by_type[rtype])
        summary.append({
            "Jenis":         rtype,
            "Sumber":        f"{n_sources} dokumen",
            "Total Baris":   len(d["df"]),
            "Tahun":         " | ".join(str(y) for y in d["years"]),
            "Terisi/Kolom":  " | ".join(filled_counts),
        })
    st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

    # ── Preview hasil gabungan (persis seperti yang akan di-export) ─────
    with st.expander("👀 Preview Hasil Gabungan Multi-Tahun", expanded=False):
        st.caption(
            "Semua akun dari semua tahun digabung. Akun yang tidak ada di "
            "suatu tahun " +
            ("diisi **0**." if fill_zero else "dibiarkan **kosong**.")
        )
        for rtype, d in reports.items():
            st.markdown(f"**{rtype}** — tahun: {', '.join(str(y) for y in d['years'])}")
            df_prev = d["df"].copy()
            # Ganti nama kolom Nilai_X → label tahun agar mudah dibaca
            rename_map = {}
            for i, y in enumerate(d["years"]):
                col = f"Nilai_{i+1}"
                if col in df_prev.columns:
                    rename_map[col] = str(y)
            df_prev = df_prev.rename(columns=rename_map)
            st.dataframe(df_prev, use_container_width=True, hide_index=True)

    # ── Info detil per kolom ──────────────────────────────────────────────
    with st.expander("📊 Detail Cakupan per Tahun", expanded=False):
        for rtype, d in reports_raw.items():
            st.markdown(f"**{rtype}**")
            val_cols = [c for c in d["df"].columns if c.startswith("Nilai_")]
            cov_data = []
            for i, vc in enumerate(val_cols):
                yr = d["years"][i] if i < len(d["years"]) else f"Kolom {i+1}"
                filled = int(d["df"][vc].notna().sum())
                total = len(d["df"])
                cov_data.append({
                    "Tahun":   yr,
                    "Terisi":  filled,
                    "Total":   total,
                    "Persen":  f"{filled/total*100:.0f}%" if total > 0 else "0%",
                })
            st.dataframe(pd.DataFrame(cov_data), use_container_width=True, hide_index=True)

    # ── Peringatan kolom kosong (informatif, bukan error) ────────────────
    for rtype, d in reports_raw.items():
        val_cols = [c for c in d["df"].columns if c.startswith("Nilai_")]
        if len(val_cols) >= 2:
            empty_cols = [d["years"][i] for i, vc in enumerate(val_cols)
                         if i < len(d["years"]) and d["df"][vc].notna().sum() == 0]
            if empty_cols:
                st.warning(
                    f"⚠ **{rtype}**: kolom tahun **{', '.join(str(y) for y in empty_cols)}** "
                    f"kosong. Kemungkinan tidak ada akun yang cocok namanya antara dokumen-dokumen "
                    f"yang Anda upload. Periksa di tab Review apakah keterangan akun sudah benar."
                )

    if st.button("📥 Generate File Excel", type="primary", use_container_width=True):
        with st.spinner("Membuat file Excel..."):
            excel_bytes = build_excel(reports, entity_name=st.session_state.entity_name)

        st.success("✅ File Excel siap!")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = re.sub(r"[^\w\-]", "_", st.session_state.entity_name)[:30]
        filename = f"Laporan_Keuangan_{safe_name}_{timestamp}.xlsx"

        st.download_button(
            "⬇️ Download Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# =============================================================================
# MAIN
# =============================================================================

def main():
    init_session_state()

    st.title("📊 Konverter Laporan Keuangan")
    st.caption("PDF & Gambar → Excel terstruktur | Claude Vision API | v2.0")

    render_sidebar()

    tab1, tab2, tab3, tab4 = st.tabs([
        "📤 Upload & OCR",
        "✏️ Review & Edit",
        "✅ Validasi",
        "📥 Export Excel",
    ])

    with tab1:
        render_upload_tab()
    with tab2:
        render_review_tab()
    with tab3:
        render_validation_tab()
    with tab4:
        render_export_tab()


if __name__ == "__main__":
    main()
