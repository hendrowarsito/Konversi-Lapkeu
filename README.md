# Konverter Laporan Keuangan (Image → Excel)

Aplikasi Streamlit untuk mengkonversi gambar laporan keuangan Indonesia (Neraca, Laba Rugi, Arus Kas) menjadi file Excel terstruktur menggunakan OCR lokal gratis.

## Fitur

- **OCR Lokal Gratis**: Tesseract + EasyOCR (tanpa biaya API)
- **Auto-detect Jenis Laporan**: Neraca / Laba Rugi / Arus Kas
- **Multi-tahun**: Ekstraksi otomatis 2 tahun pelaporan (current vs prior)
- **Preview & Edit**: Tabel editable sebelum export (koreksi manual hasil OCR)
- **Validasi Silang**:
  - Total Aset = Total Liabilitas + Ekuitas
  - Sub-total = Sum(item) di atasnya
  - Growth analysis (% perubahan antar tahun)
- **Export Excel**: Multi-sheet, dengan formula SUM (bukan hardcode), format angka Indonesia

## Instalasi

### 1. Install Python packages

```bash
pip install -r requirements.txt
```

### 2. Install Tesseract OCR (system-level)

**Windows:**
1. Download installer: https://github.com/UB-Mannheim/tesseract/wiki
2. Install dengan memilih bahasa tambahan: **Indonesian (ind)**
3. Tambahkan ke PATH (biasanya `C:\Program Files\Tesseract-OCR`)
4. Verifikasi: `tesseract --version`

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get update
sudo apt-get install tesseract-ocr tesseract-ocr-ind tesseract-ocr-eng
```

**macOS:**
```bash
brew install tesseract tesseract-lang
```

### 3. EasyOCR (opsional, fallback)

EasyOCR akan otomatis download model saat pertama kali dijalankan (~100 MB).
Tidak perlu install system-level apapun.

## Cara Menjalankan

```bash
streamlit run app.py
```

Aplikasi akan terbuka di browser pada `http://localhost:8501`.

## Workflow Penggunaan

1. **Tab Upload & OCR**: Upload gambar laporan keuangan (PNG/JPG), lalu klik "Jalankan OCR"
2. **Tab Review & Edit**: Koreksi hasil OCR. Pastikan:
   - Jenis laporan terdeteksi dengan benar
   - Tahun pelaporan benar
   - Angka-angka yang besar (trillun rupiah) ter-ekstrak dengan benar
   - Kolom "Tipe" sesuai: section (judul), item (data), total (formula SUM)
3. **Tab Validasi**: Cek validasi silang otomatis
4. **Tab Export**: Generate & download file Excel

## Tips Akurasi OCR

- Gunakan gambar resolusi tinggi (minimal 300 DPI)
- Gambar scan/foto yang rapi (tidak miring, pencahayaan merata)
- Untuk gambar kualitas rendah, gunakan EasyOCR
- Untuk gambar kualitas tinggi, Tesseract lebih cepat

## Keterbatasan

- Akurasi OCR lokal umumnya 70-85%, tergantung kualitas gambar
- **WAJIB review manual** sebelum digunakan untuk keperluan profesional
- Format layout kompleks (multi-kolom) mungkin perlu edit manual
- Bahasa: optimal untuk laporan keuangan bahasa Indonesia standar PSAK

## Struktur Output Excel

```
Workbook.xlsx
├── Info (ringkasan & status kelengkapan)
├── Laporan Posisi Keuangan
├── Laporan Laba Rugi
└── Laporan Arus Kas
```

Setiap sheet:
- Header dengan gradasi warna biru korporat
- Kolom: Keterangan | Catatan | Nilai per tahun
- Format angka: `#,##0;(#,##0);"-"` (Indonesia, negatif dalam kurung)
- Total pakai formula `=SUM(...)` — bisa diupdate kalau data diedit
- Freeze panes di header, gridline disembunyikan
