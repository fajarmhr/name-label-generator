# Wedding Name Tag Generator

Generate PDF name tag undangan pernikahan dari file Excel secara otomatis.
Layout: A4 portrait, 3x8 grid (24 name tag per halaman), motif corner dari template gambar.

---

## Struktur Folder

```
name-label-generator/
├── generate_nametag.py                    <- script utama
├── img_template.png                       <- template motif sudut (wajib ada)
├── fonts/
│   ├── PlusJakartaSans-Medium.otf         <- font Bold (untuk nama tamu)
│   ├── PlusJakartaSans-Regular.otf        <- font Regular (untuk alamat / "di" / "Tempat")
│   ├── PlusJakartaSans-Medium.ttf         <- hasil convert (auto-generate)
│   └── PlusJakartaSans-Regular.ttf        <- hasil convert (auto-generate)
├── corners/                               <- hasil crop sudut (auto-generate)
│   ├── corner_tl.png
│   ├── corner_tr.png
│   ├── corner_bl.png
│   └── corner_br.png
├── input/
│   ├── Template Daftar Tamu Undangan.xlsx <- template referensi format Excel
│   └── *.xlsx                             <- file Excel tamu (taruh di sini)
├── output/
│   └── *.pdf                              <- hasil PDF (auto-generate)
├── .gitignore
└── README.md
```

---

## Prasyarat

- Python 3.10 atau lebih baru
- File `img_template.png` (template motif sudut) di folder root
- File font `.otf` di folder `fonts/` (PlusJakartaSans-Medium & PlusJakartaSans-Regular)

---

## 1. Install Dependencies

```bash
pip install pandas openpyxl reportlab fonttools otf2ttf Pillow
```

---

## 2. Siapkan File Excel

Lihat file `input/Template Daftar Tamu Undangan.xlsx` untuk contoh format.

Format yang diperlukan:

| (baris 1 — kosong atau judul) | | |
|---|---|---|
| **No** | **Nama Lengkap** | **Alamat** |
| 1 | Budi Santoso | SMP Negeri 1 Ampel |
| 2 | Siti Rahayu | |
| ... | ... | ... |

- Nama file bebas, ekstensi `.xlsx` atau `.xls`
- Header kolom harus ada di **baris ke-2**
- Kolom `Nama Lengkap` wajib ada
- Kolom `Alamat` opsional — jika ada dan terisi, alamat ditampilkan di name tag
- Taruh file Excel di folder **`input/`** (boleh lebih dari satu file)

---

## 3. Jalankan Script

```bash
python generate_nametag.py
```

Script akan scan folder `input/` dan meminta kamu memilih file jika ada lebih dari satu:

```
File Excel yang tersedia di folder input/:
  [1] Daftar Tamu Undangan - Bapak.xlsx
  [2] Daftar Tamu Undangan - Ibu.xlsx

Pilih nomor file (1-2): 1
Dipilih: Daftar Tamu Undangan - Bapak.xlsx
Font Bold registered dari fonts/PlusJakartaSans-Medium.ttf
Font Regular registered dari fonts/PlusJakartaSans-Regular.ttf
Cropping corners from template image...
Corners cached in corners/
Loaded 76 names from input/Daftar Tamu Undangan - Bapak.xlsx
PDF saved: output/Daftar Tamu Undangan - Bapak.pdf  (4 page(s), 76 names)
Done! Buka output/Daftar Tamu Undangan - Bapak.pdf untuk preview.
```

> Jika hanya ada 1 file di `input/`, langsung dipilih otomatis tanpa perlu input.

---

## 4. Hasil Output

File PDF otomatis tersimpan di folder **`output/`**, nama mengikuti file Excel.

- Setiap halaman berisi **24 name tag** (3 kolom x 8 baris)
- Jika jumlah nama tidak habis dibagi 24, slot kosong di halaman terakhir dibiarkan
- PDF bisa dibuka di Adobe Reader, browser, atau aplikasi PDF apapun

### Format tiap name tag:

**Jika ada alamat:**
```
┌───────────────────────────────┐
│ [corner]          [corner]    │
│                               │
│        Agus & Partner         │
│       ─────────────────       │
│      SMP Negeri 2 Ampel       │
│                               │
│ [corner]          [corner]    │
└───────────────────────────────┘
```

**Jika alamat kosong:**
```
┌───────────────────────────────┐
│ [corner]          [corner]    │
│                               │
│        Agus & Partner         │
│       ─────────────────       │
│              di               │
│            Tempat             │
│                               │
│ [corner]          [corner]    │
└───────────────────────────────┘
```

- **Baris 1:** Nama tamu — font **Medium/Bold** (auto-fit 13.5pt → 9pt)
- **Baris 2-3:** Alamat (jika ada) atau "di" + "Tempat" (jika kosong)
- **Ornamen sudut:** Auto-crop dari `img_template.png`

---

## 5. Pengaturan Print

Saat print di printer:

- **Ukuran kertas:** A4 (210mm x 297mm)
- **Orientasi:** Portrait
- **Scaling:** Actual size / 100% (jangan fit-to-page)
- **Margin printer:** 0mm (borderless) atau sesuaikan

---

## 6. Ganti Template Motif

Cukup ganti file `img_template.png` dengan gambar baru, lalu jalankan script ulang.
Corner otomatis di-crop ulang dari template baru (auto-detect perubahan).

---

## Troubleshooting

| Error | Penyebab | Solusi |
|---|---|---|
| `ModuleNotFoundError` | Library belum terinstall | `pip install pandas openpyxl reportlab fonttools otf2ttf Pillow` |
| `Tidak ada file Excel di folder 'input/'` | Folder input kosong | Taruh file `.xlsx` ke folder `input/` |
| `Kolom 'Nama Lengkap' tidak ditemukan` | Header salah | Pastikan header persis `Nama Lengkap` di baris ke-2 |
| `img_template.png not found` | Template motif tidak ada | Taruh file `img_template.png` di folder root |
| Font fallback ke Helvetica | OTF tidak ada di folder fonts/ | Taruh file `.otf` ke folder `fonts/` |

---

## Spesifikasi Teknis

| Item | Detail |
|---|---|
| Ukuran halaman | A4 — 210mm x 297mm (portrait) |
| Grid | 3 kolom x 8 baris = 24 name tag/halaman |
| Ukuran name tag | 64mm x 33.9mm |
| Gap antar tag | otomatis (sisa ruang dibagi rata) |
| Margin halaman | 7mm horizontal, 8mm vertikal |
| Font nama | PlusJakartaSans Medium, 9-13.5pt (auto-fit) |
| Font alamat/di/Tempat | PlusJakartaSans Regular, 9pt |
| Warna background | #FFFFFF (putih) |
| Warna teks & border | #000000 (hitam) |
| Ornamen sudut | Auto-crop dari img_template.png |
| Format teks | Nama + alamat (2 baris) atau Nama + di + Tempat (3 baris) |

---

*Stack: Python + ReportLab + Pillow + PlusJakartaSans Font*
