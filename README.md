# Wedding Name Tag Generator

Generate PDF name tag undangan pernikahan dari file Excel secara otomatis.
Layout: 210x152mm landscape, 3x4 grid (12 name tag per halaman), motif bidadari & bunga.

---

## Struktur Folder

```
name-label-generator/
├── generate_nametag.py                <- script utama
├── img_template.png                   <- template motif sudut (wajib ada)
├── CinzelDecorative-Regular.otf       <- font (OTF, auto-convert ke TTF)
├── CinzelDecorative-Regular.ttf       <- font hasil convert (auto-generate)
├── corners/                           <- hasil crop sudut (auto-generate)
│   ├── corner_tl.png
│   ├── corner_tr.png
│   ├── corner_bl.png
│   └── corner_br.png
├── input/
│   └── *.xlsx                         <- file Excel tamu (taruh di sini)
├── output/
│   └── *.pdf                          <- hasil PDF (auto-generate)
└── README.md
```

---

## Prasyarat

- Python 3.10 atau lebih baru
- File `img_template.png` (template motif sudut) di folder root
- File `CinzelDecorative-Regular.otf` di folder root (atau koneksi internet untuk download)

---

## 1. Install Dependencies

```bash
pip install pandas openpyxl reportlab requests fonttools otf2ttf Pillow
```

---

## 2. Siapkan File Excel

Format file Excel yang diperlukan:

| (baris 1 — kosong atau judul) |
|---|
| **Nama Lengkap** |
| Budi Santoso |
| Siti Rahayu |
| ... |

- Nama file bebas (boleh apa saja), ekstensi `.xlsx` atau `.xls`
- Header kolom (`Nama Lengkap`) harus ada di **baris ke-2**
- Taruh file Excel di folder **`input/`** (boleh lebih dari satu file)

> Nama duplikat dibiarkan apa adanya, tidak dihapus.

---

## 3. Jalankan Script

```bash
python generate_nametag.py
```

Script akan scan folder `input/` dan meminta kamu memilih file jika ada lebih dari satu:

```
File Excel yang tersedia di folder input/:
  [1] Daftar_Tamu_Undangan.xlsx
  [2] Tamu_VIP.xlsx

Pilih nomor file (1-2): 1
Dipilih: Daftar_Tamu_Undangan.xlsx
Cropping corners from template image...
Corners cached in corners/
Font CinzelDecorative registered from CinzelDecorative-Regular.ttf.
Loaded 76 names from input/Daftar_Tamu_Undangan.xlsx
PDF saved: output/Daftar_Tamu_Undangan.pdf  (7 page(s), 76 names)
Done! Buka output/Daftar_Tamu_Undangan.pdf untuk preview.
```

> Jika hanya ada 1 file di `input/`, langsung dipilih otomatis tanpa perlu input.

---

## 4. Hasil Output

File PDF otomatis tersimpan di folder **`output/`**, nama mengikuti file Excel.

- Setiap halaman berisi **12 name tag** (3 kolom x 4 baris)
- Jika jumlah nama tidak habis dibagi 12, slot kosong di halaman terakhir dibiarkan
- PDF bisa dibuka di Adobe Reader, browser, atau aplikasi PDF apapun

### Format tiap name tag:

```
┌───────────────────────────────┐
│ [bidadari]      [bidadari]    │
│                               │
│        Agus & Partner         │
│       ─────────────────       │
│              di               │
│            Tempat             │
│                               │
│ [bunga]            [bunga]    │
└───────────────────────────────┘
  Background putih, teks hitam
  Font CinzelDecorative
```

- **Baris 1:** Nama tamu (auto-fit 11pt → 7pt)
- **Baris 2:** "di"
- **Baris 3:** "Tempat"
- **Sudut atas:** motif bidadari dari template
- **Sudut bawah:** motif bunga dari template

---

## 5. Pengaturan Print

Saat print di printer:

- **Ukuran kertas:** Custom — `210mm x 152mm`
- **Orientasi:** Landscape
- **Scaling:** Actual size / 100% (jangan fit-to-page)
- **Margin printer:** 0mm (borderless) atau sesuaikan

---

## Troubleshooting

| Error | Penyebab | Solusi |
|---|---|---|
| `ModuleNotFoundError` | Library belum terinstall | `pip install pandas openpyxl reportlab requests fonttools otf2ttf Pillow` |
| `Tidak ada file Excel di folder 'input/'` | Folder input kosong | Taruh file `.xlsx` ke folder `input/` |
| `Kolom 'Nama Lengkap' tidak ditemukan` | Header salah | Pastikan header persis `Nama Lengkap` di baris ke-2 |
| `img_template.png not found` | Template motif tidak ada | Taruh file `img_template.png` di folder root |
| Font fallback ke Helvetica | OTF tidak ada & gagal download | Taruh `CinzelDecorative-Regular.otf` di folder root |
| Layout geser saat print | Setting print salah | Set kertas Custom 210x152mm, scaling 100% |

---

## Spesifikasi Teknis

| Item | Detail |
|---|---|
| Ukuran halaman | 210mm x 152mm (landscape) |
| Grid | 3 kolom x 4 baris = 12 name tag/halaman |
| Ukuran name tag | 67mm x 36mm |
| Gap antar tag | 2mm horizontal & vertical |
| Margin halaman | 2mm semua sisi |
| Font | CinzelDecorative Regular, 7-11pt (auto-fit) |
| Warna background | #FFFFFF (putih) |
| Warna teks & border | #000000 (hitam) |
| Ornamen sudut | Crop dari img_template.png (16mm x 14mm per sudut) |
| Format teks | Nama / di / Tempat (3 baris centered) |

---

*Stack: Python + ReportLab + Pillow + CinzelDecorative Font*
