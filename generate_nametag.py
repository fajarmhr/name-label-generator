# INSTALL: pip install pandas openpyxl reportlab fonttools otf2ttf Pillow
# RUN    : python generate_nametag.py
# OUTPUT : output/<nama_file>.pdf

import os
import sys
import math
import pandas as pd
from PIL import Image
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.colors import HexColor

# ── Constants ──────────────────────────────────────────────────────────────────

INPUT_DIR     = "input"
OUTPUT_DIR    = "output"
CORNER_DIR    = "corners"
TEMPLATE_IMG  = "img_template.png"
FONT_DIR      = "fonts"

# Font Regular — untuk teks "di" dan "Tempat"
FONT_REG_OTF  = os.path.join(FONT_DIR, "PlusJakartaSans-Regular.otf")
FONT_REG_TTF  = os.path.join(FONT_DIR, "PlusJakartaSans-Regular.ttf")

# Font Bold — untuk nama tamu
FONT_BOLD_OTF = os.path.join(FONT_DIR, "PlusJakartaSans-Medium.otf")
FONT_BOLD_TTF = os.path.join(FONT_DIR, "PlusJakartaSans-Medium.ttf")

# Page: Custom (200mm x 137mm)
PAGE_W = 200 * mm
PAGE_H = 137 * mm

COLS     = 3
ROWS     = 4
TAGS_PER = COLS * ROWS   # 12 labels per halaman

# Label: 64mm x 32mm
TAG_W = 64 * mm
TAG_H = 32 * mm

# Margin halaman (pinggir kertas)
PAGE_MARGIN_H = 2 * mm    # margin kiri & kanan (side margin)
PAGE_MARGIN_V = 2 * mm    # margin atas & bawah (top margin)

# Gap antar nametag — dihitung otomatis dari sisa ruang setelah margin & label
# Margin 2mm fix, sisa ruang dibagi rata jadi gap
GAP_H = (PAGE_W - 2 * PAGE_MARGIN_H - COLS * TAG_W) / (COLS - 1)   # = 2mm
GAP_V = (PAGE_H - 2 * PAGE_MARGIN_V - ROWS * TAG_H) / (ROWS - 1)   # = ~1.67mm

# Ukuran ornamen sudut pada nametag
# Atas (bunga kecil) — lebih kecil
CORNER_TOP_W = 18 * mm
CORNER_TOP_H = 11.5 * mm
# Bawah (gunungan + bunga) — lebih besar
CORNER_BOT_W = 19 * mm
CORNER_BOT_H = 12 * mm

# Jarak ornamen dari tepi nametag (mm)
CORNER_MARGIN = 1.1 * mm

# Colors — monochrome black on white
COLOR_BG           = HexColor("#FFFFFF")
COLOR_BORDER_OUTER = HexColor("#000000")
COLOR_TEXT         = HexColor("#000000")
COLOR_LINE         = HexColor("#000000")

# Font sizes
FONT_MAX  = 13.5
FONT_MIN  = 9
FONT_SUB  = 9     # ukuran font tetap untuk teks di bawah garis (alamat / "di" / "Tempat")
FONT_NAME_BOLD = "PlusJakartaSans-Medium"   # font untuk nama tamu
FONT_NAME_REG  = "PlusJakartaSans-Regular"  # font untuk "di" dan "Tempat"


# ── 1. Pick Input File ────────────────────────────────────────────────────────

def pick_input_file() -> str:
    """Scan folder input/, tampilkan daftar Excel, minta user pilih."""
    exts = (".xlsx", ".xls")
    files = sorted(
        f for f in os.listdir(INPUT_DIR)
        if f.lower().endswith(exts)
    )

    if not files:
        print(f"ERROR: Tidak ada file Excel di folder '{INPUT_DIR}/'.")
        print("Taruh file .xlsx ke sana lalu jalankan ulang script ini.")
        sys.exit(1)

    if len(files) == 1:
        chosen = files[0]
        print(f"File ditemukan: {chosen}")
        return os.path.join(INPUT_DIR, chosen)

    print("\nFile Excel yang tersedia di folder input/:")
    for i, name in enumerate(files, start=1):
        print(f"  [{i}] {name}")

    while True:
        raw = input(f"\nPilih nomor file (1-{len(files)}): ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(files):
            chosen = files[int(raw) - 1]
            print(f"Dipilih: {chosen}")
            return os.path.join(INPUT_DIR, chosen)
        print(f"  Input tidak valid. Masukkan angka 1 sampai {len(files)}.")


# ── 2. Prepare Font ──────────────────────────────────────────────────────────

def _is_valid_ttf(path: str) -> bool:
    try:
        with open(path, "rb") as f:
            sig = f.read(4)
        return sig in (b'\x00\x01\x00\x00', b'true')
    except OSError:
        return False


def _convert_otf_to_ttf(otf_path: str, ttf_path: str) -> bool:
    try:
        saved_argv = sys.argv
        sys.argv = ["otf2ttf", otf_path]
        print(f"Converting {otf_path} → TTF via otf2ttf...")
        from otf2ttf import main as otf2ttf_main
        otf2ttf_main()
        sys.argv = saved_argv
        expected = os.path.splitext(otf_path)[0] + ".ttf"
        if expected != ttf_path and os.path.exists(expected):
            os.replace(expected, ttf_path)
        if _is_valid_ttf(ttf_path):
            print(f"Converted OK → {ttf_path}")
            return True
    except Exception as e:
        print(f"  OTF→TTF conversion failed: {e}")
    return False


def _prepare_single_font(otf_path: str, ttf_path: str, label: str) -> str | None:
    """
    Siapkan satu font: cek TTF → convert dari OTF → return path atau None.
    label dipakai untuk pesan log (misal "Regular", "Bold").
    """
    # Kalau TTF sudah ada dan valid, langsung pakai
    if os.path.exists(ttf_path) and _is_valid_ttf(ttf_path):
        return ttf_path

    # TTF ada tapi rusak → hapus
    if os.path.exists(ttf_path) and not _is_valid_ttf(ttf_path):
        print(f"Removing invalid {ttf_path}...")
        os.remove(ttf_path)

    # Convert dari OTF kalau ada
    if os.path.exists(otf_path):
        if _convert_otf_to_ttf(otf_path, ttf_path):
            return ttf_path

    print(f"WARNING: Font {label} tidak ditemukan. Taruh file OTF di folder fonts/.")
    return None


def prepare_fonts() -> dict[str, str | None]:
    """
    Siapkan font Regular dan Bold.
    Returns dict: {"regular": path_or_None, "bold": path_or_None}
    """
    os.makedirs(FONT_DIR, exist_ok=True)
    return {
        "regular": _prepare_single_font(FONT_REG_OTF, FONT_REG_TTF, "Regular"),
        "bold":    _prepare_single_font(FONT_BOLD_OTF, FONT_BOLD_TTF, "Bold"),
    }


# ── 3. Prepare Corner Images ─────────────────────────────────────────────────

def prepare_corner_images() -> dict[str, str] | None:
    """
    Crop 4 corners from img_template.png, make white→transparent, cache as PNG.
    Returns dict: {'TL': path, 'TR': path, 'BL': path, 'BR': path} or None.
    """
    if not os.path.exists(TEMPLATE_IMG):
        print(f"WARNING: {TEMPLATE_IMG} not found. Corners will be skipped.")
        return None

    os.makedirs(CORNER_DIR, exist_ok=True)

    # Check if already cached
    names = {"TL": "corner_tl.png", "TR": "corner_tr.png",
             "BL": "corner_bl.png", "BR": "corner_br.png"}
    paths = {k: os.path.join(CORNER_DIR, v) for k, v in names.items()}

    # Auto-detect: kalau img_template.png lebih baru dari corner cache,
    # hapus cache lama supaya di-crop ulang dari template baru
    all_cached = all(os.path.exists(p) for p in paths.values())
    if all_cached:
        template_mtime = os.path.getmtime(TEMPLATE_IMG)
        oldest_corner  = min(os.path.getmtime(p) for p in paths.values())
        if template_mtime <= oldest_corner:
            return paths
        # Template lebih baru → hapus cache lama, crop ulang
        print("Template image berubah, regenerating corners...")
        for p in paths.values():
            os.remove(p)

    print("Cropping corners from template image...")
    img = Image.open(TEMPLATE_IMG).convert("RGBA")
    w, h = img.size

    # ── Auto-crop: bagi gambar jadi 4 kuadran, lalu deteksi bounding box
    #    dari pixel non-putih di tiap kuadran. Hasilnya = crop pas di gambar,
    #    bukan persentase dari tepi (jadi bebas ganti template tanpa adjust angka).
    THRESHOLD = 235   # pixel di atas nilai ini dianggap putih / background
    PADDING   = 5     # padding pixel tambahan di sekitar bounding box

    # Bagi jadi 4 kuadran (kiri/kanan x atas/bawah)
    hw, hh = w // 2, h // 2
    quadrants = {
        "TL": (0,  0,  hw, hh),
        "TR": (hw, 0,  w,  hh),
        "BL": (0,  hh, hw, h),
        "BR": (hw, hh, w,  h),
    }

    for corner, quad_box in quadrants.items():
        quadrant = img.crop(quad_box)
        pixels = quadrant.load()
        qw, qh = quadrant.size

        # Cari bounding box pixel non-putih di kuadran ini
        min_x, min_y = qw, qh
        max_x, max_y = 0, 0
        found = False

        for py in range(qh):
            for px in range(qw):
                r, g, b, a = pixels[px, py]
                if r < THRESHOLD or g < THRESHOLD or b < THRESHOLD:
                    if px < min_x: min_x = px
                    if px > max_x: max_x = px
                    if py < min_y: min_y = py
                    if py > max_y: max_y = py
                    found = True

        if not found:
            # Kuadran kosong, skip
            print(f"  {corner}: tidak ada gambar, skip")
            continue

        # Tambah padding, clamp ke batas kuadran
        min_x = max(0,      min_x - PADDING)
        min_y = max(0,      min_y - PADDING)
        max_x = min(qw - 1, max_x + PADDING)
        max_y = min(qh - 1, max_y + PADDING)

        # Crop ke bounding box ornamen
        cropped = quadrant.crop((min_x, min_y, max_x + 1, max_y + 1))

        # Buat pixel putih/near-white jadi transparan
        cpx = cropped.load()
        for py in range(cropped.height):
            for px in range(cropped.width):
                r, g, b, a = cpx[px, py]
                if r > THRESHOLD and g > THRESHOLD and b > THRESHOLD:
                    cpx[px, py] = (r, g, b, 0)

        cropped.save(paths[corner], "PNG")
        print(f"  {corner}: {cropped.width}x{cropped.height}px")

    print(f"Corners cached in {CORNER_DIR}/")
    return paths


# ── 4. Read Names ─────────────────────────────────────────────────────────────

def read_names(excel_path: str) -> list[tuple[str, str]]:
    """
    Read Excel, return list of (nama, alamat) tuples.
    Kalau kolom 'Alamat' tidak ada, alamat diisi string kosong.
    Header at row 2 (index 1).
    """
    df = pd.read_excel(excel_path, header=1, engine="openpyxl")
    col_nama = "Nama Lengkap"
    col_alamat = "Alamat"

    if col_nama not in df.columns:
        raise ValueError(
            f"Kolom '{col_nama}' tidak ditemukan. "
            f"Kolom yang ada: {list(df.columns)}"
        )

    # Ambil nama dan alamat, bersihkan whitespace
    results = []
    for _, row in df.iterrows():
        nama = str(row[col_nama]).strip() if pd.notna(row[col_nama]) else ""
        alamat = ""
        if col_alamat in df.columns and pd.notna(row.get(col_alamat)):
            alamat = str(row[col_alamat]).strip()
        if nama:
            results.append((nama, alamat))

    print(f"Loaded {len(results)} names from {excel_path}")
    return results


# ── 5. Draw Single Name Tag ──────────────────────────────────────────────────

def draw_nametag(c: pdf_canvas.Canvas, x: float, y: float,
                 width: float, height: float, name: str, alamat: str,
                 font_bold: str, font_reg: str,
                 corners: dict[str, str] | None):
    """Draw one complete name tag. (x,y) = bottom-left corner."""

    # Background
    c.setFillColor(COLOR_BG)
    c.rect(x, y, width, height, fill=1, stroke=0)

    # Outer border — opacity 0 (invisible, tapi struktur layout tetap konsisten)
    c.saveState()
    c.setStrokeAlpha(0)
    c.setLineWidth(0.8)
    c.rect(x, y, width, height, fill=0, stroke=1)
    c.restoreState()

    # Corner images — hanya Top-Left dan Bottom-Right
    cm = CORNER_MARGIN
    if corners:
        if "TL" in corners:
            c.drawImage(corners["TL"],
                        x + cm, y + height - CORNER_TOP_H - cm,
                        width=CORNER_TOP_W, height=CORNER_TOP_H, mask="auto")
        if "BR" in corners:
            c.drawImage(corners["BR"],
                        x + width - CORNER_BOT_W - cm, y + cm,
                        width=CORNER_BOT_W, height=CORNER_BOT_H, mask="auto")

    margin_text = CORNER_BOT_W * 0.5
    usable_w = width - 2 * margin_text

    # Auto-fit: mulai dari FONT_MAX, turunkan 0.5pt per iterasi
    # sampai teks nama muat dalam 1 baris, minimum FONT_MIN
    name_size = FONT_MAX
    while name_size >= FONT_MIN:
        text_w = c.stringWidth(name, font_bold, name_size)
        if text_w <= usable_w:
            break
        name_size -= 0.5

    sub_size = FONT_SUB

    # ── Jarak antar elemen (dalam mm, dikonversi ke canvas units) ─────
    gap_name_to_line = 2.5 * mm   # jarak dari baseline nama ke garis dekoratif
    gap_line_to_sub  = 2.0 * mm   # jarak dari garis dekoratif ke baris berikutnya
    gap_di_tempat    = 1.2 * mm   # jarak dari "di" ke "Tempat" (hanya kalau alamat kosong)

    # Titik tengah name tag
    center_x = x + width / 2
    center_y = y + height / 2

    # Hitung total tinggi blok teks supaya bisa di-center vertikal
    if alamat:
        # 2 baris: nama + garis + alamat
        total_h = (name_size
                   + gap_name_to_line + 0.3
                   + gap_line_to_sub + sub_size)
    else:
        # 3 baris: nama + garis + "di" + "Tempat"
        total_h = (name_size
                   + gap_name_to_line + 0.3
                   + gap_line_to_sub + sub_size
                   + gap_di_tempat + sub_size)

    # Vertical offset sama seperti versi original (kompensasi ornamen atas)
    vertical_offset = -1.5 * mm if alamat else -3 * mm
    top_of_block = center_y + total_h / 2 + vertical_offset

    # ── Baris 1: Nama tamu (font Bold) ──────────────────────────────
    name_y = top_of_block - name_size * 0.8
    c.setFillColor(COLOR_TEXT)
    c.setFont(font_bold, name_size)
    c.drawCentredString(center_x, name_y, name)

    # ── Garis dekoratif di bawah nama ─────────────────────────────────
    line_y = name_y - gap_name_to_line
    line_x1 = x + margin_text
    line_x2 = x + width - margin_text
    c.setStrokeColor(COLOR_LINE)
    c.setLineWidth(0.3)
    c.line(line_x1, line_y, line_x2, line_y)

    if alamat:
        # ── Baris 2: Alamat (font Regular) ───────────────────────────
        alamat_y = line_y - gap_line_to_sub - sub_size * 0.8
        c.setFillColor(COLOR_TEXT)
        c.setFont(font_reg, sub_size)
        c.drawCentredString(center_x, alamat_y, alamat)
    else:
        # ── Baris 2: "di" (font Regular) ─────────────────────────────
        di_y = line_y - gap_line_to_sub - sub_size * 0.8
        c.setFillColor(COLOR_TEXT)
        c.setFont(font_reg, sub_size)
        c.drawCentredString(center_x, di_y, "di")

        # ── Baris 3: "Tempat" (font Regular) ─────────────────────────
        tempat_y = di_y - gap_di_tempat - sub_size * 0.8
        c.drawCentredString(center_x, tempat_y, "Tempat")


# ── 6. Generate PDF ──────────────────────────────────────────────────────────

def generate_pdf(guests: list[tuple[str, str]], output_path: str,
                 font_bold: str, font_reg: str,
                 corners: dict[str, str] | None):
    """Create PDF with all name tags laid out in a 3x7 grid."""

    c = pdf_canvas.Canvas(output_path, pagesize=(PAGE_W, PAGE_H))
    c.setTitle("Wedding Name Tags")

    total = len(guests)
    total_pages = math.ceil(total / TAGS_PER) if total > 0 else 1

    for page_idx in range(total_pages):
        start = page_idx * TAGS_PER
        page_guests = guests[start: start + TAGS_PER]

        for slot_idx in range(TAGS_PER):
            col_idx = slot_idx % COLS
            row_idx = slot_idx // COLS

            # Posisi tag: margin halaman + (kolom/baris * (ukuran tag + gap))
            tag_x = PAGE_MARGIN_H + col_idx * (TAG_W + GAP_H)
            tag_y = PAGE_H - PAGE_MARGIN_V - (row_idx + 1) * TAG_H - row_idx * GAP_V
            # Baris teratas dikecilkan 2mm supaya muat di area cetak printer
            current_tag_h = (TAG_H - 2 * mm) if row_idx == 0 else TAG_H

            if slot_idx < len(page_guests):
                name, alamat = page_guests[slot_idx]
            else:
                name, alamat = "", ""

            if name:
                draw_nametag(c, tag_x, tag_y, TAG_W, current_tag_h, name, alamat,
                             font_bold, font_reg, corners)

        if page_idx < total_pages - 1:
            c.showPage()

    c.save()
    print(f"PDF saved: {output_path}  ({total_pages} page(s), {total} names)")


# ── 7. Main ──────────────────────────────────────────────────────────────────

def main():
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Step 1: Pilih file input
    excel_path  = pick_input_file()
    base_name   = os.path.splitext(os.path.basename(excel_path))[0]
    output_path = os.path.join(OUTPUT_DIR, f"{base_name}.pdf")

    # Step 2: Font (Bold untuk nama, Regular untuk "di" & "Tempat")
    fonts = prepare_fonts()
    font_bold = "Helvetica-Bold"   # fallback
    font_reg  = "Helvetica"        # fallback

    if fonts["bold"]:
        try:
            pdfmetrics.registerFont(TTFont(FONT_NAME_BOLD, fonts["bold"]))
            font_bold = FONT_NAME_BOLD
            print(f"Font Bold registered dari {fonts['bold']}")
        except Exception as e:
            print(f"Could not register Bold font: {e}. Using Helvetica-Bold.")

    if fonts["regular"]:
        try:
            pdfmetrics.registerFont(TTFont(FONT_NAME_REG, fonts["regular"]))
            font_reg = FONT_NAME_REG
            print(f"Font Regular registered dari {fonts['regular']}")
        except Exception as e:
            print(f"Could not register Regular font: {e}. Using Helvetica.")

    # Step 3: Corner images
    corners = prepare_corner_images()

    # Step 4: Read names + alamat
    guests = read_names(excel_path)
    if not guests:
        print("ERROR: Tidak ada nama yang terbaca dari Excel.")
        sys.exit(1)

    # Step 5: Generate PDF
    generate_pdf(guests, output_path, font_bold, font_reg, corners)
    print(f"Done! Buka {output_path} untuk preview.")


if __name__ == "__main__":
    main()
