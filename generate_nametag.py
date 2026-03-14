# INSTALL: pip install pandas openpyxl reportlab requests fonttools otf2ttf Pillow
# RUN    : python generate_nametag.py
# OUTPUT : output/<nama_file>.pdf

import os
import sys
import math
import requests
import pandas as pd
from PIL import Image
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader

# ── Constants ──────────────────────────────────────────────────────────────────

INPUT_DIR     = "input"
OUTPUT_DIR    = "output"
CORNER_DIR    = "corners"
TEMPLATE_IMG  = "img_template.png"
FONT_TTF      = "CinzelDecorative-Regular.ttf"
FONT_OTF      = "CinzelDecorative-Regular.otf"

# Page: 210mm x 152mm landscape
PAGE_W = 210 * mm
PAGE_H = 152 * mm

COLS     = 3
ROWS     = 4
TAGS_PER = COLS * ROWS

MARGIN = 2 * mm
GAP_H  = 2 * mm
GAP_V  = 2 * mm

TAG_W = 67 * mm
TAG_H = 36 * mm

# Corner ornament size on the name tag
CORNER_W = 16 * mm
CORNER_H = 14 * mm

# Colors — monochrome black on white
COLOR_BG           = HexColor("#FFFFFF")
COLOR_BORDER_OUTER = HexColor("#000000")
COLOR_TEXT         = HexColor("#000000")
COLOR_LINE         = HexColor("#000000")

# Font sizes
FONT_MAX  = 11
FONT_MIN  = 7
FONT_NAME = "CinzelDecorative"


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


def prepare_font() -> str | None:
    import re

    if os.path.exists(FONT_TTF) and _is_valid_ttf(FONT_TTF):
        return FONT_TTF

    if os.path.exists(FONT_TTF) and not _is_valid_ttf(FONT_TTF):
        print(f"Removing invalid {FONT_TTF}...")
        os.remove(FONT_TTF)

    if os.path.exists(FONT_OTF):
        if _convert_otf_to_ttf(FONT_OTF, FONT_TTF):
            return FONT_TTF

    try:
        print("Downloading font from Google Fonts...")
        css_url = "https://fonts.googleapis.com/css?family=Cinzel+Decorative:400"
        headers = {"User-Agent": "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1)"}
        r = requests.get(css_url, headers=headers, timeout=15)
        if r.status_code == 200:
            urls = re.findall(r'url\((https://fonts\.gstatic\.com/[^)]+)\)', r.text)
            if urls:
                rf = requests.get(urls[0], timeout=30)
                if rf.status_code == 200 and len(rf.content) > 10_000:
                    with open(FONT_TTF, "wb") as f:
                        f.write(rf.content)
                    if _is_valid_ttf(FONT_TTF):
                        return FONT_TTF
                    if _convert_otf_to_ttf(FONT_TTF, FONT_TTF):
                        return FONT_TTF
    except Exception as e:
        print(f"  Google Fonts failed: {e}")

    try:
        print("Trying GitHub mirror...")
        github_url = (
            "https://github.com/google/fonts/raw/main"
            "/ofl/cinzeldecorative/CinzelDecorative-Regular.ttf"
        )
        r = requests.get(github_url, timeout=30, allow_redirects=True)
        if r.status_code == 200 and len(r.content) > 10_000:
            with open(FONT_TTF, "wb") as f:
                f.write(r.content)
            if _is_valid_ttf(FONT_TTF):
                return FONT_TTF
    except Exception as e:
        print(f"  GitHub mirror failed: {e}")

    print("WARNING: Could not obtain font. Using Helvetica as fallback.")
    return None


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

    if all(os.path.exists(p) for p in paths.values()):
        return paths

    print("Cropping corners from template image...")
    img = Image.open(TEMPLATE_IMG).convert("RGBA")
    w, h = img.size

    # Crop fractions — how much of each corner to grab
    fx = 0.36   # 36% width from edge
    fy = 0.46   # 46% height from edge
    cw = int(w * fx)
    ch = int(h * fy)

    crops = {
        "TL": (0,      0,      cw, ch),
        "TR": (w - cw, 0,      w,  ch),
        "BL": (0,      h - ch, cw, h),
        "BR": (w - cw, h - ch, w,  h),
    }

    for corner, box in crops.items():
        cropped = img.crop(box)

        # Make white/near-white pixels transparent
        pixels = cropped.load()
        for py in range(cropped.height):
            for px in range(cropped.width):
                r, g, b, a = pixels[px, py]
                if r > 235 and g > 235 and b > 235:
                    pixels[px, py] = (r, g, b, 0)

        cropped.save(paths[corner], "PNG")

    print(f"Corners cached in {CORNER_DIR}/")
    return paths


# ── 4. Read Names ─────────────────────────────────────────────────────────────

def read_names(excel_path: str) -> list[str]:
    """Read Excel, return list of names. Header at row 2 (index 1)."""
    df = pd.read_excel(excel_path, header=1, engine="openpyxl")
    col = "Nama Lengkap"
    if col not in df.columns:
        raise ValueError(
            f"Kolom '{col}' tidak ditemukan. "
            f"Kolom yang ada: {list(df.columns)}"
        )
    names = df[col].dropna().astype(str).tolist()
    names = [n.strip() for n in names if n.strip()]
    print(f"Loaded {len(names)} names from {excel_path}")
    return names


# ── 5. Draw Single Name Tag ──────────────────────────────────────────────────

def draw_nametag(c: pdf_canvas.Canvas, x: float, y: float,
                 width: float, height: float, name: str,
                 font_registered: bool, corners: dict[str, str] | None):
    """Draw one complete name tag. (x,y) = bottom-left corner."""

    # Background
    c.setFillColor(COLOR_BG)
    c.rect(x, y, width, height, fill=1, stroke=0)

    # Outer border
    c.setStrokeColor(COLOR_BORDER_OUTER)
    c.setLineWidth(0.8)
    c.rect(x, y, width, height, fill=0, stroke=1)

    # Corner images
    if corners:
        # Top-left
        c.drawImage(corners["TL"],
                    x, y + height - CORNER_H,
                    width=CORNER_W, height=CORNER_H, mask="auto")
        # Top-right
        c.drawImage(corners["TR"],
                    x + width - CORNER_W, y + height - CORNER_H,
                    width=CORNER_W, height=CORNER_H, mask="auto")
        # Bottom-left
        c.drawImage(corners["BL"],
                    x, y,
                    width=CORNER_W, height=CORNER_H, mask="auto")
        # Bottom-right
        c.drawImage(corners["BR"],
                    x + width - CORNER_W, y,
                    width=CORNER_W, height=CORNER_H, mask="auto")

    # ══════════════════════════════════════════════════════════════════════
    # TEXT LAYOUT — 3 baris teks, centered horizontal & vertical
    #
    #   ┌─────────────────────────────────┐
    #   │  [corner]            [corner]   │
    #   │                                 │
    #   │         Agus & Partner          │  ← baris 1: nama (auto-fit)
    #   │        ─────────────────        │  ← garis dekoratif
    #   │               di                │  ← baris 2: "di"
    #   │             Tempat              │  ← baris 3: "Tempat"
    #   │                                 │
    #   │  [corner]            [corner]   │
    #   └─────────────────────────────────┘
    #
    # Offset vertikal digeser 2mm ke bawah supaya teks tidak terlalu
    # mepet ke ornamen bidadari di sudut atas.
    # ══════════════════════════════════════════════════════════════════════

    # Pilih font: CinzelDecorative kalau tersedia, fallback Helvetica
    font_name = FONT_NAME if font_registered else "Helvetica"

    # Lebar area teks = lebar name tag dikurangi margin kiri-kanan
    # (setengah lebar corner supaya teks tidak tertimpa gambar sudut)
    margin_text = CORNER_W * 0.5
    usable_w = width - 2 * margin_text

    # Auto-fit: mulai dari FONT_MAX (11pt), turunkan 0.5pt per iterasi
    # sampai teks nama muat dalam 1 baris, minimum FONT_MIN (7pt)
    name_size = FONT_MAX
    while name_size >= FONT_MIN:
        text_w = c.stringWidth(name, font_name, name_size)
        if text_w <= usable_w:
            break
        name_size -= 0.5

    sub_size = 7  # ukuran font tetap untuk "di" dan "Tempat"

    # ── Jarak antar elemen (dalam mm, dikonversi ke canvas units) ─────
    gap_name_to_line = 2.5 * mm   # jarak dari baseline nama ke garis dekoratif
    gap_line_to_di   = 2.0 * mm   # jarak dari garis dekoratif ke "di"
    gap_di_tempat    = 1.2 * mm   # jarak dari "di" ke "Tempat"

    # Titik tengah name tag
    center_x = x + width / 2
    center_y = y + height / 2

    # Hitung total tinggi blok teks supaya bisa di-center vertikal:
    #   nama + gap + garis(0.3) + gap + "di" + gap + "Tempat"
    total_h = (name_size
               + gap_name_to_line + 0.3
               + gap_line_to_di + sub_size
               + gap_di_tempat + sub_size)

    # Geser blok 2mm ke bawah agar tidak terlalu dekat ornamen atas
    vertical_offset = -2 * mm
    top_of_block = center_y + total_h / 2 + vertical_offset

    # ── Baris 1: Nama tamu ────────────────────────────────────────────
    # Baseline = puncak blok dikurangi 80% tinggi font (cap height approx)
    name_y = top_of_block - name_size * 0.8
    c.setFillColor(COLOR_TEXT)
    c.setFont(font_name, name_size)
    c.drawCentredString(center_x, name_y, name)

    # ── Garis dekoratif di bawah nama ─────────────────────────────────
    line_y = name_y - gap_name_to_line
    line_x1 = x + margin_text          # mulai dari margin kiri
    line_x2 = x + width - margin_text  # sampai margin kanan
    c.setStrokeColor(COLOR_LINE)
    c.setLineWidth(0.3)
    c.line(line_x1, line_y, line_x2, line_y)

    # ── Baris 2: "di" ────────────────────────────────────────────────
    di_y = line_y - gap_line_to_di - sub_size * 0.8
    c.setFillColor(COLOR_TEXT)
    c.setFont(font_name, sub_size)
    c.drawCentredString(center_x, di_y, "di")

    # ── Baris 3: "Tempat" ────────────────────────────────────────────
    tempat_y = di_y - gap_di_tempat - sub_size * 0.8
    c.drawCentredString(center_x, tempat_y, "Tempat")


# ── 6. Generate PDF ──────────────────────────────────────────────────────────

def generate_pdf(names: list[str], output_path: str,
                 font_registered: bool, corners: dict[str, str] | None):
    """Create PDF with all name tags laid out in a 3x4 grid."""

    c = pdf_canvas.Canvas(output_path, pagesize=(PAGE_W, PAGE_H))
    c.setTitle("Wedding Name Tags")

    total = len(names)
    total_pages = math.ceil(total / TAGS_PER) if total > 0 else 1

    for page_idx in range(total_pages):
        start = page_idx * TAGS_PER
        page_names = names[start: start + TAGS_PER]

        for slot_idx in range(TAGS_PER):
            col_idx = slot_idx % COLS
            row_idx = slot_idx // COLS

            tag_x = MARGIN + col_idx * (TAG_W + GAP_H)
            tag_y = PAGE_H - MARGIN - (row_idx + 1) * TAG_H - row_idx * GAP_V

            if slot_idx < len(page_names):
                name = page_names[slot_idx]
            else:
                name = ""

            if name:
                draw_nametag(c, tag_x, tag_y, TAG_W, TAG_H, name,
                             font_registered, corners)
            else:
                c.setFillColor(COLOR_BG)
                c.setStrokeColor(COLOR_BORDER_OUTER)
                c.setLineWidth(0.3)
                c.rect(tag_x, tag_y, TAG_W, TAG_H, fill=1, stroke=1)

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

    # Step 2: Font
    font_registered = False
    font_file = prepare_font()
    if font_file:
        try:
            pdfmetrics.registerFont(TTFont(FONT_NAME, font_file))
            font_registered = True
            print(f"Font {FONT_NAME} registered from {font_file}.")
        except Exception as e:
            print(f"Could not register font: {e}. Using Helvetica.")

    # Step 3: Corner images
    corners = prepare_corner_images()

    # Step 4: Read names
    names = read_names(excel_path)
    if not names:
        print("ERROR: Tidak ada nama yang terbaca dari Excel.")
        sys.exit(1)

    # Step 5: Generate PDF
    generate_pdf(names, output_path, font_registered, corners)
    print(f"Done! Buka {output_path} untuk preview.")


if __name__ == "__main__":
    main()
