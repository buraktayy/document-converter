"""Belge Dönüştürücü için marka ikonunu üretir: krem/mercan paletinden bir
"belge + dönüşüm oku" işareti çizer ve çoklu çözünürlüklü .ico olarak
assets/icon.ico'ya yazar. assets/icon.png olarak yüksek çözünürlüklü kaynak
da bırakılır (gelecekte yeniden düzenlemek için).

Çalıştırma: python scripts/generate_icon.py
"""
from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageDraw

CREAM = (246, 243, 236, 255)
ACCENT = (232, 84, 44, 255)
ACCENT_DARK = (200, 67, 31, 255)
INK = (26, 24, 20, 255)

SIZE = 1024
ROOT = Path(__file__).resolve().parent.parent
ASSETS_DIR = ROOT / "assets"


def _arrow_points(cx: float, cy: float, radius: float, start_deg: float, end_deg: float, width: float):
    """start_deg'den end_deg'e bir yay boyunca kalın bir şerit (arrow gövdesi)
    için nokta listesi üretir."""
    steps = 24
    outer = []
    inner = []
    for i in range(steps + 1):
        t = start_deg + (end_deg - start_deg) * i / steps
        rad = math.radians(t)
        outer.append((cx + (radius + width / 2) * math.cos(rad), cy + (radius + width / 2) * math.sin(rad)))
        inner.append((cx + (radius - width / 2) * math.cos(rad), cy + (radius - width / 2) * math.sin(rad)))
    return outer + inner[::-1]


def build_icon() -> Image.Image:
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    corner = int(SIZE * 0.24)
    draw.rounded_rectangle([0, 0, SIZE, SIZE], radius=corner, fill=ACCENT)

    # --- belge (sayfa) ---
    page_w, page_h = SIZE * 0.40, SIZE * 0.52
    page_x = SIZE * 0.27
    page_y = (SIZE - page_h) / 2 - SIZE * 0.03
    fold = page_w * 0.32

    page_rect = [page_x, page_y, page_x + page_w, page_y + page_h]
    draw.rounded_rectangle(page_rect, radius=SIZE * 0.025, fill=CREAM)

    # kıvrılmış köşe (üst sağ)
    fold_pts = [
        (page_x + page_w - fold, page_y),
        (page_x + page_w, page_y + fold),
        (page_x + page_w - fold, page_y + fold),
    ]
    draw.polygon(fold_pts, fill=ACCENT_DARK)

    # belge içindeki satır çizgileri
    line_x0 = page_x + page_w * 0.16
    line_x1 = page_x + page_w * 0.84
    for i, frac in enumerate((0.42, 0.55, 0.68)):
        y = page_y + page_h * frac
        x1 = line_x1 if i != 2 else page_x + page_w * 0.62
        draw.line([(line_x0, y), (x1, y)], fill=ACCENT, width=int(SIZE * 0.018))

    # --- dönüşüm oku (sağ alt, daire üzerinde) ---
    badge_r = SIZE * 0.20
    badge_cx = SIZE * 0.715
    badge_cy = SIZE * 0.715
    draw.ellipse(
        [badge_cx - badge_r, badge_cy - badge_r, badge_cx + badge_r, badge_cy + badge_r],
        fill=CREAM,
    )

    arrow_r = badge_r * 0.55
    arrow_w = badge_r * 0.26
    arc1 = _arrow_points(badge_cx, badge_cy, arrow_r, -160, 80, arrow_w)
    draw.polygon(arc1, fill=INK)
    arc2 = _arrow_points(badge_cx, badge_cy, arrow_r, 20, 260, arrow_w)
    draw.polygon(arc2, fill=INK)

    # ok başları
    head = arrow_w * 1.1
    tip1 = (badge_cx + arrow_r * math.cos(math.radians(80)), badge_cy + arrow_r * math.sin(math.radians(80)))
    draw.polygon(
        [
            (tip1[0] - head, tip1[1] - head * 0.4),
            (tip1[0] + head * 0.3, tip1[1] + head),
            (tip1[0] + head, tip1[1] - head * 0.5),
        ],
        fill=INK,
    )
    tip2 = (badge_cx + arrow_r * math.cos(math.radians(260)), badge_cy + arrow_r * math.sin(math.radians(260)))
    draw.polygon(
        [
            (tip2[0] + head, tip2[1] + head * 0.4),
            (tip2[0] - head * 0.3, tip2[1] - head),
            (tip2[0] - head, tip2[1] + head * 0.5),
        ],
        fill=INK,
    )

    return img


def main() -> None:
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    icon = build_icon()
    icon.save(ASSETS_DIR / "icon.png")
    icon.save(ASSETS_DIR / "icon.ico", sizes=[(16, 16), (32, 32), (48, 48), (256, 256)])
    print(f"yazıldı: {ASSETS_DIR / 'icon.png'}")
    print(f"yazıldı: {ASSETS_DIR / 'icon.ico'}")


if __name__ == "__main__":
    main()
