"""Resim -> PDF dönüştürme. Kayıpsız gömme için img2pdf kullanılır; öncesinde
Pillow ile alfa düzleştirme / EXIF rotasyon / renk uzayı normalize edilir."""
from __future__ import annotations

import tempfile
from pathlib import Path

import img2pdf
from PIL import Image, ImageOps

from .. import fileops
from .types import ConversionOptions, ConversionResult


def _normalize_image(path: Path, tmp_dir: Path) -> Path:
    """img2pdf'e vermeden önce her zaman güvenli, kayıpsız bir PNG kopyası
    üretir: EXIF rotasyonu uygular, alfa kanalını beyaz zemine düzleştirir,
    CMYK/paletli gibi sıra dışı renk modlarını RGB'ye çevirir."""
    with Image.open(path) as img:
        img = ImageOps.exif_transpose(img)
        if img.mode in ("RGBA", "LA"):
            rgba = img.convert("RGBA")
            background = Image.new("RGB", rgba.size, (255, 255, 255))
            background.paste(rgba, mask=rgba.split()[-1])
            img = background
        elif img.mode == "P":
            rgba = img.convert("RGBA")
            background = Image.new("RGB", rgba.size, (255, 255, 255))
            background.paste(rgba, mask=rgba.split()[-1])
            img = background
        elif img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        normalized_path = tmp_dir / f"_norm_{fileops.sanitize_filename(path.stem)}.png"
        img.save(str(normalized_path), format="PNG")
    return normalized_path


def validate_image(path: Path) -> bool:
    try:
        with Image.open(path) as img:
            img.verify()
        return True
    except Exception:
        return False


def images_to_pdf(source: Path, dest: Path, options: ConversionOptions) -> ConversionResult:
    try:
        with tempfile.TemporaryDirectory() as tmp_dir_str:
            normalized = _normalize_image(source, Path(tmp_dir_str))
            pdf_bytes = img2pdf.convert(str(normalized))
        with fileops.atomic_write_target(dest) as tmp_path:
            tmp_path.write_bytes(pdf_bytes)
        return ConversionResult(success=True, output_paths=[dest], message=f"{dest.name} oluşturuldu.")
    except Exception as exc:  # noqa: BLE001
        return ConversionResult(success=False, message="Resim PDF'e dönüştürülemedi.", error_detail=repr(exc))
