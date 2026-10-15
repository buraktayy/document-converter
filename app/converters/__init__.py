"""Dönüştürücü kayıt defteri (registry): (kaynak uzantı, hedef uzantı) ->
dönüştürücü fonksiyon. controller.py burayı kullanarak doğru fonksiyonu bulur.
"""
from __future__ import annotations

from pathlib import Path
from typing import Callable

from ..security import normalize_ext
from . import image_tools, office_com, pdf_tools
from .types import ConversionOptions, ConversionResult

ConverterFn = Callable[[Path, Path, ConversionOptions], ConversionResult]

REGISTRY: dict[tuple[str, str], ConverterFn] = {
    ("docx", "pdf"): office_com.word_to_pdf,
    ("pdf", "docx"): pdf_tools.pdf_to_word,
    ("xlsx", "pdf"): office_com.excel_to_pdf,
    ("pdf", "xlsx"): pdf_tools.pdf_to_excel,
    ("pptx", "pdf"): office_com.powerpoint_to_pdf,
    ("pdf", "pptx"): pdf_tools.pdf_to_powerpoint,
    ("jpg", "pdf"): image_tools.images_to_pdf,
    ("png", "pdf"): image_tools.images_to_pdf,
    ("pdf", "jpg"): pdf_tools.pdf_to_images,
    ("pdf", "png"): pdf_tools.pdf_to_images,
}

# Çoklu çıktı üreten (tek dosya yerine klasör) dönüştürmeler — controller bu
# bilgiyi kullanarak çıktı yolunu (dosya mı, klasör mü) doğru hazırlar.
PRODUCES_MULTIPLE_OUTPUTS = {("pdf", "jpg"), ("pdf", "png")}


def get_converter(src_ext: str, dst_ext: str) -> ConverterFn | None:
    return REGISTRY.get((normalize_ext(src_ext), normalize_ext(dst_ext)))


def available_targets(src_ext: str) -> list[str]:
    src_ext = normalize_ext(src_ext)
    return sorted({dst for (src, dst) in REGISTRY if src == src_ext})


def produces_multiple_outputs(src_ext: str, dst_ext: str) -> bool:
    return (normalize_ext(src_ext), normalize_ext(dst_ext)) in PRODUCES_MULTIPLE_OUTPUTS
