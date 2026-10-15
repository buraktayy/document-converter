"""Girdi dosyası güvenlik doğrulamaları: format/uzantı eşleşmesi, şifreleme,
boyut sınırı. Bir dosya herhangi bir dönüştürücüye verilmeden önce buradan geçer.
"""
from __future__ import annotations

import zipfile
from pathlib import Path

import fitz  # PyMuPDF

MAX_FILE_SIZE_MB_DEFAULT = 200

# Office Open XML (docx/xlsx/pptx) içerik tipi -> normalize uzantı eşlemesi.
# Üçü de aynı PK\x03\x04 zip imzasını paylaştığı için içerik tipinden ayrıştırılır.
_OOXML_CONTENT_TYPE_EXT = {
    "word/document.xml": "docx",
    "xl/workbook.xml": "xlsx",
    "ppt/presentation.xml": "pptx",
}

_MAGIC_SIGNATURES: dict[str, tuple[bytes, ...]] = {
    "pdf": (b"%PDF-",),
    "jpg": (b"\xff\xd8\xff",),
    "png": (b"\x89PNG\r\n\x1a\n",),
}

SUPPORTED_EXTENSIONS = {"pdf", "docx", "xlsx", "pptx", "jpg", "png"}


class SecurityValidationError(Exception):
    """Bir dosya doğrulamadan geçemediğinde fırlatılır. message Türkçe ve
    kullanıcıya doğrudan gösterilebilir."""

    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


def normalize_ext(ext: str) -> str:
    ext = ext.lower().lstrip(".")
    return "jpg" if ext == "jpeg" else ext


def _sniff_ooxml(path: Path) -> str | None:
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            for marker, ext in _OOXML_CONTENT_TYPE_EXT.items():
                if marker in names:
                    return ext
    except (zipfile.BadZipFile, OSError):
        return None
    return None


def detect_format(path: Path) -> str | None:
    """Dosyanın gerçek içeriğine bakarak normalize uzantı döndürür (magic-byte
    + OOXML içerik tipi sniffing). Tanınmazsa None döner."""
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except OSError:
        return None

    for ext, signatures in _MAGIC_SIGNATURES.items():
        if any(head.startswith(sig) for sig in signatures):
            return ext

    if head.startswith(b"PK\x03\x04"):
        return _sniff_ooxml(path)

    return None


def validate_file(path: Path, claimed_ext: str, max_size_mb: int = MAX_FILE_SIZE_MB_DEFAULT) -> str:
    """Tam doğrulama: var olma, boyut, uzantı/içerik eşleşmesi, şifreleme.
    Başarılıysa normalize edilmiş gerçek formatı döndürür; aksi halde
    SecurityValidationError fırlatır (Türkçe mesajla)."""
    if not path.is_file():
        raise SecurityValidationError("Dosya bulunamadı.")

    size = path.stat().st_size
    if size == 0:
        raise SecurityValidationError("Dosya boş (0 byte).")
    if size > max_size_mb * 1024 * 1024:
        raise SecurityValidationError(
            f"Dosya çok büyük (maks. {max_size_mb} MB destekleniyor)."
        )

    claimed = normalize_ext(claimed_ext)
    if claimed not in SUPPORTED_EXTENSIONS:
        raise SecurityValidationError(f"Desteklenmeyen dosya türü: .{claimed}")

    actual = detect_format(path)
    if actual is None:
        raise SecurityValidationError(
            "Dosya içeriği tanınamadı veya bozuk olabilir."
        )
    if actual != claimed:
        raise SecurityValidationError(
            f"Dosya uzantısı (.{claimed}) gerçek içerikle (.{actual}) eşleşmiyor."
        )

    if actual == "pdf" and is_encrypted_pdf(path):
        raise SecurityValidationError("Şifreli/korumalı PDF dosyaları desteklenmiyor.")

    return actual


def is_encrypted_pdf(path: Path) -> bool:
    try:
        with fitz.open(path) as doc:
            return bool(doc.needs_pass)
    except Exception:
        return False


def pdf_has_text_layer(path: Path, sample_pages: int = 5) -> bool:
    """Word'e gönderilmeden önce taranmış (metin katmanı olmayan) PDF tespiti için."""
    try:
        with fitz.open(path) as doc:
            pages_to_check = list(doc)[:sample_pages]
            for page in pages_to_check:
                if page.get_text().strip():
                    return True
    except Exception:
        return True  # belirsizse engelleme, dönüştürücü kendi hatasını versin
    return False
