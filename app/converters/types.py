"""Converter modülleri arasında paylaşılan veri tipleri. Döngüsel import'u
önlemek için office_com.py / pdf_tools.py / image_tools.py hepsi bunu import
eder, __init__.py de bunu re-export eder."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class ConversionOptions:
    dpi: int = 200
    timeout_seconds: int = 240
    max_file_size_mb: int = 200


@dataclass(frozen=True)
class ConversionResult:
    success: bool
    output_paths: list[Path] = field(default_factory=list)
    message: str = ""
    error_detail: str | None = None
