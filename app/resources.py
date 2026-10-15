"""Paketlenmiş kaynakların (ör. ikon) hem geliştirme modunda (`python main.py`)
hem de PyInstaller ile donmuş .exe modunda doğru yolunu çözer. Donmuş modda
PyInstaller bootloader'ı `sys._MEIPASS`'ı paketlenmiş veri kökü olarak ayarlar."""
from __future__ import annotations

import sys
from pathlib import Path


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / relative
