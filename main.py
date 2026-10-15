"""Belge Dönüştürücü uygulamasının giriş noktası: `python main.py` ile veya
paketlenmiş .exe olarak çalıştırılır.

Pencereli (konsolsuz) build'de yakalanmamış bir istisna hiçbir iz bırakmadan
kaybolur; bu yüzden kalıcı bir dosya logger'ı kuruluyor ki ileride "neden
açılmıyor" sorunları yeniden build almadan teşhis edilebilsin.
"""
from __future__ import annotations

import logging
import os
import sys
import traceback
from pathlib import Path

_LOG_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "DocumentConverter"
_LOG_FILE = _LOG_DIR / "crash.log"


def _install_excepthook() -> None:
    _LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=_LOG_FILE,
        level=logging.ERROR,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    def _handle(exc_type, exc_value, exc_tb):
        logging.error("Yakalanmamış istisna:\n%s", "".join(traceback.format_exception(exc_type, exc_value, exc_tb)))

    sys.excepthook = _handle


_install_excepthook()

from app.ui import main  # noqa: E402 - excepthook kurulduktan sonra import edilmeli

if __name__ == "__main__":
    main()
