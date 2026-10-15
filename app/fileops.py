"""Çıktı dosyası güvenliği: dosya adı temizleme, path traversal koruması,
çakışma önleme, atomik yazma. Çıktı dosya adları YALNIZCA kullanıcının seçtiği
kaynak dosya adından türetilir — hiçbir zaman belge içeriğinden/metadata'dan
(örn. PDF /Title alanı) türetilmez.
"""
from __future__ import annotations

import contextlib
import os
import re
import tempfile
from pathlib import Path

_INVALID_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sanitize_filename(name: str) -> str:
    """Windows'ta geçersiz karakterleri temizler, path traversal bileşenlerini
    (.., kök ayraçları) kaldırır. Sadece taban dosya adı için kullanılır."""
    name = Path(name).name  # herhangi bir dizin bileşenini at
    name = _INVALID_CHARS.sub("_", name).strip(" .")
    if name in ("", ".", ".."):
        name = "belge"
    return name[:200]


def ensure_within(path: Path, expected_dir: Path) -> Path:
    """path'in çözümlenmiş halinin expected_dir içinde kaldığını doğrular.
    Değilse ValueError fırlatır (path traversal koruması)."""
    resolved = path.resolve()
    expected = expected_dir.resolve()
    if expected not in resolved.parents and resolved != expected:
        raise ValueError(f"Geçersiz çıktı yolu: {resolved} klasör dışında ({expected}).")
    return resolved


def unique_destination(path: Path) -> Path:
    """Dosya zaten varsa, Windows Gezgini tarzı 'ad (1).ext' ile çakışmayan
    bir yol üretir; var olan dosyanın üzerine sessizce yazmaz."""
    if not path.exists():
        return path
    stem, suffix, parent = path.stem, path.suffix, path.parent
    counter = 1
    while True:
        candidate = parent / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


@contextlib.contextmanager
def atomic_write_target(dest_path: Path):
    """Hedefle AYNI klasörde geçici bir dosya yolu verir (os.replace'in atomik
    olması için aynı diskte olması şart). Bloğun içinde o geçici yola yazılır;
    blok hatasız biterse hedefe atomik taşınır, hata olursa geçici dosya silinir.
    """
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{dest_path.stem}_", suffix=dest_path.suffix, dir=dest_path.parent
    )
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        yield tmp_path
        os.replace(tmp_path, dest_path)
    except BaseException:
        tmp_path.unlink(missing_ok=True)
        raise


def build_output_path(source: Path, output_dir: Path, target_ext: str, suffix: str = "") -> Path:
    """Kaynak dosya adından, beklenen klasör içinde, çakışmayan ve sanitize
    edilmiş bir çıktı yolu üretir."""
    base_name = sanitize_filename(source.stem) + suffix
    candidate = output_dir / f"{base_name}.{target_ext}"
    candidate = ensure_within(candidate, output_dir)
    return unique_destination(candidate)
