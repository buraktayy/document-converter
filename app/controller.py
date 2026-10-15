"""UI ile dönüştürücüler arasındaki orkestrasyon katmanı: format tespiti,
doğrulama, dönüştürücü seçimi ve arka plan thread'inde çalıştırma.

UI thread'i asla COM/PDF işlemleriyle bloklanmaz: her iş ayrı bir daemon
thread'de çalışır, sonuç bir queue.Queue üzerinden UI'nin root.after() ile
periyodik olarak okuduğu mesajlarla iletilir.
"""
from __future__ import annotations

import logging
import queue
import threading
from dataclasses import dataclass, field
from pathlib import Path

from . import converters, fileops, security
from .converters.types import ConversionOptions, ConversionResult

logger = logging.getLogger(__name__)


@dataclass
class ConversionJob:
    source: Path
    target_format: str
    output_dir: Path
    options: ConversionOptions = field(default_factory=ConversionOptions)


def detect_source_format(path: Path) -> str | None:
    """Seçilen dosyanın GERÇEK formatını (uzantıya değil, dosya içeriğine göre)
    döndürür; tanınmazsa None. Bu sayede sahte uzantılı dosyalar UI'da da
    doğru/şüpheli işaretlenebilir."""
    return security.detect_format(path)


def available_targets(src_ext: str) -> list[str]:
    return converters.available_targets(src_ext)


def run_conversion_async(job: ConversionJob, result_queue: "queue.Queue[tuple]") -> threading.Thread:
    """Dönüştürmeyi arka plan thread'inde başlatır. result_queue'ya şu olaylar
    yazılır: ("started", None), ("done", ConversionResult), ("error", str)."""
    thread = threading.Thread(target=_run_job, args=(job, result_queue), daemon=True)
    thread.start()
    return thread


def _run_job(job: ConversionJob, result_queue: "queue.Queue[tuple]") -> None:
    result_queue.put(("started", None))
    try:
        actual_src_ext = security.validate_file(
            job.source, job.source.suffix, max_size_mb=job.options.max_file_size_mb
        )
        target_ext = security.normalize_ext(job.target_format)

        converter_fn = converters.get_converter(actual_src_ext, target_ext)
        if converter_fn is None:
            result_queue.put(
                ("error", f".{actual_src_ext} -> .{target_ext} dönüşümü desteklenmiyor.")
            )
            return

        dest = fileops.build_output_path(job.source, job.output_dir, target_ext)
        result: ConversionResult = converter_fn(job.source, dest, job.options)

        if result.success:
            result_queue.put(("done", result))
        else:
            result_queue.put(("error", result.message))
            if result.error_detail:
                logger.error("Dönüştürme hatası (%s): %s", job.source, result.error_detail)
    except security.SecurityValidationError as exc:
        result_queue.put(("error", exc.message))
    except Exception:  # noqa: BLE001 - beklenmeyen her hata kullanıcıya genel mesajla döner
        logger.exception("Beklenmeyen dönüştürme hatası: %s", job.source)
        result_queue.put(("error", "Beklenmeyen bir hata oluştu."))
