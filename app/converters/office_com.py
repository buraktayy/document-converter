"""Word/Excel/PowerPoint -> PDF COM otomasyonu ile yüksek sadakatli dönüştürme.

Not: PDF -> Word dönüşümü burada DEĞİL, pdf_tools.py'de pdf2docx ile yapılır.
Word'ün yerleşik "PDF aç ve dönüştür" özelliği test edilirken, otomasyonla
bastırılamayan bir onay diyaloğu ("Word, PDF'inizi düzenlenebilir bir Word
belgesine dönüştürecek...") nedeniyle COM üzerinden süresiz donduğu tespit
edildi (Visible=True/False, ConfirmConversions=False, DisplayAlerts=False
hiçbiri bu diyaloğu engellemiyor) — bu yüzden o yöndeki dönüşüm için
pdf2docx kütüphanesine geçildi.

Güvenlik kuralları:
- gencache / makepy KULLANILMAZ (bilinen kararsızlık kaynağı) — sabit tamsayı
  sabitleri kullanılır.
- Her Application nesnesi DispatchEx ile İZOLE, görünmez, uyarısız ve
  makroları zorla devre dışı bırakarak (AutomationSecurity) oluşturulur.
- Belgeler her zaman salt-okunur açılır; bu fonksiyonlar kaynak dosyayı asla
  değiştirmez, sadece okuyup başka bir yola aktarır (export/save-as).
- Donma riskine karşı: her office_app() çağrısı bir izleme (watchdog) zamanlayıcısı
  kurar; süre aşılırsa ilgili Office process'i PID üzerinden zorla sonlandırılır.
"""
from __future__ import annotations

import contextlib
import threading
from pathlib import Path

import pythoncom
import win32api
import win32com.client
import win32con
import win32process
from pywintypes import com_error

from .. import fileops
from .types import ConversionOptions, ConversionResult

# --- Sabitler (gencache KULLANILMADAN, version-independent integer sabitler) ---
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WD_REVISIONS_VIEW_FINAL = 0
XL_TYPE_PDF = 0
PP_SAVE_AS_PDF = 32


class ConversionError(Exception):
    """Türkçe, kullanıcıya gösterilebilir mesaj taşıyan dönüştürme hatası."""

    def __init__(self, message: str, detail: str | None = None):
        super().__init__(message)
        self.message = message
        self.detail = detail


class _ProcessWatchdog:
    """Bir Office process'inin belirli bir süre içinde tamamlanmasını bekler;
    aşılırsa process'i PID üzerinden zorla sonlandırır (donmuş COM çağrılarını
    kurtarmak için — COM nesneleri thread-güvenli olmadığından doğrudan iptal
    edilemez, ama işletim sistemi seviyesinde process sonlandırılabilir)."""

    def __init__(self, timeout_seconds: float):
        self._timeout = timeout_seconds
        self._timer: threading.Timer | None = None
        self._pid: int | None = None
        self.timed_out = False

    def arm(self, pid: int) -> None:
        self._pid = pid
        self._timer = threading.Timer(self._timeout, self._kill)
        self._timer.daemon = True
        self._timer.start()

    def _kill(self) -> None:
        self.timed_out = True
        if self._pid is None:
            return
        try:
            handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, self._pid)
            try:
                win32api.TerminateProcess(handle, -1)
            finally:
                win32api.CloseHandle(handle)
        except Exception:
            pass

    def disarm(self) -> None:
        if self._timer is not None:
            self._timer.cancel()


@contextlib.contextmanager
def office_app(prog_id: str, timeout_seconds: float = 240):
    """Worker thread üzerinden çağrılmalıdır. CoInitialize/CoUninitialize'ı
    kendi thread'i için yönetir; izole, görünmez, makro-devre-dışı bir Office
    Application örneği oluşturur ve donma korumalı şekilde verir."""
    pythoncom.CoInitialize()
    app = None
    watchdog = _ProcessWatchdog(timeout_seconds)
    try:
        app = win32com.client.DispatchEx(prog_id)
        try:
            app.Visible = False
        except Exception:
            pass  # PowerPoint bazı sürümlerde Visible=False'a izin vermez
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        if hasattr(app, "AutomationSecurity"):
            try:
                app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:
                pass
        if prog_id == "Excel.Application":
            try:
                app.AskToUpdateLinks = False
                app.EnableEvents = False
            except Exception:
                pass

        try:
            _, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
            watchdog.arm(pid)
        except Exception:
            pass

        yield app

        if watchdog.timed_out:
            raise ConversionError("İşlem çok uzun sürdü ve durduruldu.")
    finally:
        watchdog.disarm()
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
            del app
        pythoncom.CoUninitialize()


def _translate_error(exc: Exception) -> str:
    if isinstance(exc, ConversionError):
        return exc.message
    if isinstance(exc, com_error):
        return "Office uygulaması dönüştürmeyi tamamlayamadı (dosya bozuk veya desteklenmeyen bir özellik içeriyor olabilir)."
    if isinstance(exc, PermissionError):
        return "Hedef dosya başka bir programda açık olabilir, lütfen kapatıp tekrar deneyin."
    return "Beklenmeyen bir hata oluştu."


def word_to_pdf(source: Path, dest: Path, options: ConversionOptions) -> ConversionResult:
    try:
        with fileops.atomic_write_target(dest) as tmp_path:
            with office_app("Word.Application", options.timeout_seconds) as app:
                doc = app.Documents.Open(
                    str(source), ReadOnly=True, ConfirmConversions=False, AddToRecentFiles=False
                )
                try:
                    try:
                        doc.ActiveWindow.View.RevisionsView = WD_REVISIONS_VIEW_FINAL
                        doc.ActiveWindow.View.ShowRevisionsAndComments = False
                    except Exception:
                        pass  # görünüm ayarlanamasa da export devam eder
                    doc.ExportAsFixedFormat(
                        OutputFileName=str(tmp_path),
                        ExportFormat=WD_EXPORT_FORMAT_PDF,
                        OptimizeFor=WD_EXPORT_OPTIMIZE_FOR_PRINT,
                    )
                finally:
                    doc.Close(SaveChanges=False)
        return ConversionResult(success=True, output_paths=[dest], message=f"{dest.name} oluşturuldu.")
    except Exception as exc:  # noqa: BLE001 - kullanıcıya çevrilmiş mesaj döndürülüyor
        return ConversionResult(success=False, message=_translate_error(exc), error_detail=repr(exc))


def excel_to_pdf(source: Path, dest: Path, options: ConversionOptions) -> ConversionResult:
    try:
        with fileops.atomic_write_target(dest) as tmp_path:
            with office_app("Excel.Application", options.timeout_seconds) as app:
                wb = app.Workbooks.Open(str(source), ReadOnly=True, UpdateLinks=0, AddToMru=False)
                try:
                    wb.ExportAsFixedFormat(Type=XL_TYPE_PDF, Filename=str(tmp_path))
                finally:
                    wb.Close(SaveChanges=False)
        return ConversionResult(success=True, output_paths=[dest], message=f"{dest.name} oluşturuldu.")
    except Exception as exc:  # noqa: BLE001
        return ConversionResult(success=False, message=_translate_error(exc), error_detail=repr(exc))


def powerpoint_to_pdf(source: Path, dest: Path, options: ConversionOptions) -> ConversionResult:
    try:
        with fileops.atomic_write_target(dest) as tmp_path:
            with office_app("PowerPoint.Application", options.timeout_seconds) as app:
                pres = app.Presentations.Open(
                    str(source), ReadOnly=True, Untitled=False, WithWindow=False
                )
                try:
                    pres.SaveAs(str(tmp_path), FileFormat=PP_SAVE_AS_PDF)
                finally:
                    pres.Close()
        return ConversionResult(success=True, output_paths=[dest], message=f"{dest.name} oluşturuldu.")
    except Exception as exc:  # noqa: BLE001
        return ConversionResult(success=False, message=_translate_error(exc), error_detail=repr(exc))
