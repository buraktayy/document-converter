# DocumentConverter.spec
# Build with: pyinstaller --noconfirm --clean DocumentConverter.spec
#
# customtkinter, tkinterdnd2 ve python-pptx runtime'da PyInstaller'ın statik
# analiziyle görünmeyen veri/native dosyalar yüklüyor (tema JSON'ları, tkdnd
# native dll/tcl, pptx şablonu). Her biri için collect_all() ile savunma
# amaçlı toplama yapılıyor; var olan bir contrib hook'u çiğnemez (idempotent).

from pathlib import Path

from PyInstaller.utils.hooks import collect_all

ROOT = Path(SPECPATH)

datas = [(str(ROOT / "assets" / "icon.ico"), "assets")]
binaries = []
hiddenimports = [
    # pywin32 / COM otomasyonu (office_com.py) — win32timezone sessizce
    # eksik kalan klasik bir hidden import.
    "win32timezone",
    "win32com",
    "win32com.client",
    "win32com.server",
    "pythoncom",
    "pywintypes",
    "win32api",
    "win32con",
    "win32process",
]

for pkg in (
    "customtkinter",
    "tkinterdnd2",
    "pymupdf",
    "pdfplumber",
    "pdf2docx",
    "pptx",     # python-pptx'in import adı "pptx"
):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

# PyMuPDF 1.28+ iki CPython extension modülü (_extra, _mupdf) ve bunların
# kullandığı mupdfcpp DLL'i ile çalışır. PyInstaller normalde hidden import
# analiziyle bunları bulur; ancak bazı PyMuPDF / hooks-contrib kombinasyonları
# extension'ları COLLECT aşamasında atlayabiliyor. Native dosyaları doğrudan
# pymupdf paket dizinine ekleyerek "cannot import name '_extra'" başlangıç
# hatasını engelle. `fitz` ayrı collect_all gerektirmez: PyMuPDF'in uyumluluk
# paketi normal import analizi tarafından zaten toplanır.
import pymupdf

pymupdf_dir = Path(pymupdf.__file__).resolve().parent
pymupdf_native_files = []
for pattern in ("*.pyd", "*.dll"):
    pymupdf_native_files.extend(pymupdf_dir.glob(pattern))

required_pymupdf_native = {"_extra.pyd", "_mupdf.pyd", "mupdfcpp64.dll"}
found_pymupdf_native = {path.name for path in pymupdf_native_files}
missing_pymupdf_native = required_pymupdf_native - found_pymupdf_native
if missing_pymupdf_native:
    raise RuntimeError(
        "PyMuPDF native dosyaları bulunamadı: "
        + ", ".join(sorted(missing_pymupdf_native))
    )

binaries += [(str(path), "pymupdf") for path in pymupdf_native_files]

a = Analysis(
    ["main.py"],
    pathex=[str(ROOT)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DocumentConverter",
    debug=False,
    strip=False,
    upx=False,
    console=False,
    icon=str(ROOT / "assets" / "icon.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name="DocumentConverter",
)
