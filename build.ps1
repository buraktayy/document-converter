# build.ps1 - Belge Dönüştürücü'nün bağımsız .exe'sini üretir.
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$pythonCommand = Get-Command python -ErrorAction SilentlyContinue
if ($pythonCommand) {
    $python = $pythonCommand.Source
} else {
    $python = Join-Path $env:LOCALAPPDATA "Programs\Python\Python311\python.exe"
}
if (-not (Test-Path -LiteralPath $python)) {
    throw "Python bulunamadı. Python 3.11 kurulumunu kontrol edin."
}

Write-Host "PyInstaller ve hooks-contrib kuruluyor/güncelleniyor..." -ForegroundColor Cyan
& $python -m pip install --upgrade pyinstaller pyinstaller-hooks-contrib
if ($LASTEXITCODE -ne 0) { throw "pip install pyinstaller başarısız oldu" }

Write-Host "İkon üretiliyor..." -ForegroundColor Cyan
& $python scripts/generate_icon.py
if ($LASTEXITCODE -ne 0) { throw "İkon üretimi başarısız oldu" }

Write-Host "Önceki build kalıntıları temizleniyor..." -ForegroundColor Cyan
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue ".\build", ".\dist"

Write-Host "PyInstaller çalıştırılıyor..." -ForegroundColor Cyan
& $python -m PyInstaller --noconfirm --clean DocumentConverter.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller build başarısız oldu" }

$exePath = Join-Path $PSScriptRoot "dist\DocumentConverter\DocumentConverter.exe"
if (-not (Test-Path $exePath)) { throw "Build tamamlandı ama exe bulunamadı: $exePath" }

$pymupdfDir = Join-Path $PSScriptRoot "dist\DocumentConverter\_internal\pymupdf"
$requiredNativeFiles = @("_extra.pyd", "_mupdf.pyd", "mupdfcpp64.dll")
$missingNativeFiles = @(
    $requiredNativeFiles | Where-Object {
        -not (Test-Path -LiteralPath (Join-Path $pymupdfDir $_))
    }
)
if ($missingNativeFiles.Count -gt 0) {
    throw "Build eksik PyMuPDF native dosyaları içeriyor: $($missingNativeFiles -join ', ')"
}

Write-Host "Build tamamlandı: $exePath" -ForegroundColor Green
