# New-DesktopShortcut.ps1 - Masaüstüne "Belge Dönüştürücü" kısayolu oluşturur/günceller.
# build.ps1'den ayrı tutulur ki her yeniden build, masaüstünü sessizce tekrar dokunmasın.
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$exePath = Join-Path $PSScriptRoot "dist\DocumentConverter\DocumentConverter.exe"
if (-not (Test-Path $exePath)) { throw "Exe bulunamadı: $exePath - önce build.ps1 çalıştırın." }
$exePath = (Resolve-Path $exePath).Path
$workDir = Split-Path $exePath -Parent

$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "Belge Dönüştürücü.lnk"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $exePath
$shortcut.WorkingDirectory = $workDir
$shortcut.IconLocation = "$exePath,0"
$shortcut.Description = "Belge Dönüştürücü"
$shortcut.Save()

Write-Host "Masaüstü kısayolu oluşturuldu: $shortcutPath" -ForegroundColor Green
