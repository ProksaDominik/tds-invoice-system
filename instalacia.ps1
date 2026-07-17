# -*- coding: utf-8 -*-
# instalacia.ps1  jednorazová inštalácia aplikácie na novom počítači.
# Spustite kliknutím pravým tlačidlom na tento súbor -> "Spustiť s PowerShell",
# alebo v PowerSheli: powershell -ExecutionPolicy Bypass -File instalacia.ps1

$ErrorActionPreference = "Stop"
$projekt = $PSScriptRoot

Write-Host "== Inštalácia: Správa faktúr TDS s.r.o. ==" -ForegroundColor Cyan
Write-Host "Priečinok projektu: $projekt"
Write-Host ""

# 1. Python
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Host "Python nie je nainštalovaný alebo nie je v PATH." -ForegroundColor Yellow
    Write-Host "Stiahnite Python 3.10+ z https://www.python.org/downloads/"
    Write-Host "Pri inštalácii zaškrtnite 'Add Python to PATH', potom spustite tento skript znova."
    exit 1
}
Write-Host "Python nájdený: $($python.Source)" -ForegroundColor Green

# 2. Python knižnice
Write-Host ""
Write-Host "Inštalujem Python knižnice z requirements.txt..."
python -m pip install -r (Join-Path $projekt "requirements.txt")

# 3. Tesseract OCR (potrebné pre čítanie skenovaných/vytlačených faktúr)
Write-Host ""
$tesseractExe = @(
    "C:\Program Files\Tesseract-OCR\tesseract.exe",
    "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $tesseractExe -and (Get-Command tesseract -ErrorAction SilentlyContinue)) {
    $tesseractExe = (Get-Command tesseract).Source
}

if (-not $tesseractExe) {
    Write-Host "Tesseract OCR nie je nainštalovaný, inštalujem cez winget..."
    winget install --id UB-Mannheim.TesseractOCR --silent --accept-package-agreements --accept-source-agreements
} else {
    Write-Host "Tesseract OCR už je nainštalovaný: $tesseractExe" -ForegroundColor Green
}

# jazykové dáta (slk.traineddata, eng.traineddata) sú súčasťou priečinka
# tessdata/ v tomto projekte, netreba ich sťahovať znova.

# 4. Odkaz na pracovnej ploche
Write-Host ""
Write-Host "Vytváram odkaz na ploche..."
$desktop = [Environment]::GetFolderPath("Desktop")
$pythonDir = Split-Path $python.Source
$pythonw = Join-Path $pythonDir "pythonw.exe"
$shortcutPath = Join-Path $desktop "Sprava faktur TDS.lnk"
$ikona = Join-Path $projekt "icon.ico"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $pythonw
$shortcut.Arguments = '"app.py"'
$shortcut.WorkingDirectory = $projekt
if (Test-Path $ikona) {
    $shortcut.IconLocation = "$ikona,0"
}
$shortcut.Description = "Správa faktúr - TDS s.r.o."
$shortcut.Save()

Write-Host "Odkaz vytvorený: $shortcutPath" -ForegroundColor Green

Write-Host ""
Write-Host "== Hotovo ==" -ForegroundColor Cyan
Write-Host "Skontrolujte 'ico_firmy' v config.json, ak ste ho ešte neskopírovali s projektom."
Write-Host "Ak používate Google Kalendár, skopírujte aj credentials.json (token.json sa vytvorí"
Write-Host "znova pri prvom schválení faktúry - otvorí sa prehliadač na prihlásenie)."
