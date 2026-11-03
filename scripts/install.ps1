# -*- coding: utf-8 -*-
# install.ps1 - jednorazová inštalácia aplikácie na novom počítači.
# Spustite kliknutím pravým tlačidlom na tento súbor -> "Spustiť s PowerShell",
# alebo v PowerSheli: powershell -ExecutionPolicy Bypass -File scripts\install.ps1

$ErrorActionPreference = "Stop"
# skript je v podpriečinku scripts\, projekt je o úroveň vyššie
$projekt = Split-Path $PSScriptRoot -Parent

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

# 3. PaddlePaddle - výpočtové jadro pre PaddleOCR (čítanie skenovaných faktúr).
#    S grafickou kartou NVIDIA verzia pre GPU (CUDA 12.9 - podporuje aj nové
#    karty RTX 50xx), inak verzia pre procesor. Modely OCR (~135 MB) sa
#    stiahnu automaticky pri prvej naskenovanej faktúre.
Write-Host ""
if (Get-Command nvidia-smi -ErrorAction SilentlyContinue) {
    Write-Host "Nájdená grafická karta NVIDIA - inštalujem PaddlePaddle pre GPU..."
    python -m pip install paddlepaddle-gpu==3.4.0 `
        -i https://www.paddlepaddle.org.cn/packages/stable/cu129/ `
        --extra-index-url https://pypi.org/simple
} else {
    Write-Host "Inštalujem PaddlePaddle pre procesor (CPU)..."
    python -m pip install "paddlepaddle>=3.2"
}

# 4. Konfigurácia - config.json nie je v gite; pri prvej inštalácii sa
#    vytvorí zo vzoru, existujúci (skopírovaný) súbor sa nikdy neprepíše.
$konfig = Join-Path $projekt "config.json"
if (-not (Test-Path $konfig)) {
    Copy-Item (Join-Path $projekt "config.example.json") $konfig
    Write-Host "Vytvorený config.json zo vzoru - doplňte IČO firmy a cesty." -ForegroundColor Yellow
}

# 5. Odkaz na pracovnej ploche
Write-Host ""
Write-Host "Vytváram odkaz na ploche..."
$desktop = [Environment]::GetFolderPath("Desktop")
$pythonDir = Split-Path $python.Source
$pythonw = Join-Path $pythonDir "pythonw.exe"
$shortcutPath = Join-Path $desktop "Sprava faktur TDS.lnk"
$ikona = Join-Path $projekt "assets\icon.ico"

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
Write-Host "Skontrolujte 'ico_firmy' a cesty k súborom v config.json."
Write-Host "Ak používate Google Kalendár, skopírujte aj credentials.json (token.json sa vytvorí"
Write-Host "znova pri prvom schválení faktúry - otvorí sa prehliadač na prihlásenie)."
