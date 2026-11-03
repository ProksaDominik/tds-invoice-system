@echo off
setlocal

echo ============================================
echo  Dochadzka TDS - jednorazove nastavenie PC
echo ============================================
echo.

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python sa nenasiel - instalujem cez winget...
    winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
    echo.
    echo Python bol prave nainstalovany. Windows potrebuje restartovat
    echo tento skript, aby ho nasiel - spustite prosim setup.bat ESTE RAZ.
    pause
    exit /b 0
) else (
    echo Python: OK
)

if not exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    echo Tesseract OCR sa nenasiel - instalujem cez winget...
    winget install -e --id UB-Mannheim.TesseractOCR --accept-source-agreements --accept-package-agreements
) else (
    echo Tesseract OCR: OK
)

echo.
echo Instalujem Python kniznice...
python -m pip install --upgrade pip
python -m pip install -r "%~dp0requirements.txt"

echo.
echo Vytvaram odkaz na plochu...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$desktop = [Environment]::GetFolderPath('Desktop'); $ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut(\"$desktop\Dochadzka TDS.lnk\"); $sc.TargetPath = '%~dp0start.bat'; $sc.WorkingDirectory = '%~dp0'; $sc.IconLocation = '%~dp0icon.ico'; $sc.Save()"

echo.
echo ============================================
echo  Hotovo! Na ploche je teraz odkaz "Dochadzka TDS".
echo ============================================
pause
