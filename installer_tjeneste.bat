@echo off
REM Installerer SupportsystemBackup som Windows-tjeneste. Kjor som administrator.
cd /d "%~dp0"
echo Sjekker pywin32...
python -c "import win32serviceutil" 2>nul
if errorlevel 1 (
    echo Installerer pywin32...
    python -m pip install pywin32
)
python backup_service.py install
python backup_service.py start
echo.
echo OK - Tjenesten "SupportsystemBackup" er installert og startet.
pause
