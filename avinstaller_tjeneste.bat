@echo off
REM Avinstallerer tjenesten SupportsystemBackup. Kjor som administrator.
cd /d "%~dp0"
python backup_service.py stop
python backup_service.py remove
echo.
echo OK - Tjenesten er avinstallert.
pause
