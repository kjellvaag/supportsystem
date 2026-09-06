@echo off
REM Oppretter planlagt oppgave som kjorer backup.py i 'loop'-modus.
REM Leser intervall fra backup_config.json. Ingen administrator-rettigheter kreves.
cd /d "%~dp0"
for /f "delims=" %%a in ('python -c "import sys;print(sys.executable)"') do set PY=%%a
echo Oppretter oppgave "SupportsystemBackup" (loop-modus) ...
schtasks /create /tn "SupportsystemBackup" /tr "\"%PY%\" \"%~dp0backup.py\" loop" /sc onstart /ru "%USERNAME%" /f
echo.
echo OK - Oppgaven er opprettet. Test med: schtasks /run /tn "SupportsystemBackup"
pause
