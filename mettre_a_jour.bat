@echo off
REM Ouvre l interface de mise a jour du classeur (double-clic)
cd /d "%~dp0"
where pythonw >NUL 2>NUL
if %errorlevel%==0 (
  start "" pythonw interface.py
) else (
  python interface.py
  if errorlevel 1 pause
)
