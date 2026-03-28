@echo off
setlocal

cd /d "%~dp0"

if exist ".venv\Scripts\pythonw.exe" start "" ".venv\Scripts\pythonw.exe" "app.py" & goto :eof
if exist ".venv\Scripts\python.exe" start "" ".venv\Scripts\python.exe" "app.py" & goto :eof

where pythonw >nul 2>nul
if "%errorlevel%"=="0" start "" pythonw "app.py" & goto :eof

where python >nul 2>nul
if "%errorlevel%"=="0" start "" python "app.py" & goto :eof

echo Python was not found.
echo Please install Python or recreate the project's .venv.
pause
