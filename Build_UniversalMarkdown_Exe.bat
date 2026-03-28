@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Project virtual environment was not found.
    echo Please recreate .venv before building the exe.
    pause
    exit /b 1
)

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onedir --windowed --name UniversalMarkdown_Portable --collect-submodules markitdown --collect-submodules markitdown_ocr --collect-data magika --copy-metadata markitdown --copy-metadata markitdown-ocr --copy-metadata openai --copy-metadata magika app.py
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

if exist "dist\UniversalMarkdown_Portable\_UserSettings" rmdir /s /q "dist\UniversalMarkdown_Portable\_UserSettings"

echo Build completed: %cd%\dist\UniversalMarkdown_Portable\UniversalMarkdown_Portable.exe
echo Keep the whole UniversalMarkdown_Portable folder together when sharing it.
pause
