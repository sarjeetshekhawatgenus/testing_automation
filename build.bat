@echo off
setlocal enabledelayedexpansion

echo ================================================================
echo   Building MasterAutomation Windows Application (.exe)
echo ================================================================

cd /d "%~dp0"

echo [1/3] Checking environment...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    pause
    exit /b 1
)

echo [2/3] Checking dependencies...
python -c "import PyInstaller, streamlit, pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing application and build dependencies...
    python -m pip install -r requirements.txt pyinstaller==6.22.3
    if errorlevel 1 (
        echo [ERROR] Dependency installation failed.
        pause
        exit /b 1
    )
)

echo [3/3] Compiling Executable with PyInstaller...
python -m PyInstaller --noconfirm --clean --distpath dist-uan-suffix MasterAutomation.spec

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. Please check the logs above.
    pause
    exit /b 1
)

echo.
echo ================================================================
echo   Build Successful!
echo   Executable: dist-uan-suffix\MasterAutomation\MasterAutomation.exe
echo   IMPORTANT: Distribute the entire MasterAutomation folder, not just the exe
echo ================================================================
echo.
pause
