@echo off
chcp 65001 >nul
echo ========================================
echo  Close_ExWin - Build Script
echo ========================================
echo.

cd /d "%~dp0"

set PYTHON=D:\Anaconda3\envs\WinTool\python.exe
set PIP=D:\Anaconda3\envs\WinTool\Scripts\pip.exe
set PYINSTALLER=D:\Anaconda3\envs\WinTool\Scripts\pyinstaller.exe

if not exist "%PYTHON%" (
    echo [ERROR] WinTool env not found: %PYTHON%
    pause & exit /b 1
)

echo [1/3] Installing packages...
"%PIP%" install pyinstaller pywin32 pystray pillow -q
if errorlevel 1 (
    echo [ERROR] pip install failed
    pause & exit /b 1
)

echo [2/3] Building EXE...
"%PYINSTALLER%" --onefile --windowed --clean --name "Close_ExWin" Close_ExWIN.py
if errorlevel 1 (
    echo [ERROR] PyInstaller failed
    pause & exit /b 1
)

echo [3/3] Done!
copy /Y "dist\Close_ExWin.exe" "Close_ExWin.exe" >nul
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
echo.
echo  Close_ExWin.exe is ready.
echo.
pause
