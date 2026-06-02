@echo off
chcp 65001 >nul
echo ========================================
echo  Close_ExWin - EXE 打包腳本
echo ========================================
echo.

cd /d "%~dp0"

set PYTHON=D:\Anaconda3\envs\WinTool\python.exe
set PIP=D:\Anaconda3\envs\WinTool\Scripts\pip.exe
set PYINSTALLER=D:\Anaconda3\envs\WinTool\Scripts\pyinstaller.exe

if not exist "%PYTHON%" (
    echo [錯誤] 找不到 WinTool 環境：%PYTHON%
    pause & exit /b 1
)

echo [1/3] 安裝必要套件...
"%PIP%" install pyinstaller pywin32 pystray pillow -q
if errorlevel 1 (
    echo [錯誤] 套件安裝失敗
    pause & exit /b 1
)

echo [2/3] 打包 EXE...
"%PYINSTALLER%" --onefile --windowed --clean --name "Close_ExWin" Close_ExWIN.py
if errorlevel 1 (
    echo [錯誤] 打包失敗
    pause & exit /b 1
)

echo [3/3] 完成！
copy /Y "dist\Close_ExWin.exe" "Close_ExWin.exe" >nul
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
echo.
echo  Close_ExWin.exe 已產生，直接執行即可
echo.
pause
