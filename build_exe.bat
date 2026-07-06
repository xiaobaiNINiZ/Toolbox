@echo off
chcp 65001 >nul
echo ========================================
echo   Video Rename Tool - Build EXE
echo ========================================
echo.

REM Check dependencies
pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    pip install -r requirements.txt
    echo.
)

echo Building...
echo.

REM Build EXE (use English name to avoid encoding issues)
pyinstaller --onefile --windowed --name "VideoRenameTool" --clean video_rename_tool.py

echo.
echo ========================================
if exist "dist\VideoRenameTool.exe" (
    echo Build SUCCESS!
    echo EXE location: dist\VideoRenameTool.exe
    echo.
    echo You can rename it to: 视频批量重命名工具.exe
) else (
    echo Build FAILED! Check error messages.
)
echo ========================================
echo.
pause