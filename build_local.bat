@echo off
REM Local Windows build script
REM Run this from the repo root after activating your Python environment

echo Installing / updating dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Building ABR Analysis Tool...
pyinstaller ABR_Analysis_Tool.spec --clean

echo.
if exist "dist\ABR Analysis Tool\ABR Analysis Tool.exe" (
    echo BUILD SUCCEEDED
    echo Executable: dist\ABR Analysis Tool\ABR Analysis Tool.exe
) else (
    echo BUILD FAILED - check output above
)
pause
