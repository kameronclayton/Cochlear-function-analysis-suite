@echo off
REM Local Windows build script for CoFAST
REM Run this from the repo root after activating your Python environment

echo Installing / updating dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Building CoFAST...
pyinstaller CoFAST.spec --clean

echo.
if exist "dist\CoFAST\CoFAST.exe" (
    echo BUILD SUCCEEDED
    echo Executable: dist\CoFAST\CoFAST.exe
) else (
    echo BUILD FAILED - check output above
)
pause
