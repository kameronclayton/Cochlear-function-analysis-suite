@echo off
REM ── ABR / DPOAE Analysis Tool launcher ──────────────────────────────────────
REM Double-click this file to launch the analysis tool.
REM Searches common Python locations (Anaconda, standard installer, PATH).

setlocal

REM ── Try to find a working python executable ─────────────────────────────────
set PYTHON=

REM 1. Python already on PATH?
where python >nul 2>&1 && set PYTHON=python

REM 2. Anaconda in ProgramData (system-wide install)?
if not defined PYTHON (
    if exist "C:\ProgramData\Anaconda3\python.exe" (
        set PYTHON=C:\ProgramData\Anaconda3\python.exe
    )
)

REM 3. Anaconda in user folder?
if not defined PYTHON (
    if exist "%USERPROFILE%\Anaconda3\python.exe" (
        set PYTHON=%USERPROFILE%\Anaconda3\python.exe
    )
)

REM 4. Miniconda in ProgramData?
if not defined PYTHON (
    if exist "C:\ProgramData\miniconda3\python.exe" (
        set PYTHON=C:\ProgramData\miniconda3\python.exe
    )
)

REM 5. Miniconda in user folder?
if not defined PYTHON (
    if exist "%USERPROFILE%\miniconda3\python.exe" (
        set PYTHON=%USERPROFILE%\miniconda3\python.exe
    )
)

if not defined PYTHON (
    echo.
    echo ERROR: Python could not be found.
    echo Please install Python 3.9+ from https://www.python.org or Anaconda from
    echo https://www.anaconda.com, then re-run this file.
    echo.
    pause
    exit /b 1
)

echo Using Python: %PYTHON%
echo.

REM ── Check / install required packages ───────────────────────────────────────
%PYTHON% -c "import numpy, pandas, openpyxl, matplotlib" >nul 2>&1
if errorlevel 1 (
    echo Required packages not found. Installing now -- this may take a minute...
    echo.
    %PYTHON% -m pip install numpy pandas openpyxl matplotlib
    echo.
)

REM ── Launch the tool ──────────────────────────────────────────────────────────
%PYTHON% "%~dp0abr_analysis_tool.py"

if errorlevel 1 (
    echo.
    echo The tool exited with an error. Press any key to close.
    pause
)

endlocal
