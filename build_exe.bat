@echo off
setlocal

set "ROOT=%~dp0"
set "VENV=%ROOT%.venv"

if not exist "%VENV%\Scripts\python.exe" (
    echo Creating Python environment...
    py -3.11 -m venv "%VENV%" >nul 2>&1
    if errorlevel 1 python -m venv "%VENV%"
    if errorlevel 1 (
        echo [ERROR] Python 3.11 or newer is required.
        echo Install Python from https://www.python.org/downloads/windows/
        pause
        exit /b 1
    )
)

echo Installing DocDiff dependencies...
"%VENV%\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 exit /b 1
"%VENV%\Scripts\python.exe" -m pip install -r "%ROOT%requirements.txt"
if errorlevel 1 exit /b 1

pushd "%ROOT%"
echo Building portable DocDiff.exe...
"%VENV%\Scripts\python.exe" -m PyInstaller --noconfirm --clean DocDiff.spec
set "RESULT=%errorlevel%"
popd

if not "%RESULT%"=="0" exit /b %RESULT%
echo.
echo Portable executable created at dist\DocDiff.exe
exit /b 0
