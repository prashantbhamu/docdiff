@echo off
setlocal

set "ROOT=%~dp0"

if not exist "%ROOT%.venv\Scripts\python.exe" call "%ROOT%setup_dev.bat"
if errorlevel 1 exit /b 1

"%ROOT%.venv\Scripts\python.exe" -c "import PyInstaller" >nul 2>&1
if errorlevel 1 call "%ROOT%setup_dev.bat"
if errorlevel 1 exit /b 1

pushd "%ROOT%"
echo Building portable DocDiff.exe...
"%ROOT%.venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean DocDiff.spec
set "RESULT=%errorlevel%"
popd

if not "%RESULT%"=="0" exit /b %RESULT%
echo.
echo Portable executable created at dist\DocDiff.exe
exit /b 0
