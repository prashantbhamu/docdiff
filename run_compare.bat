@echo off
setlocal

set "ROOT=%~dp0"
title DocDiff - Document Comparator

if exist "%ROOT%dist\DocDiff.exe" (
    "%ROOT%dist\DocDiff.exe"
    exit /b %errorlevel%
)

if not exist "%ROOT%src\docdiff_app.py" (
    echo [ERROR] DocDiff source files are missing.
    pause
    exit /b 1
)

if not exist "%ROOT%.venv\Scripts\python.exe" (
    echo DocDiff development environment is not installed.
    echo Run setup_dev.bat first, or build the portable EXE with build_exe.bat.
    pause
    exit /b 1
)

"%ROOT%.venv\Scripts\python.exe" "%ROOT%src\docdiff_app.py"
exit /b %errorlevel%
