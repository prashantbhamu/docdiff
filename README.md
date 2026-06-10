# DocDiff

DocDiff is a local Windows application for comparing PDF, Word, and text documents. It reconstructs document blocks, aligns related content, and produces a side-by-side HTML report with word-level additions and deletions.

## Use the portable application

Build the executable with `build_exe.bat`, then run `dist\DocDiff.exe`. The browser opens automatically at `http://127.0.0.1:8765`.

The executable is self-contained. A target Windows PC does not need Python or the project dependencies installed. Generated reports are saved in a `reports` folder beside the executable.

## Run from source

1. Install Python 3.11 or newer.
2. Run `setup_dev.bat`.
3. Run `run_compare.bat`.

The launcher uses `dist\DocDiff.exe` when present. Otherwise, it runs the application from `.venv`.

## Command-line comparison

```powershell
.venv\Scripts\python.exe src\compare_docs.py original.pdf revised.pdf --output report.html
```

## Build

```powershell
build_exe.bat
```

PyInstaller creates the untracked binary at `dist\DocDiff.exe`. Executables, uploaded documents, generated reports, and build directories are intentionally excluded from Git.

## Privacy

DocDiff binds only to `127.0.0.1`. Uploaded files are processed locally and temporary copies are deleted after comparison. The application does not require external web fonts or cloud services.
