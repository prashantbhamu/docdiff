# DocDiff

DocDiff is a local Windows application for comparing PDF, Word, and text documents. It reconstructs document blocks, aligns related content, and produces a side-by-side HTML report with word-level additions and deletions.

## Use the portable application

Build the executable with `build_exe.bat`, then run `dist\DocDiff.exe`. The browser opens automatically at `http://127.0.0.1:8765`.

The executable is self-contained. A target Windows PC does not need Python or the project dependencies installed. Generated reports are saved in a `reports` folder beside the executable.

## Run from source

1. Install Python 3.11 or newer.
2. Create a virtual environment and install `requirements.txt`.
3. Run `python src\docdiff_app.py`.

The application opens in the browser at `http://127.0.0.1:8765`.

## Command-line comparison

```powershell
python src\compare_docs.py original.pdf revised.pdf --output report.html
```

## Build

```powershell
build_exe.bat
```

The build script creates its own virtual environment, installs the required packages, and runs PyInstaller. It creates the untracked binary at `dist\DocDiff.exe`. Executables, uploaded documents, generated reports, and build directories are intentionally excluded from Git.

## Privacy

DocDiff binds only to `127.0.0.1`. Uploaded files are processed locally and temporary copies are deleted after comparison. The application does not require external web fonts or cloud services.
