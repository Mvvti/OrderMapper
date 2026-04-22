# ExcelMapper

ExcelMapper is a desktop Python + pywebview application for processing Excel-based orders and generating an output workbook.

## Requirements

- Windows
- Python 3.12
- Project dependencies installed in `.venv`

## Run Locally

```bat
.venv\Scripts\python.exe gui_app.py
```

## Build EXE (PyInstaller, onedir)

```bat
build.bat
```

After build, the application is available at:

```text
dist\ExcelMapper\ExcelMapper.exe
```

## Project Structure

- `gui_app.py` - backend and JS API bridge
- `frontend/` - UI files (HTML/CSS/JS)
- `src/` - data processing pipeline
- `app.spec` - PyInstaller configuration
- `hooks/hook_base_path.py` - runtime hook for base path handling
