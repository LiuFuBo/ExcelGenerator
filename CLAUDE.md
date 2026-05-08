# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Run Commands

- **Dev run:** `python3.11 app.py`
- **Mac .dmg build:** `bash build_mac.sh` — produces `dist/ExcelGenerator.app` and `ExcelGenerator.dmg`
- **Windows .exe build:** run `build_windows.bat` on a Windows machine, or rely on GitHub Actions (`push to main` triggers `.github/workflows/build_exe.yml`)
- **Install deps:** `pip install PySide6 openpyxl pypinyin PyInstaller`

## Architecture

Single-file PySide6/Qt app. `ExcelGeneratorApp` (QMainWindow) in `app.py` handles all UI and data logic.

**Core data flow (generation):**
1. Generate random sales dates (month 1st → today, 9:00–21:00)
2. Sort date + order columns only (other columns untouched)
3. Regenerate order numbers from sorted dates (YYMMDD from same-row date)

**File modes:**
- **New file:** user types Excel name → creates `.xlsx` in app's directory
- **Append file:** user picks existing `.xlsx` via QFileDialog → finds or creates columns, appends data, sorts date+order columns

**Key methods:**
- `_build_ui` — constructs all widgets
- `start_generation` — validates inputs, computes file path, delegates to `_generate_new_file` or `_generate_append_file`
- `_sort_date_and_order` — sorts only date_col and order_col values, preserving other columns
- `generate_order_number_from_date` — extracts YYMMDD from date string for S+YYMMDD+14-digit format

## Cross-Platform Notes

- Mac build uses `--windowed` (directory .app bundle); Windows GitHub Actions build uses `--onefile --windowed` (single .exe)
- Both include `--add-data="jj_header_rounded.png"` for the app icon resource
- **Do NOT display the cover image inside the app UI** — `jj_header_rounded.png` is only used as the app/dock icon, not as an in-app image
- In .app bundles, `sys.executable` points to `Contents/MacOS/`, so base_dir calculation uses `os.path.dirname(sys.executable)` + navigation to `Contents/Resources/` when needed

## Known Pitfalls

- **StyleProxy:** openpyxl `cell.alignment` returns an unhashable StyleProxy. Always construct real `Alignment(horizontal=..., vertical=..., wrap_text=...)` objects instead of using `cell.alignment` directly in comparisons or assignments that trigger openpyxl's internal indexing.
- **traceback import:** Any `traceback.format_exc()` usage requires `import traceback` — missing it causes the error handler itself to fail and mask real errors.
- **Shell syntax:** GitHub Actions workflow must use `shell: cmd` for PyInstaller commands — `^` line continuation is cmd.exe syntax, not PowerShell.
- **Sorting scope:** Only sort date and order columns. Never move other column data.

## Push Policy

Only push to GitHub when the user explicitly requests it. Do not auto-push after changes.