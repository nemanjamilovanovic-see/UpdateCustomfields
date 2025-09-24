# UpdateCustomfields

Desktop tool to batch-update LiveSD topics' `customfields.YearInMaintenance` from an Excel file with columns `tp_ID` and `Godina`.

## Run (source)
- Install Python 3.10+ and dependencies: `pip install -r requirements.txt`
- Configure server/timeouts in `updatecustomfields.yaml`.
- Start the GUI: `python UpdateCustomfields.py`

## Excel format
- Required columns: `tp_ID` (topic id), `Godina` (year)
- Rows with missing values are skipped.

## Configuration (updatecustomfields.yaml)
- `live_url`: Base API, e.g. `https://livesd.company.com/live20Rest/`
- `concurrency`: Worker threads, e.g. 10
- `connect_timeout`: Seconds for TCP connect, e.g. 10
- `search_read_timeout`: Seconds for search read, e.g. 30
- `save_read_timeout`: Seconds for save read, e.g. 120
- `retry_backoff_s`: Sleep before retry after timeout, e.g. 0.3

## Build EXE (PyInstaller)
- Using VS Code task: `Terminal -> Run Task -> Build EXE with PyInstaller`
- Or from terminal (with .spec):
	- `pyinstaller --clean --noconfirm --distpath dist --workpath build UpdateCustomfields.spec`
- Result: `dist/UpdateCustomfields/UpdateCustomfields/UpdateCustomfields.exe`

### Run built app (Windows)
```cmd
cd "c:\\Users\\nemanja.milovanovic\\Desktop\\New folder (10)\\ZaMilicu\\UpdateCustomfields"
"dist\\UpdateCustomfields\\UpdateCustomfields.exe"
```

### One-file build (optional)
If you prefer a single EXE, update the spec or build with `--onefile` (may increase start time). When using `.spec`, set `EXE(..., name='UpdateCustomfields', console=False)` and run `pyinstaller --onefile UpdateCustomfields.spec`.

## Notes
- The app reuses connections per thread and verifies saves after timeouts.
- If a record already has the same year, it is skipped to speed up processing.

