# Fuel, Mandatory Insurance, and Indices Update

A Tkinter desktop app that scrapes prices and writes parameter files for KolNatun.

Three tabs:

- **Fuel prices** — Paz and delekulator
- **Mandatory car insurance** — Capital Market Authority (CMA)
- **Indices** — CBS API and US BLS CPI

## Requirements

- Python 3.7 or later
- Windows (Access files are created via COM)
- Internet connection
- Chrome (for Selenium-based scrapes)

## Install and run

```bash
pip install -r requirements.txt
python main_app.py
```

Or run from the project folder:

- `הפעל_תוכנה.bat` — uses the local venv
- `run_main.bat` — uses Python on PATH

Build an EXE with `07092026.spec` (PyInstaller).

## Output

Files are written under `C:\Users\shir.feldman\Desktop\parametrsUpdate` (set in `config.py`):

| Area | Folder | Files |
|------|--------|--------|
| Fuel | `DELEK` | text, MDB, updates `par_dlk.dat` |
| Mandatory insurance | `BituahRechev` | image, MDB, updates `par_rech.dat` |
| Indices | `Madadim` | `madadimMMYY.txt` |

`par_dlk.dat` and `par_rech.dat` are read from `p:\kolnatun\updates\paramPro` and written to the local output folder.
