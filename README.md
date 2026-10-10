# Python Data Automation — Fuel, Insurance & Economic Indices

A Python desktop application that collects fuel prices, mandatory car insurance rates, and economic indices through API requests and web scraping, then generates structured files for a recurring monthly update process.

Built from scratch to replace manual data collection and file preparation with a single application.

## Features

The interface provides three tabs, each covering a separate data collection workflow:

| Workflow | Data sources | Output |
|---|---|---|
| Fuel prices | Paz and Delekulator | Text files, Microsoft Access databases, and updated fuel parameter files |
| Mandatory car insurance | Israel's Capital Market Authority (CMA) | Insurance tables as images, Microsoft Access databases, and updated insurance parameter files |
| Economic indices | Israel's Central Bureau of Statistics (CBS) API and US Bureau of Labor Statistics (BLS) | Monthly index files |

Runs are started manually through the desktop interface. Collection tasks run in background threads to keep the interface responsive.

## Technical highlights

- **API integration:** retrieve data through HTTP requests, including the CBS indices API and CMA endpoints.
- **Web scraping:** collect data using Selenium and HTML parsing.
- **File processing:** transform retrieved data into text, image, and database outputs, and update existing parameter files.
- **Desktop interface:** use Tkinter to provide separate workflows for fuel, insurance, and indices.
- **Modular structure:** separate the interface, data collection modules, and configuration.

## Technologies

Python · Tkinter · Requests · Selenium · Beautiful Soup · curl_cffi · Microsoft Access / Windows COM · Matplotlib · PyInstaller

## Requirements

- Python with Tkinter available and a version compatible with `requirements.txt`.
- Windows for the Microsoft Access / COM workflows.
- Microsoft Access installed for workflows that create databases through `Access.Application`.
- Google Chrome for Selenium-based collection.
- Internet access to the external data sources.
- Existing parameter files for workflows that update `par_dlk.dat` and `par_rech.dat`.

## Install and run

Clone or download the repository, then open a terminal in the project folder:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Before running, update these settings in `config.py` to match your environment:

| Setting | Purpose |
|---|---|
| `BASE_PARAMETERS_PATH` | Base folder for generated output files |
| `DELEK_PARAM_SOURCE_PATH` | Folder containing the existing `par_dlk.dat` fuel parameter file |
| `BITUAH_RECHEV_PARAM_SOURCE_PATH` | Folder containing the existing `par_rech.dat` insurance parameter file |

The source parameter files are environment-specific inputs and are not included in the repository.

Start the application:

```powershell
python main_app.py
```

Select the relevant tab and start the update workflow.

Alternative launchers are included:

- `run_main.bat` — uses Python on PATH.
- `הפעל_תוכנה.bat` — uses an existing local virtual environment.

Check the launcher paths before using them.

## Generated files

Output folders are created under the configured `BASE_PARAMETERS_PATH`:

| Workflow | Folder | Files |
|---|---|---|
| Fuel | `DELEK` | Text and `.mdb` files; updated `par_dlk.dat` |
| Mandatory insurance | `BituahRechev` | Table images and `.mdb` files; updated `par_rech.dat` |
| Indices | `Madadim` | `madadimMMYY.txt` |

Existing `.dat` files are read from their configured source folders. Updated files are written to the local output folders.

## Build a Windows executable

An existing PyInstaller specification is included:

```powershell
pip install pyinstaller
pyinstaller Auto070926V2.spec
```

The build produces an application folder under `dist`. Keep its supporting files alongside the executable.

## Project structure

| File or folder | Responsibility |
|---|---|
| `main_app.py` | Tkinter interface and workflow execution |
| `config.py` | Paths, source URLs, file naming, and date settings |
| `UpdateDelek/` | Fuel data collection and file generation |
| `BituahRechev/` | Insurance data collection, tables, and database generation |
| `Madadim/` | Economic index collection and monthly file updates |
| `requirements.txt` | Python dependencies |
| `Auto070926V2.spec` | Windows executable packaging configuration |

## Operational notes

The application depends on external websites and APIs. Changes to their endpoints or page structure may require updates to the collection modules. Workflows are triggered manually through the application.
