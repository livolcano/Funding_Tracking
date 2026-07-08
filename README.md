# funding_tracking_mvp

A lightweight MVP tool for tracking KAKEN funding information.

It supports two workflows:

- Researcher mode: read researcher names + affiliations from Excel, search KAKEN, and export each researcher's latest funded project.
- Project keyword mode: search funded projects by keyword and export a filtered project list.

The project provides both:

- CLI (`main.py`) for automation
- Desktop UI (`ui.py`) using Tkinter for interactive usage

## What This Project Does

Given researcher or keyword input, the tool can:

- Query KAKEN search pages
- Parse project metadata (title, PI, institution, period, status, detail URL)
- Extract project detail fields such as keywords and budget text
- Convert JPY budget to USD using a public exchange-rate API (researcher mode)
- Save results into CSV / SQLite / Excel

## Project Structure

- `main.py`: core logic + CLI entry
- `ui.py`: Tkinter wizard UI
- `input/researcher_excels/`: default input folder for researcher Excel files
- `output/`: generated outputs

## Requirements

- Python 3.14+
- Internet access (for KAKEN + exchange-rate API)

Dependencies are listed in `requirements.txt` and `pyproject.toml`.

## Quick Start (CLI)

1. Create and activate a virtual environment.
2. Install dependencies.
3. Run one of the following modes.

```bash
python3.14 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Mode A: Keyword Project List

```bash
python main.py \
  --keyword-query terahertz \
  --page-size 20 \
  --max-pages 1 \
  --project-top-n-per-keyword 5 \
  --project-list-excel output/kaken_project_list.xlsx
```

### Mode B: Researcher Excel Search

```bash
python main.py \
  --input-excel input/researcher_excels/your_file.xlsx \
  --name-column researcher_name \
  --affiliation-column affiliation \
  --output-csv output/kaken_latest_projects.csv \
  --output-sqlite output/kaken_latest_projects.db
```

If `--input-excel` is omitted, the tool uses the first Excel file under `input/researcher_excels/`.

## Run UI

```bash
python ui.py
```

Then use the 3-step wizard:

1. Choose search mode (Project or Researcher)
2. Configure source/filters
3. Select output and run search

## Output Files

- Researcher mode:
  - CSV (default: `output/kaken_latest_projects.csv`)
  - SQLite DB (default: `output/kaken_latest_projects.db`)
- Keyword mode:
  - Excel project list (default: `output/kaken_project_list.xlsx`)

## Notes

- KAKEN response structure can change over time; parsing selectors may need updates.
- Use `--delay-seconds` to reduce request pressure.
- Use `--insecure-skip-verify` only in trusted environments.

## Smoke Test Status

A smoke test was run on 2026-07-08 in this workspace:

- `python main.py --help` passed
- keyword query run passed and generated `output/smoke_project_list.xlsx`
