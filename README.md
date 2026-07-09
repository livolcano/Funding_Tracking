# funding_tracking_mvp

A lightweight MVP tool for tracking KAKEN funding information.

Current scope: this project currently supports only Japan KAKEN government-funded project information.

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

## For Non-Coders: Let GitHub Copilot Do The Setup

If you are not comfortable with command lines, you can ask GitHub Copilot in VS Code to do all setup steps for you.

1. Open this repository folder in VS Code.
2. Open Copilot Chat.
3. Paste the prompt below and send it.

```text
Please initialize and run this project for me in this exact workspace.

Goal:
- Automatically set up Python environment
- Install dependencies
- Run one smoke test command
- Show me a short summary of what happened

Required actions:
1) Detect whether Python 3.14 is available.
  - If available, create `.venv` with Python 3.14.
  - If not available, use the best available Python 3.x and clearly tell me which version was used.
2) Activate the virtual environment in terminal.
3) Install dependencies from `requirements.txt`.
4) Run this smoke test command once:
  - `python main.py --keyword-query terahertz --page-size 20 --max-pages 1 --project-top-n-per-keyword 5 --project-list-excel output/smoke_project_list.xlsx`
5) Verify that `output/smoke_project_list.xlsx` was generated.
6) At the end, give me:
  - environment Python version,
  - whether install succeeded,
  - whether smoke test succeeded,
  - output file path.

Do not modify source code unless required to make the setup runnable.
```

Tip: If Copilot asks for permission to run terminal commands, click Allow.

### Mode A: Keyword Project List

```bash
python main.py \
  --keyword-query 6G \
  --page-size 20 \
  --max-pages 1 \
  --project-top-n-per-keyword 5 \
  --project-list-excel output/kaken_project_list.xlsx
```

Notes for keyword mode:

- Default suggested keyword in UI is `6G`.
- `--project-top-n-per-keyword` (and UI max result count per keyword) supports `1-200` only.
- After a run, UI shows summary statistics:
  - Total Projects
  - Total Budget (JPY)
  - Total Budget (USD)

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

## Run Web UI (localhost)

```bash
python -m pip install -r requirements.txt
streamlit run web_ui.py --server.address 127.0.0.1 --server.port 8501
```

Then open: `http://127.0.0.1:8501`

The web UI provides the same two modes:

1. Project Search (keyword mode)
2. Researcher Search (manual input or Excel path)

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
  - Includes extra columns:
    - `project_title_en`
    - `project_pi_name_en`
    - `budget_total`
    - `budget_total_usd`

## Notes

- Scope: currently only Japan KAKEN government funding data is supported.
- Roadmap: support for other countries and additional funding sources will be gradually added in future versions.
- KAKEN response structure can change over time; parsing selectors may need updates.
- Use `--delay-seconds` to reduce request pressure.
- Use `--insecure-skip-verify` only in trusted environments.

## Smoke Test Status

A smoke test was run on 2026-07-08 in this workspace:

- `python main.py --help` passed
- keyword query run passed and generated `output/smoke_project_list.xlsx`
