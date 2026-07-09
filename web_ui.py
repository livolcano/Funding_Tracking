from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile

import streamlit as st
from openpyxl import Workbook

from main import (
    parse_budget_total_jpy,
    process_keyword_queries,
    process_researchers,
    save_to_sqlite,
    write_csv,
    write_project_list_excel,
)

DEFAULT_PAGE_SIZE = 50
DEFAULT_MAX_PAGES = 2
DEFAULT_DELAY_SECONDS = 1.5
DEFAULT_TOP_N_PER_KEYWORD = 50


def build_temp_excel_from_manual_input(raw_text: str) -> Path:
    rows: list[tuple[str, str]] = []
    for line_no, line in enumerate(raw_text.splitlines(), start=1):
        line = line.strip()
        if not line:
            continue
        if "," not in line:
            raise ValueError(f"Line {line_no} format error. Use: name,affiliation")
        name, affiliation = line.split(",", 1)
        name = name.strip()
        affiliation = affiliation.strip()
        if not name:
            raise ValueError(f"Line {line_no} is missing researcher name.")
        rows.append((name, affiliation))

    if not rows:
        raise ValueError("No valid manual rows found.")

    with NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        temp_path = Path(tmp.name)

    wb = Workbook()
    ws = wb.active
    ws.title = "researchers"
    ws.append(["researcher_name", "affiliation"])
    for name, affiliation in rows:
        ws.append([name, affiliation])
    wb.save(temp_path)

    return temp_path


def render_project_mode() -> None:
    st.subheader("Project Keyword Search")

    keyword_text = st.text_area(
        "Keywords (one per line)",
        value="6G\n",
        height=120,
    )
    top_n_per_keyword = st.number_input(
        "Max result count per keyword",
        min_value=1,
        max_value=200,
        value=50,
        step=1,
    )
    only_not_completed = st.checkbox("Only include projects that are not completed", value=False)
    started_after_year_raw = st.text_input("Project start year is greater than (optional)", value="")
    output_excel = st.text_input("Output Excel", value="output/kaken_project_list.xlsx")

    if st.button("Run Project Search", type="primary"):
        try:
            keywords = [line.strip() for line in keyword_text.splitlines() if line.strip()]
            if not keywords:
                raise ValueError("Please provide at least one keyword.")

            started_after_year = int(started_after_year_raw) if started_after_year_raw.strip() else None
            records = process_keyword_queries(
                keyword_queries=keywords,
                page_size=DEFAULT_PAGE_SIZE,
                max_pages=DEFAULT_MAX_PAGES,
                top_n_per_keyword=int(top_n_per_keyword),
                only_not_completed=only_not_completed,
                started_after_year=started_after_year,
                verify_ssl=False,
                ca_bundle=None,
            )

            output_path = Path(output_excel)
            write_project_list_excel(records, output_path)

            budget_jpy_sum = 0
            budget_usd_sum = 0.0
            for record in records:
                budget_jpy = parse_budget_total_jpy(str(record.get("budget_total", "")))
                if budget_jpy is not None:
                    budget_jpy_sum += budget_jpy
                budget_usd_raw = str(record.get("budget_total_usd", "")).strip()
                if budget_usd_raw:
                    try:
                        budget_usd_sum += float(budget_usd_raw)
                    except ValueError:
                        pass

            st.success(f"Completed. Saved {len(records)} project rows to: {output_path}")
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Projects", len(records))
            col2.metric("Total Budget (JPY)", f"{budget_jpy_sum:,}")
            col3.metric("Total Budget (USD)", f"{budget_usd_sum:,.2f}")
            if records:
                st.dataframe(records)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Project search failed: {exc}")


def render_researcher_mode() -> None:
    st.subheader("Researcher Search")

    source_mode = st.radio("Input source", ["Manual Input", "Upload Excel"], horizontal=True)

    temp_excel: Path | None = None
    input_excel_path: Path | None = None

    if source_mode == "Manual Input":
        manual_text = st.text_area(
            "One researcher per line. Format: name,affiliation",
            value=(
                "Bo Qian,National Institute of Informatics\n"
                "Keping Yu,Hosei University\n"
            ),
            height=150,
        )
        if manual_text.strip():
            try:
                temp_excel = build_temp_excel_from_manual_input(manual_text)
                input_excel_path = temp_excel
            except Exception as exc:  # noqa: BLE001
                st.warning(f"Input not ready: {exc}")
    else:
        excel_path_str = st.text_input("Excel file path", value="input/researcher_excels/researcher_list.xlsx")
        path = Path(excel_path_str)
        if path.exists():
            input_excel_path = path
        else:
            st.warning(f"Excel file not found: {path}")

    output_csv = st.text_input("Output CSV", value="output/kaken_latest_projects.csv")
    output_sqlite = st.text_input("Output SQLite", value="output/kaken_latest_projects.db")

    if st.button("Run Researcher Search", type="primary"):
        if input_excel_path is None:
            st.error("Please provide a valid input source first.")
            return

        try:
            records = process_researchers(
                input_excel=input_excel_path,
                sheet_name=0,
                name_column="researcher_name",
                affiliation_column="affiliation",
                page_size=DEFAULT_PAGE_SIZE,
                max_pages=DEFAULT_MAX_PAGES,
                max_researchers=None,
                delay_seconds=DEFAULT_DELAY_SECONDS,
                verify_ssl=False,
                ca_bundle=None,
            )

            if not records:
                raise RuntimeError("No records were produced.")

            csv_path = Path(output_csv)
            sqlite_path = Path(output_sqlite)
            write_csv(records, csv_path)
            save_to_sqlite(records, sqlite_path)

            st.success(
                "Completed. "
                f"Saved {len(records)} researcher rows to CSV: {csv_path} and SQLite: {sqlite_path}"
            )
            st.dataframe(records)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Researcher search failed: {exc}")
        finally:
            if temp_excel and temp_excel.exists():
                temp_excel.unlink(missing_ok=True)


def main() -> None:
    st.set_page_config(page_title="KAKEN Search Web UI", page_icon="🌐", layout="wide")
    st.title("KAKEN Search Wizard (Web)")
    st.caption("Localhost web UI for project/researcher KAKEN search")

    mode = st.radio("Select Search Mode", ["Project Search", "Researcher Search"], horizontal=True)

    if mode == "Project Search":
        render_project_mode()
    else:
        render_researcher_mode()


if __name__ == "__main__":
    main()
