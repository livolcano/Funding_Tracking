from __future__ import annotations

import threading
from pathlib import Path
from tempfile import NamedTemporaryFile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

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


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("KAKEN Search Wizard")
        self.root.geometry("980x820")

        self.search_mode = tk.StringVar(value="")
        self.source_mode = tk.StringVar(value="manual")

        self.excel_path_var = tk.StringVar(value="")
        self.output_csv_var = tk.StringVar(value="output/kaken_latest_projects.csv")
        self.output_sqlite_var = tk.StringVar(value="output/kaken_latest_projects.db")
        self.project_list_excel_var = tk.StringVar(value="output/kaken_project_list.xlsx")
        self.project_started_after_year_var = tk.StringVar(value="")
        self.project_top_n_per_keyword_var = tk.StringVar(value=str(DEFAULT_TOP_N_PER_KEYWORD))
        self.project_only_not_completed_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Select a mode to begin.")

        self._build_ui()

    def _build_ui(self) -> None:
        self.container = ttk.Frame(self.root, padding=12)
        self.container.pack(fill=tk.BOTH, expand=True)

        self.step1_mode_frame = ttk.LabelFrame(self.container, text="Step 1: Select Search Mode", padding=10)
        self.step1_mode_frame.pack(fill=tk.X)

        ttk.Radiobutton(
            self.step1_mode_frame,
            text="Search Projects",
            value="project",
            variable=self.search_mode,
        ).grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(
            self.step1_mode_frame,
            text="Search Researchers",
            value="researcher",
            variable=self.search_mode,
        ).grid(row=0, column=1, sticky=tk.W, padx=(12, 0))
        ttk.Button(self.step1_mode_frame, text="Continue", command=self._complete_step1).grid(
            row=0, column=2, sticky=tk.E, padx=(12, 0)
        )

        self.step2_project_frame = ttk.LabelFrame(self.container, text="Step 2: Configure Project Search", padding=10)
        self.step2_project_frame.columnconfigure(0, weight=1)
        ttk.Label(
            self.step2_project_frame,
            text="Keywords (one per line)",
        ).grid(row=0, column=0, sticky=tk.W)
        self.keyword_text = tk.Text(self.step2_project_frame, height=6, wrap=tk.WORD)
        self.keyword_text.grid(row=1, column=0, sticky=tk.EW, pady=(6, 0))
        self.keyword_text.insert("1.0", "6G\n")

        project_filter_frame = ttk.Frame(self.step2_project_frame)
        project_filter_frame.grid(row=2, column=0, sticky=tk.EW, pady=(8, 0))
        project_filter_frame.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            project_filter_frame,
            text="Only include projects that are not completed",
            variable=self.project_only_not_completed_var,
        ).grid(row=0, column=0, sticky=tk.W)

        ttk.Label(project_filter_frame, text="Project start year is greater than (optional)").grid(
            row=1, column=0, sticky=tk.W, pady=(8, 0)
        )
        ttk.Entry(project_filter_frame, textvariable=self.project_started_after_year_var).grid(
            row=1, column=1, sticky=tk.EW, padx=(8, 0), pady=(8, 0)
        )

        ttk.Label(project_filter_frame, text="Max results per keyword (1-200)").grid(
            row=2, column=0, sticky=tk.W, pady=(8, 0)
        )
        ttk.Entry(project_filter_frame, textvariable=self.project_top_n_per_keyword_var).grid(
            row=2, column=1, sticky=tk.EW, padx=(8, 0), pady=(8, 0)
        )

        ttk.Button(self.step2_project_frame, text="Continue", command=self._complete_step2_project).grid(
            row=3, column=0, sticky=tk.E, pady=(10, 0)
        )

        self.step2_researcher_frame = ttk.LabelFrame(self.container, text="Step 2: Configure Researcher Search", padding=10)
        self.step2_researcher_frame.columnconfigure(0, weight=1)

        source_switch = ttk.Frame(self.step2_researcher_frame)
        source_switch.grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(
            source_switch,
            text="Manual Input",
            value="manual",
            variable=self.source_mode,
            command=self._toggle_source_mode,
        ).grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(
            source_switch,
            text="Upload Excel",
            value="excel",
            variable=self.source_mode,
            command=self._toggle_source_mode,
        ).grid(row=0, column=1, sticky=tk.W, padx=(12, 0))

        self.manual_frame = ttk.Frame(self.step2_researcher_frame)
        self.manual_frame.grid(row=1, column=0, sticky=tk.EW, pady=(8, 0))
        self.manual_frame.columnconfigure(0, weight=1)

        ttk.Label(
            self.manual_frame,
            text="One researcher per line. Format: name,affiliation",
        ).grid(row=0, column=0, sticky=tk.W)
        self.manual_text = tk.Text(self.manual_frame, height=8, wrap=tk.WORD)
        self.manual_text.grid(row=1, column=0, sticky=tk.EW, pady=(6, 0))
        self.manual_text.insert(
            "1.0",
            "Bo Qian,National Institute of Informatics\n"
            "Keping Yu,Hosei University\n",
        )

        self.excel_frame = ttk.Frame(self.step2_researcher_frame)
        self.excel_frame.grid(row=2, column=0, sticky=tk.EW, pady=(8, 0))
        self.excel_frame.columnconfigure(1, weight=1)

        ttk.Label(self.excel_frame, text="Excel File").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.excel_frame, textvariable=self.excel_path_var).grid(
            row=0, column=1, sticky=tk.EW, padx=(8, 8)
        )
        ttk.Button(self.excel_frame, text="Browse", command=self._choose_excel).grid(row=0, column=2)

        ttk.Button(
            self.step2_researcher_frame,
            text="Continue",
            command=self._complete_step2_researcher,
        ).grid(row=3, column=0, sticky=tk.E, pady=(10, 0))

        self.step3_project_output_frame = ttk.LabelFrame(
            self.container,
            text="Step 3: Project List Output",
            padding=10,
        )
        self.step3_project_output_frame.columnconfigure(1, weight=1)

        ttk.Label(self.step3_project_output_frame, text="Project List Excel").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.step3_project_output_frame, textvariable=self.project_list_excel_var).grid(
            row=0, column=1, sticky=tk.EW, padx=(8, 8)
        )
        ttk.Button(
            self.step3_project_output_frame,
            text="Browse",
            command=self._choose_project_list_excel,
        ).grid(row=0, column=2)
        self.search_project_button = ttk.Button(
            self.step3_project_output_frame,
            text="Search",
            command=self._start_search,
        )
        self.search_project_button.grid(row=1, column=2, sticky=tk.E, pady=(10, 0))

        self.step3_researcher_output_frame = ttk.LabelFrame(
            self.container,
            text="Step 3: Researcher Search Output",
            padding=10,
        )
        self.step3_researcher_output_frame.columnconfigure(1, weight=1)

        ttk.Label(self.step3_researcher_output_frame, text="Output CSV").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.step3_researcher_output_frame, textvariable=self.output_csv_var).grid(
            row=0, column=1, sticky=tk.EW, padx=(8, 8)
        )
        ttk.Button(
            self.step3_researcher_output_frame,
            text="Browse",
            command=self._choose_output_csv,
        ).grid(row=0, column=2)

        ttk.Label(self.step3_researcher_output_frame, text="Output SQLite").grid(
            row=1, column=0, sticky=tk.W, pady=(8, 0)
        )
        ttk.Entry(self.step3_researcher_output_frame, textvariable=self.output_sqlite_var).grid(
            row=1, column=1, sticky=tk.EW, padx=(8, 8), pady=(8, 0)
        )
        ttk.Button(
            self.step3_researcher_output_frame,
            text="Browse",
            command=self._choose_output_sqlite,
        ).grid(row=1, column=2, pady=(8, 0))
        self.search_researcher_button = ttk.Button(
            self.step3_researcher_output_frame,
            text="Search",
            command=self._start_search,
        )
        self.search_researcher_button.grid(row=2, column=2, sticky=tk.E, pady=(10, 0))

        status_frame = ttk.LabelFrame(self.container, text="Status", padding=10)
        status_frame.pack(fill=tk.X, pady=(12, 0))
        ttk.Label(status_frame, textvariable=self.status_var).pack(anchor=tk.W)

        actions = ttk.Frame(self.container)
        actions.pack(fill=tk.X, pady=(12, 0))
        self.exit_button = ttk.Button(actions, text="Exit", command=self._exit_app)
        self.exit_button.pack(side=tk.LEFT)

        self.completion_frame = ttk.Frame(self.container)
        self.restart_button = ttk.Button(
            self.completion_frame,
            text="Start New Search",
            command=self._restart_flow,
        )
        self.restart_button.pack(side=tk.LEFT)
        ttk.Button(self.completion_frame, text="Close Window", command=self._exit_app).pack(
            side=tk.LEFT,
            padx=(8, 0),
        )

    def _hide_step_2_and_3(self) -> None:
        self.step2_project_frame.pack_forget()
        self.step2_researcher_frame.pack_forget()
        self.step3_project_output_frame.pack_forget()
        self.step3_researcher_output_frame.pack_forget()

    def _set_status(self, message: str) -> None:
        self.root.after(0, lambda: self.status_var.set(message))

    def _show_completion_actions(self) -> None:
        self.root.after(0, lambda: self.completion_frame.pack(fill=tk.X, pady=(12, 0)))

    def _hide_completion_actions(self) -> None:
        self.completion_frame.pack_forget()

    def _complete_step1(self) -> None:
        mode = self.search_mode.get()
        if mode not in {"project", "researcher"}:
            messagebox.showwarning("Warning", "Please select a search mode first.")
            return

        self._hide_step_2_and_3()
        self._hide_completion_actions()
        if mode == "project":
            self.step2_project_frame.pack(fill=tk.X, pady=(12, 0))
            self.status_var.set("Project search mode selected. Configure the project search options.")
        else:
            self.step2_researcher_frame.pack(fill=tk.X, pady=(12, 0))
            self._toggle_source_mode()
            self.status_var.set("Researcher search mode selected. Provide researcher input data.")

    def _complete_step2_project(self) -> None:
        keywords = [line.strip() for line in self.keyword_text.get("1.0", tk.END).splitlines() if line.strip()]
        if not keywords:
            messagebox.showwarning("Warning", "Please enter at least one keyword.")
            return
        top_n_raw = self.project_top_n_per_keyword_var.get().strip()
        try:
            top_n = int(top_n_raw)
        except ValueError:
            messagebox.showwarning("Warning", "Max results per keyword must be an integer.")
            return
        if top_n < 1 or top_n > 200:
            messagebox.showwarning("Warning", "Max results per keyword must be between 1 and 200.")
            return

        started_after_year_raw = self.project_started_after_year_var.get().strip()
        if started_after_year_raw:
            try:
                int(started_after_year_raw)
            except ValueError:
                messagebox.showwarning("Warning", "Project start year must be an integer.")
                return

        self.step3_project_output_frame.pack_forget()
        self.step3_project_output_frame.pack(fill=tk.X, pady=(12, 0))
        self.status_var.set("Project search is configured. Choose the output file and click Search.")

    def _complete_step2_researcher(self) -> None:
        if self.source_mode.get() == "excel":
            path = Path(self.excel_path_var.get().strip())
            if not path.exists():
                messagebox.showwarning("Warning", "Please choose an existing Excel file.")
                return
        else:
            raw = self.manual_text.get("1.0", tk.END).strip()
            if not raw:
                messagebox.showwarning("Warning", "Please enter at least one researcher line.")
                return
        self.step3_researcher_output_frame.pack_forget()
        self.step3_researcher_output_frame.pack(fill=tk.X, pady=(12, 0))
        self.status_var.set("Researcher search is configured. Choose output files and click Search.")

    def _toggle_source_mode(self) -> None:
        is_manual = self.source_mode.get() == "manual"
        self.manual_text.configure(state=tk.NORMAL if is_manual else tk.DISABLED)
        for child in self.excel_frame.winfo_children():
            child.configure(state=tk.DISABLED if is_manual else tk.NORMAL)

    def _choose_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose Excel File",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xltx *.xltm")],
        )
        if path:
            self.excel_path_var.set(path)

    def _choose_output_csv(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save CSV As",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if path:
            self.output_csv_var.set(path)

    def _choose_output_sqlite(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save SQLite As",
            defaultextension=".db",
            filetypes=[("SQLite", "*.db")],
        )
        if path:
            self.output_sqlite_var.set(path)

    def _choose_project_list_excel(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save Project List Excel As",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            self.project_list_excel_var.set(path)

    def _exit_app(self) -> None:
        self.root.destroy()

    def _restart_flow(self) -> None:
        self._hide_step_2_and_3()
        self._hide_completion_actions()
        self.search_mode.set("")
        self.status_var.set("Select a mode to begin a new search.")

    def _start_search(self) -> None:
        self.search_project_button.configure(state=tk.DISABLED)
        self.search_researcher_button.configure(state=tk.DISABLED)
        self._hide_completion_actions()
        self.status_var.set("Search is running normally. Please wait...")
        thread = threading.Thread(target=self._run_search, daemon=True)
        thread.start()

    def _run_search(self) -> None:
        temp_excel_path: Path | None = None
        try:
            mode = self.search_mode.get()
            if mode == "project":
                keyword_queries = [
                    line.strip() for line in self.keyword_text.get("1.0", tk.END).splitlines() if line.strip()
                ]
                if not keyword_queries:
                    raise ValueError("Keyword input is empty.")

                started_after_year_raw = self.project_started_after_year_var.get().strip()
                started_after_year = int(started_after_year_raw) if started_after_year_raw else None
                top_n_raw = self.project_top_n_per_keyword_var.get().strip()
                top_n_per_keyword = int(top_n_raw)

                records = process_keyword_queries(
                    keyword_queries=keyword_queries,
                    page_size=DEFAULT_PAGE_SIZE,
                    max_pages=DEFAULT_MAX_PAGES,
                    top_n_per_keyword=top_n_per_keyword,
                    only_not_completed=self.project_only_not_completed_var.get(),
                    started_after_year=started_after_year,
                    verify_ssl=False,
                    ca_bundle=None,
                )

                output_excel = Path(self.project_list_excel_var.get().strip())
                write_project_list_excel(records, output_excel)

                total_budget_jpy = 0
                total_budget_usd = 0.0
                for record in records:
                    budget_jpy = parse_budget_total_jpy(str(record.get("budget_total", "")))
                    if budget_jpy is not None:
                        total_budget_jpy += budget_jpy
                    budget_usd_raw = str(record.get("budget_total_usd", "")).strip()
                    if budget_usd_raw:
                        try:
                            total_budget_usd += float(budget_usd_raw)
                        except ValueError:
                            pass

                self._set_status(
                    "Completed successfully. "
                    f"Project list saved to {output_excel}. "
                    f"Total projects: {len(records)}; "
                    f"Total budget JPY: {total_budget_jpy:,}; "
                    f"Total budget USD: {total_budget_usd:,.2f}"
                )
                self._show_completion_actions()
                return

            input_excel = self._resolve_input_excel()
            temp_excel_path = None if input_excel.exists() else input_excel

            records = process_researchers(
                input_excel=input_excel,
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

            output_csv = Path(self.output_csv_var.get().strip())
            output_sqlite = Path(self.output_sqlite_var.get().strip())

            write_csv(records, output_csv)
            save_to_sqlite(records, output_sqlite)
            self._set_status(
                "Completed successfully. "
                f"CSV saved to {output_csv}; SQLite saved to {output_sqlite}. "
                f"Total researchers: {len(records)}"
            )
            self._show_completion_actions()
        except Exception as exc:  # noqa: BLE001
            self._set_status(f"Search failed: {exc}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            if temp_excel_path and temp_excel_path.exists():
                temp_excel_path.unlink(missing_ok=True)
            self.root.after(0, lambda: self.search_project_button.configure(state=tk.NORMAL))
            self.root.after(0, lambda: self.search_researcher_button.configure(state=tk.NORMAL))

    def _resolve_input_excel(self) -> Path:
        if self.source_mode.get() == "excel":
            path = Path(self.excel_path_var.get().strip())
            if not path.exists():
                raise FileNotFoundError(f"Excel file not found: {path}")
            return path

        raw = self.manual_text.get("1.0", tk.END).strip()
        if not raw:
            raise ValueError("Manual input is empty.")

        rows: list[tuple[str, str]] = []
        for line_no, line in enumerate(raw.splitlines(), start=1):
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


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
