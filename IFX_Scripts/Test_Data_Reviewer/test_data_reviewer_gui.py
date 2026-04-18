from __future__ import annotations

import csv
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import matplotlib
matplotlib.use("TkAgg")
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.widgets import RectangleSelector

import Test_Data_Reviewer as analysis
import stdf_to_flat_csv


METRIC_FILTER_OPTIONS = (
    (analysis.METRIC_YIELD, "Fails"),
    (analysis.METRIC_CPK_LOW, "Cpk < 1.67"),
    (analysis.METRIC_CPK_HIGH, "Cpk > 20"),
    (analysis.METRIC_SITE_DELTA, "Site-to-Site Delta"),
    (analysis.METRIC_UNIQUE_VALUES, "Unique Values"),
    (analysis.METRIC_SKEWNESS, "Skewness"),
    (analysis.METRIC_MULTIMODALITY, "Multimodality"),
)


def _progress_status_text(phase: str, payload: dict[str, object]) -> str:
    if phase == "input_selection_start":
        total_selected = int(payload.get("total_selected", 0))
        return f"Resolving {total_selected} selected input file(s)..."
    if phase == "input_source_reused":
        file_name = str(payload.get("file_name", ""))
        csv_name = str(payload.get("csv_name", ""))
        return f"Reusing existing CSV for {file_name}: {csv_name}"
    if phase == "input_source_converting":
        file_name = str(payload.get("file_name", ""))
        index = int(payload.get("source_index", 0))
        total = int(payload.get("total_sources", 0))
        return f"Converting source file {index}/{total}: {file_name}"
    if phase == "input_source_converted":
        file_name = str(payload.get("file_name", ""))
        return f"Converted source file: {file_name}"
    if phase == "input_selection_done":
        total_csv = int(payload.get("total_csv", 0))
        return f"Selected input resolution complete: {total_csv} CSV file(s) ready"
    if phase == "starting":
        total_files = int(payload.get("total_files", 0))
        return f"Preparing review dataset from {total_files} CSV file(s)..."
    if phase == "file_start":
        file_index = int(payload.get("file_index", 0))
        total_files = int(payload.get("total_files", 0))
        file_name = str(payload.get("file_name", ""))
        return f"Reading file {file_index}/{total_files}: {file_name}"
    if phase == "file_loading":
        file_name = str(payload.get("file_name", ""))
        candidate_tests = int(payload.get("candidate_tests", 0))
        return f"Loading raw measurement data from {file_name} ({candidate_tests} candidate tests)"
    if phase == "file_loaded":
        file_name = str(payload.get("file_name", ""))
        affected_tests = int(payload.get("affected_tests", 0))
        return f"Assessing problematic tests in {file_name} ({affected_tests} affected tests)"
    if phase == "test_progress":
        file_name = str(payload.get("file_name", ""))
        test_index = int(payload.get("test_index", 0))
        total_tests = int(payload.get("total_tests", 0))
        test_col = str(payload.get("test_col", ""))
        return f"Processing test {test_index}/{total_tests} in {file_name}: {test_col}"
    if phase == "file_done":
        file_name = str(payload.get("file_name", ""))
        findings_so_far = int(payload.get("findings_so_far", 0))
        return f"Finished {file_name}; findings collected so far: {findings_so_far}"
    if phase == "completed":
        total_findings = int(payload.get("total_findings", 0))
        return f"Analysis complete: {total_findings} problematic test(s) collected"
    return phase


class AnalysisProgressDialog(tk.Toplevel):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master)
        self.title("Preparing review dataset")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", lambda: None)

        container = ttk.Frame(self, padding=14)
        container.pack(fill="both", expand=True)

        self.message_var = tk.StringVar(value="Initializing analysis...")
        self.detail_var = tk.StringVar(value="")

        ttk.Label(container, textvariable=self.message_var, font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(container, textvariable=self.detail_var, foreground="#555555").pack(anchor="w", pady=(6, 10))
        self.progress = ttk.Progressbar(container, orient="horizontal", mode="indeterminate", length=420)
        self.progress.pack(fill="x")
        self.progress.start(12)

        self.update_idletasks()

    def update_status(self, phase: str, payload: dict[str, object]) -> None:
        status_text = _progress_status_text(phase, payload)
        if phase == "starting":
            total_files = int(payload.get("total_files", 0))
            self.message_var.set("Preparing review dataset...")
            self.detail_var.set(f"Scanning {total_files} input file(s).")
        elif phase == "file_start":
            file_index = int(payload.get("file_index", 0))
            total_files = int(payload.get("total_files", 0))
            file_name = str(payload.get("file_name", ""))
            self.message_var.set(f"Reading file {file_index}/{total_files}")
            self.detail_var.set(file_name)
        elif phase == "file_loading":
            file_name = str(payload.get("file_name", ""))
            candidate_tests = int(payload.get("candidate_tests", 0))
            self.message_var.set("Loading raw measurement data...")
            self.detail_var.set(f"{file_name} | {candidate_tests} candidate test(s)")
        elif phase == "file_loaded":
            file_name = str(payload.get("file_name", ""))
            affected_tests = int(payload.get("affected_tests", 0))
            self.message_var.set("Assessing problematic tests...")
            self.detail_var.set(f"{file_name} | {affected_tests} problematic test(s)")
        elif phase == "test_progress":
            file_name = str(payload.get("file_name", ""))
            test_index = int(payload.get("test_index", 0))
            total_tests = int(payload.get("total_tests", 0))
            test_col = str(payload.get("test_col", ""))
            self.message_var.set(f"Processing test {test_index}/{total_tests}")
            self.detail_var.set(f"{file_name} | Test {test_col}")
        elif phase == "file_done":
            file_name = str(payload.get("file_name", ""))
            findings_so_far = int(payload.get("findings_so_far", 0))
            self.message_var.set("Finishing current input file...")
            self.detail_var.set(f"{file_name} | {findings_so_far} findings collected so far")
        elif phase == "completed":
            total_findings = int(payload.get("total_findings", 0))
            self.message_var.set("Analysis complete")
            self.detail_var.set(f"Collected {total_findings} problematic test(s).")
        else:
            self.message_var.set(status_text)
            self.detail_var.set("")

        self.update_idletasks()


class DtrInfoWindow(tk.Toplevel):
    def __init__(
        self,
        master: tk.Misc,
        *,
        dtr_tables: list[tuple[str, list[str], list[list[str]]]],
    ) -> None:
        super().__init__(master)
        self.title("DTR infos")
        self.geometry("980x620")
        self.minsize(780, 420)
        self._dtr_tables = dtr_tables

        container = ttk.Frame(self, padding=10)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Input file").grid(row=0, column=0, sticky="w")
        self.selected_file_var = tk.StringVar(value=dtr_tables[0][0])
        selector = ttk.Combobox(
            container,
            state="readonly",
            textvariable=self.selected_file_var,
            values=[item[0] for item in dtr_tables],
            width=72,
        )
        selector.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        selector.bind("<<ComboboxSelected>>", self._on_selection_change)

        table_frame = ttk.Frame(container)
        table_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(10, 0))

        self.table = ttk.Treeview(table_frame, show="headings")
        scroll_y = tk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        scroll_x = tk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.table.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)
        container.columnconfigure(1, weight=1)
        self.update_idletasks()
        self.geometry(f"+{self.master.winfo_rootx() + 140}+{self.master.winfo_rooty() + 140}")

        self._render_selected_table(dtr_tables[0][0])

    def _on_selection_change(self, _event) -> None:
        self._render_selected_table(self.selected_file_var.get())

    def _render_selected_table(self, file_name: str) -> None:
        selected = next(item for item in self._dtr_tables if item[0] == file_name)
        _, headers, rows = selected
        self.table.delete(*self.table.get_children())
        self.table["columns"] = headers
        for header in headers:
            self.table.heading(header, text=header)
            self.table.column(header, anchor="w", width=max(120, min(420, len(header) * 10 + 40)))
        for row in rows:
            self.table.insert("", "end", values=row)

class TestDataReviewerGui(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Test Data Reviewer - GUI Draft")
        self.geometry("1680x980")
        self.minsize(1380, 820)

        self.dataset: analysis.ReviewDataset | None = None
        self._finding_by_iid: dict[str, analysis.ReviewFinding] = {}
        self._plot_data_cache: dict[tuple[str, int], analysis.ReviewPlotData] = {}
        self._hover_targets: dict[object, dict[str, object]] = {}
        self._cdf_selector: RectangleSelector | None = None
        self._cdf_default_xlim: tuple[float, float] | None = None
        self._cdf_default_ylim: tuple[float, float] | None = None
        self.selected_input_paths: list[Path] = []
        self._dtr_info_cache: dict[str, tuple[list[str], list[list[str]]] | None] = {}

        self.input_selection_var = tk.StringVar()
        self.yield_threshold_var = tk.StringVar(value="100.0")
        self.cpk_low_var = tk.StringVar(value="1.67")
        self.cpk_high_var = tk.StringVar(value="20.0")
        self.outlier_mad_var = tk.StringVar(value="6.0")
        self.waterfall_cleanup_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Choose one or more input files and click Analyze Dataset.")
        self.hover_var = tk.StringVar(value="Hover inside a plot to inspect a data point.")
        self.decision_var = tk.StringVar(value=analysis.REVIEW_DECISION_UNREVIEWED)
        self.metric_filter_vars = {
            metric_key: tk.BooleanVar(value=False)
            for metric_key, _ in METRIC_FILTER_OPTIONS
        }

        self._build_layout()
        self._draw_placeholder_plots("Analyze a dataset and select a problematic test.")

    def _build_layout(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        config_frame = ttk.LabelFrame(root, text="Review setup", padding=10)
        config_frame.pack(fill="x")

        ttk.Label(config_frame, text="Input files").grid(row=0, column=0, sticky="w")
        ttk.Entry(config_frame, textvariable=self.input_selection_var, width=100).grid(row=0, column=1, columnspan=4, sticky="ew", padx=(8, 8))
        top_button_frame = ttk.Frame(config_frame)
        top_button_frame.grid(row=0, column=5, columnspan=4, sticky="e")
        ttk.Button(top_button_frame, text="Browse", command=self._browse_input_files).pack(side="left")
        ttk.Button(top_button_frame, text="Analyze Dataset", command=self._analyze_dataset).pack(side="left", padx=(4, 0))
        ttk.Button(top_button_frame, text="Show DTR infos", command=self._show_dtr_infos).pack(side="left", padx=(12, 0))
        ttk.Button(top_button_frame, text="Generate Excel report", command=self._export_report).pack(side="left", padx=(4, 0))

        threshold_frame = ttk.Frame(config_frame)
        threshold_frame.grid(row=1, column=0, columnspan=9, sticky="w", pady=(8, 0))
        ttk.Label(threshold_frame, text="Yield").grid(row=0, column=0, sticky="w")
        ttk.Entry(threshold_frame, textvariable=self.yield_threshold_var, width=8).grid(row=0, column=1, sticky="w", padx=(4, 10))
        ttk.Label(threshold_frame, text="Cpk low").grid(row=0, column=2, sticky="w")
        ttk.Entry(threshold_frame, textvariable=self.cpk_low_var, width=8).grid(row=0, column=3, sticky="w", padx=(4, 10))
        ttk.Label(threshold_frame, text="Cpk high").grid(row=0, column=4, sticky="w")
        ttk.Entry(threshold_frame, textvariable=self.cpk_high_var, width=8).grid(row=0, column=5, sticky="w", padx=(4, 10))
        ttk.Label(threshold_frame, text="Outlier MAD").grid(row=0, column=6, sticky="w")
        ttk.Entry(threshold_frame, textvariable=self.outlier_mad_var, width=8).grid(row=0, column=7, sticky="w", padx=(4, 0))
        ttk.Checkbutton(
            threshold_frame,
            text="Waterfall clean-up",
            variable=self.waterfall_cleanup_var,
        ).grid(row=0, column=8, sticky="w", padx=(18, 0))

        metric_frame = ttk.LabelFrame(config_frame, text="Metric filters", padding=8)
        metric_frame.grid(row=2, column=0, columnspan=8, sticky="ew", pady=(10, 0))
        for index, (metric_key, label) in enumerate(METRIC_FILTER_OPTIONS):
            col_idx = index * 2
            ttk.Checkbutton(
                metric_frame,
                variable=self.metric_filter_vars[metric_key],
                command=self._apply_metric_filters,
            ).grid(row=0, column=col_idx, sticky="w")
            ttk.Label(metric_frame, text=label).grid(row=0, column=col_idx + 1, sticky="w", padx=(2, 12), pady=2)

        ttk.Button(config_frame, text="Reset CDF zoom", command=self._reset_cdf_zoom).grid(row=2, column=8, sticky="e", pady=(10, 0))

        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(5, weight=1)
        config_frame.columnconfigure(8, weight=1)

        content = ttk.Panedwindow(root, orient="horizontal")
        content.pack(fill="both", expand=True, pady=(12, 0))

        left_frame = ttk.Frame(content, padding=(0, 0, 8, 0))
        right_frame = ttk.Frame(content)
        content.add(left_frame, weight=2)
        content.add(right_frame, weight=3)

        file_filter_frame = ttk.LabelFrame(left_frame, text="Input file filter", padding=10)
        file_filter_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(file_filter_frame, text="Select one or more files. No selection means all input files.").pack(anchor="w")
        file_filter_list_frame = ttk.Frame(file_filter_frame)
        file_filter_list_frame.pack(fill="x", pady=(6, 0))
        self.file_filter_listbox = tk.Listbox(
            file_filter_list_frame,
            selectmode="extended",
            exportselection=False,
            height=5,
        )
        file_filter_scroll = tk.Scrollbar(file_filter_list_frame, orient="vertical", command=self.file_filter_listbox.yview)
        self.file_filter_listbox.configure(yscrollcommand=file_filter_scroll.set)
        self.file_filter_listbox.pack(side="left", fill="x", expand=True)
        file_filter_scroll.pack(side="right", fill="y")
        self.file_filter_listbox.bind("<<ListboxSelect>>", self._on_file_filter_change)
        ttk.Button(file_filter_frame, text="Clear file filter", command=self._clear_file_filter).pack(anchor="e", pady=(6, 0))

        module_filter_frame = ttk.LabelFrame(left_frame, text="Test module filter", padding=10)
        module_filter_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(module_filter_frame, text="Modules are detected from the first four characters of each test name.").pack(anchor="w")
        module_filter_list_frame = ttk.Frame(module_filter_frame)
        module_filter_list_frame.pack(fill="x", pady=(6, 0))
        self.module_filter_listbox = tk.Listbox(
            module_filter_list_frame,
            selectmode="extended",
            exportselection=False,
            height=5,
        )
        module_filter_scroll = tk.Scrollbar(module_filter_list_frame, orient="vertical", command=self.module_filter_listbox.yview)
        self.module_filter_listbox.configure(yscrollcommand=module_filter_scroll.set)
        self.module_filter_listbox.pack(side="left", fill="x", expand=True)
        module_filter_scroll.pack(side="right", fill="y")
        self.module_filter_listbox.bind("<<ListboxSelect>>", self._on_module_filter_change)
        ttk.Button(module_filter_frame, text="Clear module filter", command=self._clear_module_filter).pack(anchor="e", pady=(6, 0))

        tree_frame = ttk.LabelFrame(left_frame, text="Problematic tests", padding=8)
        tree_frame.pack(fill="both", expand=True)

        columns = ("decision", "file", "module", "test_nr", "test_name", "status", "priority", "yield", "cpk")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        headings = {
            "decision": "Decision",
            "file": "File",
            "module": "Module",
            "test_nr": "Test Nr",
            "test_name": "Test Name",
            "status": "Status",
            "priority": "Priority",
            "yield": "Yield (%)",
            "cpk": "Cpk",
        }
        widths = {
            "decision": 120,
            "file": 220,
            "module": 80,
            "test_nr": 90,
            "test_name": 260,
            "status": 220,
            "priority": 90,
            "yield": 90,
            "cpk": 80,
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], anchor="w")

        tree_scroll_y = tk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_scroll_x = tk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        right_split = ttk.Panedwindow(right_frame, orient="vertical")
        right_split.pack(fill="both", expand=True)

        plots_frame = ttk.LabelFrame(right_split, text="Interactive plots", padding=8)
        detail_frame = ttk.LabelFrame(right_split, text="Review details", padding=10)
        right_split.add(plots_frame, weight=3)
        right_split.add(detail_frame, weight=2)

        self.figure = Figure(figsize=(11.6, 7.4), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.figure, master=plots_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill="both", expand=True)
        self.canvas.mpl_connect("motion_notify_event", self._on_plot_hover)
        self.canvas.mpl_connect("button_press_event", self._on_plot_click)

        ttk.Label(plots_frame, textvariable=self.hover_var, anchor="w").pack(fill="x", pady=(8, 0))

        ttk.Label(detail_frame, text="Decision").pack(anchor="w")
        decision_combo = ttk.Combobox(detail_frame, state="readonly", textvariable=self.decision_var, values=analysis.REVIEW_DECISIONS)
        decision_combo.pack(fill="x", pady=(4, 10))
        decision_combo.bind("<<ComboboxSelected>>", self._update_selected_decision)

        self.summary_label = ttk.Label(detail_frame, text="No test selected.", justify="left")
        self.summary_label.pack(fill="x", pady=(0, 10))

        ttk.Label(detail_frame, text="Findings").pack(anchor="w")
        self.findings_text = tk.Text(detail_frame, height=10, wrap="word")
        self.findings_text.pack(fill="both", expand=True)

        ttk.Label(detail_frame, text="Reviewer notes").pack(anchor="w", pady=(10, 0))
        self.notes_text = tk.Text(detail_frame, height=6, wrap="word")
        self.notes_text.pack(fill="both", expand=True)
        self.notes_text.bind("<FocusOut>", self._persist_notes)

        status_bar = ttk.Label(root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x", pady=(10, 0))

    def _browse_input_files(self) -> None:
        initial_dir = None
        if self.selected_input_paths:
            initial_dir = str(self.selected_input_paths[0].parent)

        selected = filedialog.askopenfilenames(
            title="Select input files",
            initialdir=initial_dir,
            filetypes=[
                (
                    "Supported input files",
                    (
                        "*.csv",
                        "*.stdf",
                        "*.std",
                        "*.eff",
                        "*.stdf.gz",
                        "*.stdf.xz",
                        "*.stdf.bz2",
                        "*.std.gz",
                        "*.std.xz",
                        "*.std.bz2",
                        "*.stdf.tar.gz",
                        "*.std.tar.gz",
                    ),
                ),
                ("CSV files", ("*.csv",)),
                ("STDF/EFF files", ("*.stdf", "*.std", "*.eff", "*.stdf.gz", "*.stdf.xz", "*.stdf.bz2", "*.std.gz", "*.std.xz", "*.std.bz2", "*.stdf.tar.gz", "*.std.tar.gz")),
                ("All files", ("*.*",)),
            ],
        )
        if not selected:
            return

        self.selected_input_paths = [Path(path) for path in selected]
        self._dtr_info_cache.clear()
        selected_names = [path.name for path in self.selected_input_paths]
        if len(selected_names) <= 3:
            display = "; ".join(selected_names)
        else:
            display = f"{len(selected_names)} files selected | " + "; ".join(selected_names[:3]) + " ..."
        self.input_selection_var.set(display)

    def _analyze_dataset(self) -> None:
        progress_dialog: AnalysisProgressDialog | None = None
        try:
            if not self.selected_input_paths:
                raise ValueError("Choose one or more input files")

            progress_dialog = AnalysisProgressDialog(self)
            self.status_var.set("Preparing selected input files...")

            def _progress_callback(phase: str, payload: dict[str, object]) -> None:
                status_text = _progress_status_text(phase, payload)
                self.status_var.set(status_text)
                self.update_idletasks()
                if progress_dialog is not None and progress_dialog.winfo_exists():
                    progress_dialog.update_status(phase, payload)

            input_folder, selected_csv_paths = self._resolve_selected_analysis_csv_paths(progress_callback=_progress_callback)

            dataset = analysis.collect_review_dataset(
                input_folder=input_folder,
                output_folder=input_folder / "Outputs",
                modules=None,
                outlier_mad_multiplier=float(self.outlier_mad_var.get().strip()),
                yield_threshold=float(self.yield_threshold_var.get().strip()),
                cpk_low=float(self.cpk_low_var.get().strip()),
                cpk_high=float(self.cpk_high_var.get().strip()),
                max_files=None,
                single_file=None,
                selected_csv_paths=selected_csv_paths,
                encoding=analysis.DEFAULT_ENCODING,
                progress_callback=_progress_callback,
                waterfall_cleanup_enabled=self.waterfall_cleanup_var.get(),
            )
        except Exception as exc:
            if progress_dialog is not None and progress_dialog.winfo_exists():
                progress_dialog.destroy()
            self.status_var.set(f"Analysis failed: {exc}")
            messagebox.showerror("Analysis failed", str(exc))
            return

        if progress_dialog is not None and progress_dialog.winfo_exists():
            progress_dialog.destroy()

        self.dataset = dataset
        self._plot_data_cache.clear()
        self._finding_by_iid.clear()
        self._populate_file_filter_list(dataset.processed_files)
        self._populate_module_filter_list(dataset.available_modules)
        self._apply_metric_filters()
        total_filters = len(dataset.findings)
        cleanup_state = "on" if dataset.waterfall_cleanup_enabled else "off"
        self.status_var.set(
            f"Loaded {total_filters} problematic test(s) from {len(dataset.processed_files)} processed file(s); {len(dataset.available_modules)} module(s) detected; waterfall clean-up {cleanup_state}."
        )

    def _selected_std_input_paths(self) -> list[Path]:
        std_paths: list[Path] = []
        for path in self.selected_input_paths:
            if path.suffix.lower() == ".csv":
                continue
            if stdf_to_flat_csv._source_kind_for_name(path.name) != "stdf":
                continue
            std_paths.append(path)
        return std_paths

    def _load_dtr_table_for_source(self, source_path: Path) -> tuple[list[str], list[list[str]]] | None:
        cache_key = str(source_path.resolve()).lower()
        if cache_key in self._dtr_info_cache:
            return self._dtr_info_cache[cache_key]

        artifacts_folder = source_path.parent / "Outputs" / "Artifacts"
        dtr_path = artifacts_folder / stdf_to_flat_csv.dtr_name_for_source(source_path.name)
        if not dtr_path.is_file():
            with tempfile.TemporaryDirectory(prefix="tdr_dtr_") as tmp_dir:
                tmp_root = Path(tmp_dir)
                summary = stdf_to_flat_csv.convert_stdf_file(
                    source_path,
                    tmp_root / stdf_to_flat_csv.csv_name_for_source(source_path.name),
                    artifacts_output_folder=tmp_root,
                )
                if not summary.dtr_files:
                    self._dtr_info_cache[cache_key] = None
                    return None
                dtr_path = summary.dtr_files[0]
                with dtr_path.open("r", encoding="utf-8", newline="") as handle:
                    reader = csv.reader(handle, delimiter=stdf_to_flat_csv.DELIMITER)
                    rows = list(reader)
                if not rows:
                    self._dtr_info_cache[cache_key] = None
                    return None
                headers = [str(item) for item in rows[0]]
                body = [[str(item) for item in row] for row in rows[1:]]
                self._dtr_info_cache[cache_key] = (headers, body)
                return self._dtr_info_cache[cache_key]

        with dtr_path.open("r", encoding="utf-8", newline="") as handle:
            reader = csv.reader(handle, delimiter=stdf_to_flat_csv.DELIMITER)
            rows = list(reader)
        if not rows:
            self._dtr_info_cache[cache_key] = None
            return None
        headers = [str(item) for item in rows[0]]
        body = [[str(item) for item in row] for row in rows[1:]]
        self._dtr_info_cache[cache_key] = (headers, body)
        return self._dtr_info_cache[cache_key]

    def _show_dtr_infos(self) -> None:
        std_input_paths = self._selected_std_input_paths()
        if not std_input_paths:
            messagebox.showwarning(
                "No DTR infos available",
                "No DTR infos could be extracted from the selected input files. STD input files are required for that.",
            )
            return

        dtr_tables: list[tuple[str, list[str], list[list[str]]]] = []
        for source_path in std_input_paths:
            table = self._load_dtr_table_for_source(source_path)
            if table is None:
                continue
            headers, rows = table
            dtr_tables.append((source_path.name, headers, rows))

        if not dtr_tables:
            messagebox.showwarning(
                "No DTR infos available",
                "No DTR infos could be extracted from the selected input files. STD input files are required for that.",
            )
            return

        DtrInfoWindow(self, dtr_tables=dtr_tables)

    def _existing_csv_for_selected_source(self, source_path: Path) -> Path | None:
        candidate_paths: list[Path] = []
        stripped_name = source_path.name
        suffixes = Path(stripped_name).suffixes
        while suffixes and suffixes[-1].lower() in (
            stdf_to_flat_csv.COMPRESSED_SUFFIXES
            | stdf_to_flat_csv.ARCHIVE_SUFFIXES
            | stdf_to_flat_csv.SUPPORTED_INPUT_SUFFIXES
        ):
            stripped_name = Path(stripped_name).with_suffix("").name
            candidate_paths.append(source_path.parent / f"{stripped_name}.csv")
            suffixes = Path(stripped_name).suffixes

        candidate_paths.append(source_path.parent / stdf_to_flat_csv.csv_name_for_source(source_path.name))

        seen: set[str] = set()
        for candidate_path in candidate_paths:
            key = str(candidate_path.resolve()).lower()
            if key in seen:
                continue
            seen.add(key)
            if candidate_path.is_file() and analysis._is_analysis_input_csv_path(candidate_path):
                return candidate_path
        return None

    def _resolve_selected_analysis_csv_paths(
        self,
        progress_callback: callable | None = None,
    ) -> tuple[Path, list[Path]]:
        selected_paths = [Path(path) for path in self.selected_input_paths]
        if not selected_paths:
            raise ValueError("Choose one or more input files")

        if progress_callback is not None:
            progress_callback(
                "input_selection_start",
                {"total_selected": len(selected_paths)},
            )

        parent_folders = {path.parent.resolve() for path in selected_paths}
        if len(parent_folders) != 1:
            raise ValueError("Select input files from the same folder")

        input_folder = next(iter(parent_folders))
        csv_paths: list[Path] = []
        pending_source_paths: list[Path] = []

        for path in selected_paths:
            if not path.is_file():
                raise FileNotFoundError(f"Selected input file not found: {path}")
            if path.suffix.lower() == ".csv":
                if not analysis._is_analysis_input_csv_path(path):
                    continue
                csv_paths.append(path)
                continue

            existing_csv = self._existing_csv_for_selected_source(path)
            if existing_csv is not None and existing_csv.is_file():
                csv_paths.append(existing_csv)
                if progress_callback is not None:
                    progress_callback(
                        "input_source_reused",
                        {
                            "file_name": path.name,
                            "csv_name": existing_csv.name,
                        },
                    )
            else:
                pending_source_paths.append(path)

        if pending_source_paths:
            artifacts_folder = input_folder / "Outputs" / "Artifacts"
            converted_csv_paths: list[Path] = []
            total_sources = len(pending_source_paths)
            for source_index, source_path in enumerate(pending_source_paths, start=1):
                if progress_callback is not None:
                    progress_callback(
                        "input_source_converting",
                        {
                            "file_name": source_path.name,
                            "source_index": source_index,
                            "total_sources": total_sources,
                        },
                    )
                try:
                    summary = stdf_to_flat_csv.convert_stdf_file(
                        source_path,
                        input_folder / stdf_to_flat_csv.csv_name_for_source(source_path.name),
                        artifacts_output_folder=artifacts_folder,
                    )
                except Exception as exc:
                    reason = str(exc).strip() or exc.__class__.__name__
                    raise RuntimeError(
                        f"Failed to convert selected input file '{source_path.name}' to CSV. Reason: {reason}"
                    ) from exc
                converted_csv_paths.extend(summary.output_files)
                if progress_callback is not None:
                    progress_callback(
                        "input_source_converted",
                        {"file_name": source_path.name},
                    )

            if not converted_csv_paths:
                raise RuntimeError("No STDF/EFF input files were converted into CSV")
            csv_paths.extend(converted_csv_paths)

        unique_csv_paths: list[Path] = []
        seen_paths: set[str] = set()
        for path in csv_paths:
            resolved = str(path.resolve()).lower()
            if resolved in seen_paths:
                continue
            seen_paths.add(resolved)
            unique_csv_paths.append(path)

        if not unique_csv_paths:
            raise RuntimeError("No supported CSV/STDF/EFF input files were selected")

        if progress_callback is not None:
            progress_callback(
                "input_selection_done",
                {"total_csv": len(unique_csv_paths)},
            )

        return input_folder, unique_csv_paths

    def _selected_finding(self) -> analysis.ReviewFinding | None:
        selected = self.tree.selection()
        if not selected:
            return None
        return self._finding_by_iid.get(selected[0])

    def _cleanup_summary_for_finding(
        self,
        finding: analysis.ReviewFinding,
    ) -> analysis.ReviewCleanupFunctionSummary | None:
        if self.dataset is None:
            return None
        file_summaries = self.dataset.cleanup_summaries_by_file.get(finding.file_name, ())
        for summary in file_summaries:
            if int(summary.function_index) == int(finding.function_index):
                return summary
        return None

    def _active_metric_filters(self) -> set[str]:
        return {
            metric_key
            for metric_key, variable in self.metric_filter_vars.items()
            if variable.get()
        }

    def _selected_file_filters(self) -> set[str]:
        selected_indices = self.file_filter_listbox.curselection()
        return {str(self.file_filter_listbox.get(index)) for index in selected_indices}

    def _selected_module_filters(self) -> set[str]:
        selected_indices = self.module_filter_listbox.curselection()
        return {str(self.module_filter_listbox.get(index)) for index in selected_indices}

    def _populate_file_filter_list(self, file_names: tuple[str, ...]) -> None:
        self.file_filter_listbox.delete(0, "end")
        for file_name in file_names:
            self.file_filter_listbox.insert("end", file_name)

    def _populate_module_filter_list(self, module_names: tuple[str, ...]) -> None:
        self.module_filter_listbox.delete(0, "end")
        for module_name in module_names:
            self.module_filter_listbox.insert("end", module_name)

    def _clear_file_filter(self) -> None:
        self.file_filter_listbox.selection_clear(0, "end")
        self._apply_metric_filters()

    def _clear_module_filter(self) -> None:
        self.module_filter_listbox.selection_clear(0, "end")
        self._apply_metric_filters()

    def _on_file_filter_change(self, _event) -> None:
        self._apply_metric_filters()

    def _on_module_filter_change(self, _event) -> None:
        self._apply_metric_filters()

    def _filtered_findings(self) -> list[analysis.ReviewFinding]:
        if self.dataset is None:
            return []
        active_filters = self._active_metric_filters()
        selected_files = self._selected_file_filters()
        selected_modules = self._selected_module_filters()
        filtered = list(self.dataset.findings)
        if selected_files:
            filtered = [finding for finding in filtered if finding.file_name in selected_files]
        if selected_modules:
            filtered = [finding for finding in filtered if finding.module in selected_modules]
        if active_filters:
            filtered = [
                finding
                for finding in filtered
                if not active_filters.isdisjoint(set(finding.metric_keys))
            ]
        return filtered

    def _apply_metric_filters(self) -> None:
        selected_finding = self._selected_finding()
        selected_key = None
        if selected_finding is not None:
            selected_key = (selected_finding.file_name, selected_finding.test_col)

        for item_id in self.tree.get_children():
            self.tree.delete(item_id)
        self._finding_by_iid.clear()

        filtered_findings = self._filtered_findings()
        for index, finding in enumerate(filtered_findings, start=1):
            item_id = f"finding-{index}"
            self._finding_by_iid[item_id] = finding
            self.tree.insert(
                "",
                "end",
                iid=item_id,
                values=(
                    finding.decision,
                    finding.file_name,
                    finding.module,
                    finding.test_col,
                    finding.test_name,
                    finding.status,
                    finding.priority,
                    "" if finding.yield_pct is None else f"{finding.yield_pct:.1f}",
                    "" if finding.cpk is None else f"{finding.cpk:.2f}",
                ),
            )

        if self.dataset is not None:
            active_filters = self._active_metric_filters()
            selected_files = self._selected_file_filters()
            selected_modules = self._selected_module_filters()
            metric_suffix = "all metrics" if not active_filters else f"{len(active_filters)} metric filter(s) active"
            file_suffix = "all files" if not selected_files else f"{len(selected_files)} file filter(s) active"
            module_suffix = "all modules" if not selected_modules else f"{len(selected_modules)} module filter(s) active"
            self.status_var.set(
                f"Showing {len(filtered_findings)} of {len(self.dataset.findings)} problematic test(s); {metric_suffix}; {file_suffix}; {module_suffix}."
            )

        preferred_item_id = None
        if selected_key is not None:
            for item_id, finding in self._finding_by_iid.items():
                if (finding.file_name, finding.test_col) == selected_key:
                    preferred_item_id = item_id
                    break

        if preferred_item_id is None and self.tree.get_children():
            preferred_item_id = self.tree.get_children()[0]

        if preferred_item_id is not None:
            self.tree.selection_set(preferred_item_id)
            self.tree.focus(preferred_item_id)
            self._on_tree_select(None)
        else:
            self.summary_label.configure(text="No problematic test matches the current file, module, and metric filters.")
            self.findings_text.delete("1.0", "end")
            self.notes_text.delete("1.0", "end")
            self._draw_placeholder_plots("No problematic test matches the current file, module, and metric filters.")

    def _on_tree_select(self, _event) -> None:
        finding = self._selected_finding()
        if finding is None:
            return

        self.decision_var.set(analysis._normalize_review_decision(finding.decision))
        cleanup_summary = self._cleanup_summary_for_finding(finding)
        cleanup_lines = [
            f"Function: {finding.function_index} - {finding.function_label or 'N/A'}",
            f"Cleanup before function: {finding.cleanup_removed_before_review} chip(s) already removed",
        ]
        if cleanup_summary is not None:
            cleanup_lines.append(
                f"Cleanup after function: removed {cleanup_summary.removed_chip_count} chip(s) | remaining active chips: {cleanup_summary.remaining_chip_count}"
            )

        self.summary_label.configure(
            text=(
                f"File: {finding.file_name}\n"
                f"Sheet: {finding.sheet_name}\n"
                f"Test: {finding.test_col} - {finding.test_name}\n"
                f"Status: {finding.status}\n"
                f"Priority: {finding.priority}\n"
                f"Temp: {finding.temp_label}\n"
                f"Fails: {finding.fail_chips} | Outliers: {finding.outliers} | N: {finding.sample_count}\n"
                + "\n".join(cleanup_lines)
            )
        )
        self.findings_text.delete("1.0", "end")
        self.findings_text.insert("1.0", finding.findings or "")
        self.notes_text.delete("1.0", "end")
        self.notes_text.insert("1.0", finding.te_notes or "")
        self._render_selected_plots(finding)

    def _update_selected_decision(self, _event) -> None:
        finding = self._selected_finding()
        if finding is None:
            return
        finding.decision = analysis._normalize_review_decision(self.decision_var.get())
        selected = self.tree.selection()
        if selected:
            values = list(self.tree.item(selected[0], "values"))
            values[0] = finding.decision
            self.tree.item(selected[0], values=values)

    def _persist_notes(self, _event) -> None:
        finding = self._selected_finding()
        if finding is None:
            return
        finding.te_notes = self.notes_text.get("1.0", "end").strip()

    def _plot_cache_key(self, finding: analysis.ReviewFinding) -> tuple[str, int]:
        return (str(finding.file_path), int(finding.test_col))

    def _get_plot_data(self, finding: analysis.ReviewFinding) -> analysis.ReviewPlotData:
        cache_key = self._plot_cache_key(finding)
        plot_data = self._plot_data_cache.get(cache_key)
        if plot_data is None:
            file_cache = None
            if self.dataset is not None:
                file_cache = self.dataset.file_plot_caches.get(str(finding.file_path))
            plot_data = analysis.load_review_plot_data(
                finding,
                file_cache=file_cache,
                encoding=self.dataset.encoding if self.dataset is not None else analysis.DEFAULT_ENCODING,
            )
            self._plot_data_cache[cache_key] = plot_data
        return plot_data

    def _draw_placeholder_plots(self, message: str) -> None:
        self.figure.clear()
        axis = self.figure.add_subplot(111)
        axis.axis("off")
        axis.text(0.5, 0.5, message, ha="center", va="center", fontsize=12)
        self._hover_targets.clear()
        self._cdf_selector = None
        self._cdf_default_xlim = None
        self._cdf_default_ylim = None
        self.hover_var.set("Hover inside a plot to inspect a data point.")
        self.figure.subplots_adjust(left=0.03, right=0.97, top=0.96, bottom=0.06)
        self.canvas.draw_idle()

    def _render_selected_plots(self, finding: analysis.ReviewFinding) -> None:
        try:
            plot_data = self._get_plot_data(finding)
        except Exception as exc:
            self._draw_placeholder_plots(f"Plot generation failed: {exc}")
            return

        self.figure.clear()
        grid = self.figure.add_gridspec(2, 2, height_ratios=[1.0, 1.08], hspace=0.32, wspace=0.22)
        cdf_axis = self.figure.add_subplot(grid[0, 0])
        scatter_axis = self.figure.add_subplot(grid[0, 1])
        wafer_axis = self.figure.add_subplot(grid[1, :])

        self._hover_targets.clear()
        self._render_cdf_axis(cdf_axis, plot_data)
        self._render_scatter_axis(scatter_axis, plot_data)
        self._render_wafer_axis(wafer_axis, plot_data)
        self._attach_cdf_zoom(cdf_axis)

        self.figure.subplots_adjust(left=0.06, right=0.97, top=0.95, bottom=0.08, wspace=0.22, hspace=0.34)
        self.canvas.draw_idle()

    def _render_cdf_axis(self, axis, plot_data: analysis.ReviewPlotData) -> None:
        values = np.sort(np.asarray(plot_data.finite_values, dtype=float))
        y_values = 100.0 * np.arange(1, values.size + 1) / values.size
        low_limit = plot_data.low_limit
        high_limit = plot_data.high_limit
        low = -np.inf if low_limit is None else float(low_limit)
        high = np.inf if high_limit is None else float(high_limit)
        fail_mask = (values < low) | (values > high) if (low_limit is not None or high_limit is not None) else np.zeros(values.size, dtype=bool)
        pass_mask = ~fail_mask

        title = analysis._build_plot_title(
            test_name=plot_data.finding.test_name,
            test_col=str(plot_data.finding.test_col),
            temp_label=plot_data.finding.temp_label,
            cpk=plot_data.finding.cpk,
            mean_v=plot_data.mean_value,
            median_v=plot_data.median_value,
        )
        if np.any(pass_mask):
            axis.scatter(values[pass_mask], y_values[pass_mask], s=16, alpha=0.82, color="#1F77B4", label=f"Pass={int(np.count_nonzero(pass_mask))}")
        if np.any(fail_mask):
            axis.scatter(values[fail_mask], y_values[fail_mask], s=20, alpha=0.95, color="#D62728", label=f"Fail={int(np.count_nonzero(fail_mask))}")
        if low_limit is not None:
            axis.axvline(float(low_limit), color="#D62728", linestyle="-", linewidth=1.4, label=f"LTL={analysis._fmt_num(float(low_limit))}")
        if high_limit is not None:
            axis.axvline(float(high_limit), color="#D62728", linestyle="-", linewidth=1.4, label=f"UTL={analysis._fmt_num(float(high_limit))}")
        if plot_data.finding.ltl_6s is not None:
            axis.axvline(float(plot_data.finding.ltl_6s), color="#FF7F0E", linestyle="--", linewidth=1.1)
        if plot_data.finding.utl_6s is not None:
            axis.axvline(float(plot_data.finding.utl_6s), color="#FF7F0E", linestyle="--", linewidth=1.1)
        if plot_data.finding.ltl_12s is not None:
            axis.axvline(float(plot_data.finding.ltl_12s), color="#9467BD", linestyle=":", linewidth=1.1)
        if plot_data.finding.utl_12s is not None:
            axis.axvline(float(plot_data.finding.utl_12s), color="#9467BD", linestyle=":", linewidth=1.1)

        analysis._apply_probability_percent_axis(axis)
        axis.set_title(title, fontsize=10)
        axis.set_xlabel("Value")
        axis.grid(True, alpha=0.25)
        axis.legend(loc="best", fontsize=8, framealpha=0.9)
        self._cdf_default_xlim = axis.get_xlim()
        self._cdf_default_ylim = axis.get_ylim()
        self._register_hover_target(
            axis,
            x_values=values,
            y_values=y_values,
            formatter=lambda idx: (
                f"CDF | value={analysis._fmt_num(float(values[idx]))} | percentile={y_values[idx]:.2f}% | "
                f"status={'FAIL' if fail_mask[idx] else 'PASS'}"
            ),
        )

    def _render_scatter_axis(self, axis, plot_data: analysis.ReviewPlotData) -> None:
        values = np.asarray(plot_data.finite_values, dtype=float)
        x_values = np.arange(values.size, dtype=int)
        low_limit = plot_data.low_limit
        high_limit = plot_data.high_limit
        low = -np.inf if low_limit is None else float(low_limit)
        high = np.inf if high_limit is None else float(high_limit)
        fail_mask = (values < low) | (values > high) if (low_limit is not None or high_limit is not None) else np.zeros(values.size, dtype=bool)
        pass_mask = ~fail_mask

        if np.any(pass_mask):
            axis.scatter(x_values[pass_mask], values[pass_mask], s=13, alpha=0.75, color="#4F81BD", label="In spec")
        if np.any(fail_mask):
            axis.scatter(x_values[fail_mask], values[fail_mask], s=19, alpha=0.92, color="#C0504D", label="Out of spec")
        if low_limit is not None:
            axis.axhline(float(low_limit), color="#D62728", linestyle="--", linewidth=1.1)
        if high_limit is not None:
            axis.axhline(float(high_limit), color="#D62728", linestyle="--", linewidth=1.1)

        axis.set_title("Scatter plot", fontsize=10)
        axis.set_xlabel("Device number")
        axis.set_ylabel("Test value")
        axis.grid(True, alpha=0.18)
        if np.any(pass_mask) or np.any(fail_mask):
            axis.legend(loc="best", fontsize=8)
        self._register_hover_target(
            axis,
            x_values=x_values.astype(float),
            y_values=values,
            formatter=lambda idx: (
                f"Scatter | device={int(x_values[idx])} | value={analysis._fmt_num(float(values[idx]))} | "
                f"status={'FAIL' if fail_mask[idx] else 'PASS'}"
            ),
        )

    def _render_wafer_axis(self, axis, plot_data: analysis.ReviewPlotData) -> None:
        if not analysis._supports_wafer_maps(plot_data.finding.file_name):
            axis.axis("off")
            axis.text(0.5, 0.5, "Wafer map not available for this file naming pattern.", ha="center", va="center")
            return

        prepared = analysis._prepare_wafer_map_frame(plot_data.numeric_series, meta_cols=plot_data.meta_cols)
        df, wafers, vmin, vmax, warning_text = prepared
        if df is None or wafers is None or len(wafers) == 0:
            axis.axis("off")
            axis.text(0.5, 0.5, "Wafer map data unavailable for the selected test.", ha="center", va="center")
            return

        try:
            cmap = matplotlib.colormaps["turbo"]
        except Exception:
            cmap = matplotlib.colormaps["viridis"]

        low_limit = plot_data.low_limit
        high_limit = plot_data.high_limit
        low = -np.inf if low_limit is None else float(low_limit)
        high = np.inf if high_limit is None else float(high_limit)

        x_arrays: list[np.ndarray] = []
        y_arrays: list[np.ndarray] = []
        hover_lines: list[str] = []
        value_arrays: list[np.ndarray] = []

        x_min = float(df["X"].min()) if not df.empty else 0.0
        x_max = float(df["X"].max()) if not df.empty else 1.0
        x_span = max(4.0, (x_max - x_min) + 4.0)

        scatter_artist = None
        for wafer_index, wafer_name in enumerate(wafers):
            wafer_frame = df[df["WAFER"].astype(str) == str(wafer_name)].copy()
            if wafer_frame.empty:
                continue
            grouped = wafer_frame.groupby(["Y", "X"], as_index=False).agg(v=("v", "median"))
            if grouped.empty:
                continue

            x_offset = wafer_index * x_span
            plot_x = grouped["X"].to_numpy(dtype=float) + x_offset
            plot_y = grouped["Y"].to_numpy(dtype=float)
            plot_v = grouped["v"].to_numpy(dtype=float)
            fail_mask = (plot_v < low) | (plot_v > high) if (low_limit is not None or high_limit is not None) else np.zeros(plot_v.size, dtype=bool)
            edge_colors = np.where(fail_mask, "#C00000", "#666666")

            scatter_artist = axis.scatter(
                plot_x,
                plot_y,
                c=plot_v,
                cmap=cmap,
                vmin=vmin,
                vmax=vmax,
                marker="s",
                s=150,
                linewidths=0.9,
                edgecolors=edge_colors,
            )
            axis.text(
                float(np.mean(plot_x)),
                float(np.max(plot_y)) + 1.1,
                f"Wafer {wafer_name}",
                ha="center",
                va="bottom",
                fontsize=9,
                fontweight="bold",
            )

            x_arrays.append(plot_x)
            y_arrays.append(plot_y)
            value_arrays.append(plot_v)
            hover_lines.extend(
                [
                    (
                        f"Wafer map | wafer={wafer_name} | X={analysis._fmt_wafer_coordinate(orig_x)} | "
                        f"Y={analysis._fmt_wafer_coordinate(orig_y)} | value={analysis._fmt_num(float(value))} | "
                        f"status={'FAIL' if is_fail else 'PASS'}"
                    )
                    for orig_x, orig_y, value, is_fail in zip(
                        grouped["X"].to_numpy(dtype=float),
                        grouped["Y"].to_numpy(dtype=float),
                        plot_v,
                        fail_mask,
                        strict=False,
                    )
                ]
            )

        axis.set_title("Wafer map", fontsize=10)
        axis.set_xlabel("X coordinate")
        axis.set_ylabel("Y coordinate")
        axis.grid(True, alpha=0.16)
        axis.set_aspect("equal", adjustable="datalim")
        if warning_text:
            axis.text(0.01, 0.01, warning_text, transform=axis.transAxes, fontsize=8, color="#A61C00", va="bottom")
        if scatter_artist is not None:
            self.figure.colorbar(scatter_artist, ax=axis, shrink=0.82, pad=0.02, label=plot_data.finding.unit or "Value")

        if x_arrays:
            all_x = np.concatenate(x_arrays)
            all_y = np.concatenate(y_arrays)
            self._register_hover_target(
                axis,
                x_values=all_x,
                y_values=all_y,
                formatter=lambda idx: hover_lines[idx],
            )

    def _register_hover_target(self, axis, *, x_values, y_values, formatter) -> None:
        annotation = axis.annotate(
            "",
            xy=(0, 0),
            xytext=(10, 10),
            textcoords="offset points",
            bbox={"boxstyle": "round,pad=0.25", "fc": "#FFF7D6", "ec": "#666666", "alpha": 0.95},
            arrowprops={"arrowstyle": "->", "color": "#666666"},
        )
        annotation.set_visible(False)
        self._hover_targets[axis] = {
            "x": np.asarray(x_values, dtype=float),
            "y": np.asarray(y_values, dtype=float),
            "annotation": annotation,
            "formatter": formatter,
        }

    def _on_plot_hover(self, event) -> None:
        active_axis = event.inaxes
        redraw = False
        for axis, payload in self._hover_targets.items():
            annotation = payload["annotation"]
            if axis is not active_axis or event.x is None or event.y is None:
                if annotation.get_visible():
                    annotation.set_visible(False)
                    redraw = True
                continue

            x_values = payload["x"]
            y_values = payload["y"]
            if x_values.size == 0:
                if annotation.get_visible():
                    annotation.set_visible(False)
                    redraw = True
                continue

            display_points = axis.transData.transform(np.column_stack([x_values, y_values]))
            distances = np.hypot(display_points[:, 0] - event.x, display_points[:, 1] - event.y)
            best_index = int(np.argmin(distances))
            if float(distances[best_index]) > 18.0:
                if annotation.get_visible():
                    annotation.set_visible(False)
                    redraw = True
                continue

            annotation.xy = (float(x_values[best_index]), float(y_values[best_index]))
            annotation.set_text(payload["formatter"](best_index))
            annotation.set_visible(True)
            self.hover_var.set(payload["formatter"](best_index))
            redraw = True

        if active_axis is None:
            self.hover_var.set("Hover inside a plot to inspect a data point.")
        if redraw:
            self.canvas.draw_idle()

    def _attach_cdf_zoom(self, cdf_axis) -> None:
        self._cdf_selector = RectangleSelector(
            cdf_axis,
            self._on_cdf_zoom_select,
            useblit=False,
            button=[1],
            minspanx=5,
            minspany=5,
            spancoords="pixels",
            interactive=False,
        )

    def _on_cdf_zoom_select(self, eclick, erelease) -> None:
        if eclick.xdata is None or eclick.ydata is None or erelease.xdata is None or erelease.ydata is None:
            return
        if not hasattr(eclick, "inaxes") or eclick.inaxes is None:
            return
        axis = eclick.inaxes
        x0, x1 = sorted([float(eclick.xdata), float(erelease.xdata)])
        y0, y1 = sorted([float(eclick.ydata), float(erelease.ydata)])
        if abs(x1 - x0) < 1e-12 or abs(y1 - y0) < 1e-12:
            return
        axis.set_xlim(x0, x1)
        axis.set_ylim(y0, y1)
        self.canvas.draw_idle()

    def _on_plot_click(self, event) -> None:
        if event.dblclick and event.inaxes is not None:
            self._reset_cdf_zoom()

    def _reset_cdf_zoom(self) -> None:
        if self._cdf_default_xlim is None or self._cdf_default_ylim is None:
            return
        if self.figure.axes:
            cdf_axis = self.figure.axes[0]
            cdf_axis.set_xlim(*self._cdf_default_xlim)
            cdf_axis.set_ylim(*self._cdf_default_ylim)
            self.canvas.draw_idle()

    def _export_report(self) -> None:
        if self.dataset is None:
            messagebox.showinfo("Nothing to export", "Analyze a dataset first.")
            return

        default_path = self.dataset.output_folder / "Test_Data_Reviewer_GUI_Report.xlsx"
        target = filedialog.asksaveasfilename(
            title="Save GUI review workbook",
            defaultextension=".xlsx",
            initialfile=default_path.name,
            initialdir=str(default_path.parent),
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not target:
            return

        self._persist_notes(None)
        try:
            saved_path = analysis.export_review_dataset_workbook(
                self.dataset,
                destination_path=Path(target),
            )
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))
            return

        self.status_var.set(f"Saved GUI review workbook: {saved_path}")
        messagebox.showinfo("Workbook saved", str(saved_path))


def main() -> int:
    app = TestDataReviewerGui()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
