import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import threading
import time
import re

class GrantReportDemoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CAST Grant Reporting Demo (Apricot → Filtered Client List)")
        self.root.geometry("1280x820")

        # Readability
        self.default_font = ("Arial", 12)
        self.header_font = ("Arial", 18, "bold")
        self.root.option_add("*Font", self.default_font)
        self._setup_style()

        self.workbook_path = None
        self.workbook_sheets = {}   # sheet_name -> df
        self.views = {}             # view_name -> df
        self.current_view_name = None
        self.df_current = None
        self.df_display = None

        self.id_column = "Legacy Client ID"
        self.max_preview_rows = 250

        # Grant-focused filters (single-select dropdowns)
        self.filter_vars = {}   # field -> StringVar
        self.filter_combos = {} # field -> Combobox
        self.filter_fields = [
            "Funder",
            "Type of Victimization",
            "Gender",
            "Homelessness",
            "Program",
            "Race/Ethnicity",
            "Victim Type",
            "Age at Time of Trafficking",
            "Veteran Status",
            "LGBTQ/Two-Spirited",
            "Disability",
            "Immigrant Status",
            "Country of Citizenship",
            "Primary Language",
        ]

        # Fields that can contain multi-select values separated by '|', ',' or '&'
        self.multi_select_fields = {
            "Race/Ethnicity",
            "Disability",
            "Victim Type",
            "Homelessness",
            "Age at Time of Trafficking",
        }

        # Keep the preview table “grant-like” (always show these first, if present)
        self.preview_columns = [
            self.id_column,
            "Funder",
            "Type of Victimization",
            "Gender",
            "Homelessness",
            "Program",
            "Race/Ethnicity",
            "Victim Type",
            "Age at Time of Trafficking",
            "Veteran Status",
            "LGBTQ/Two-Spirited",
            "Disability",
            "Immigrant Status",
            "Country of Citizenship",
            "Primary Language",
        ]

        self.create_widgets()

    def _setup_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview.Heading", font=("Arial", 11, "bold"))
        style.configure("Treeview", font=("Arial", 11), rowheight=28)

    def create_widgets(self):
        # Header
        header_frame = tk.Frame(self.root, bg="#2c3e50", height=70)
        header_frame.pack(fill="x")
        header_label = tk.Label(
            header_frame,
            text="CAST Grant Reporting Demo",
            font=self.header_font,
            bg="#2c3e50",
            fg="white",
        )
        header_label.pack(pady=(14, 2))
        header_sub = tk.Label(
            header_frame,
            text="Load Apricot export → de-duplicate clients → filter by grant requirements → preview client list",
            bg="#2c3e50",
            fg="#d0d7de",
        )
        header_sub.pack(pady=(0, 10))

        # File Upload Section
        upload_frame = tk.LabelFrame(self.root, text="1) Load Exported Excel Workbook", padx=10, pady=10)
        upload_frame.pack(fill="x", padx=10, pady=5)
        
        self.btn_upload = tk.Button(upload_frame, text="Choose Excel File…", command=self.load_file)
        self.btn_upload.pack(side="left", padx=5)
        
        self.lbl_status = tk.Label(upload_frame, text="No file loaded", fg="gray")
        self.lbl_status.pack(side="left", padx=10)

        # View selector
        tk.Label(upload_frame, text="Data View:").pack(side="left", padx=(20, 5))
        self.combo_view = ttk.Combobox(upload_frame, state="readonly", width=45)
        self.combo_view.pack(side="left", padx=5)
        self.combo_view.bind("<<ComboboxSelected>>", self.on_view_selected)

        # Status line
        status_frame = tk.Frame(self.root)
        status_frame.pack(fill="x", padx=10, pady=(0, 6))
        self.lbl_pipeline_status = tk.Label(status_frame, text="Status: waiting for workbook…", fg="#333333")
        self.lbl_pipeline_status.pack(side="left")

        # Grant filters
        filter_frame = tk.LabelFrame(self.root, text="2) Grant Filters", padx=10, pady=10)
        filter_frame.pack(fill="x", padx=10, pady=6)

        grid = tk.Frame(filter_frame)
        grid.pack(fill="x")

        cols_per_row = 4
        for idx, field in enumerate(self.filter_fields):
            r = idx // cols_per_row
            c = (idx % cols_per_row) * 2

            tk.Label(grid, text=f"{field}:").grid(row=r, column=c, sticky="w", padx=(0, 6), pady=4)
            var = tk.StringVar(value="All")
            combo = ttk.Combobox(grid, textvariable=var, state="disabled", width=28)
            combo.grid(row=r, column=c + 1, sticky="we", padx=(0, 14), pady=4)
            combo.bind("<<ComboboxSelected>>", self.apply_grant_filters)

            self.filter_vars[field] = var
            self.filter_combos[field] = combo

        controls = tk.Frame(filter_frame)
        controls.pack(fill="x", pady=(8, 0))
        self.btn_clear_grant_filters = tk.Button(
            controls, text="Clear Filters", command=self.clear_grant_filters, state="disabled"
        )
        self.btn_clear_grant_filters.pack(side="left")
        self.lbl_active_filters = tk.Label(controls, text="Active filters: (none)", fg="#555555")
        self.lbl_active_filters.pack(side="left", padx=12)

        # Data Preview Section (Excel-like)
        preview_frame = tk.LabelFrame(self.root, text="3) Client List Preview (Excel-like)", padx=10, pady=10)
        preview_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Treeview for Excel-like display
        self.tree = ttk.Treeview(preview_frame, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)

        scrollbar_y = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        scrollbar_x.pack(side="bottom", fill="x")
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # Footer counts
        footer = tk.Frame(self.root)
        footer.pack(fill="x", padx=10, pady=(0, 10))
        self.lbl_counts = tk.Label(footer, text="Filtered clients: 0 of 0", fg="#111111", font=("Arial", 12, "bold"))
        self.lbl_counts.pack(side="left")
        self.lbl_rows = tk.Label(footer, text="Rows: 0 of 0", fg="#333333")
        self.lbl_rows.pack(side="left", padx=14)

    def log(self, message):
        # Keep the demo clean: show status in a single readable line
        if hasattr(self, "lbl_pipeline_status"):
            self.lbl_pipeline_status.config(text=f"Status: {message}")
        self.root.update()

    def load_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not filename:
            return

        self.workbook_path = filename
        self.lbl_status.config(text=f"Loaded: {filename.split('/')[-1]}", fg="green")
        
        # Start simulated processing thread
        threading.Thread(target=self.process_file, args=(filename,), daemon=True).start()

    def process_file(self, filename):
        try:
            self.log("Reading Apricot Workbook (all tabs)...")
            xl = pd.ExcelFile(filename)
            time.sleep(1) # Fake delay for effect

            # Load all sheets
            self.workbook_sheets = {}
            for sheet in xl.sheet_names:
                try:
                    df = xl.parse(sheet)
                    self.workbook_sheets[sheet] = df
                except Exception:
                    # If a sheet fails, skip it for demo purposes
                    continue

            self.log(f"Loaded {len(self.workbook_sheets)} sheets.")
            self.log("Understanding tabs (columns + common jumbled fields)...")
            time.sleep(0.5)

            # Build views:
            # - Individual sheet views
            # - A combined/merged view (like Fake Report)
            self.views = {}
            for sheet_name, df in self.workbook_sheets.items():
                self.views[f"Sheet: {sheet_name}"] = df

            merged_view = self.build_merged_report_view(xl)
            if merged_view is not None:
                self.views["Merged Report View (Fake Report-style)"] = merged_view

            # If "New Client Demographics" exists, expose it prominently (often ~60 clients)
            if "New Client Demographics" in self.workbook_sheets:
                self.views["New Client Demographics (combined tab)"] = self.workbook_sheets["New Client Demographics"]

            # Pick a default view
            if "New Client Demographics (combined tab)" in self.views:
                self.current_view_name = "New Client Demographics (combined tab)"
            elif "Merged Report View (Fake Report-style)" in self.views:
                self.current_view_name = "Merged Report View (Fake Report-style)"
            else:
                self.current_view_name = list(self.views.keys())[0] if self.views else None

            # Update GUI
            self.root.after(0, self.populate_view_selector)
            self.root.after(0, lambda: self.set_view(self.current_view_name))
            self.root.after(0, self.enable_grant_filters)
            self.log("Workbook ready. Use grant filters to drill down.")

        except Exception as e:
            self.log(f"Error: {str(e)}")

    def build_merged_report_view(self, xl):
        """
        Build a merged dataset that resembles the "Fake Report" output:
        one row per client (deduped) with key demographic columns.
        """
        self.log("Creating merged report view (deduplicate to 1 row/client)...")

        # Prefer the already-combined tab when present, but normalize it
        if "New Client Demographics" in self.workbook_sheets:
            df = self.workbook_sheets["New Client Demographics"].copy()
            df = self.ensure_id_column(df)
            df = self.apply_light_normalizations(df)
            df = self.dedupe_clients(df)
            return df

        # Otherwise: merge across known sheets (demo version of the Guide process)
        sheet_mappings = {
            "Race - Rows": [self.id_column, "Race/Ethnicity", "Program", "Funder", "Type of Victimization"],
            "Gender - Rows": [self.id_column, "Gender", "Funder", "Type of Victimization"],
            "Age of Victim - Rows": [self.id_column, "Date of Birth", "Funder", "Type of Victimization"],
            "Disability, Veteran, LG - Rows": [
                self.id_column,
                "Veteran Status",
                "LGBTQ/Two-Spirited",
                "Disability",
                "Country of Citizenship",
                "Primary Language",
                "Immigrant Status",
            ],
            "Victimization Type - Rows": [
                self.id_column,
                "Type of Victimization",
                "Victim Type",
                "Age at Time of Trafficking",
                "Funder",
            ],
            "Homelessness - Rows": [self.id_column, "Homelessness", "Funder", "Type of Victimization"],
        }

        merged_df = None
        loaded_any = False
        for sheet, cols in sheet_mappings.items():
            if sheet not in self.workbook_sheets:
                continue
            df = self.workbook_sheets[sheet].copy()
            df = self.ensure_id_column(df)
            keep_cols = [c for c in cols if c in df.columns]
            df = df[keep_cols]
            df = df.drop_duplicates(subset=[self.id_column])
            loaded_any = True
            if merged_df is None:
                merged_df = df
            else:
                merged_df = pd.merge(merged_df, df, on=self.id_column, how="outer")

        if not loaded_any or merged_df is None:
            return None

        merged_df = self.apply_light_normalizations(merged_df)
        merged_df = self.dedupe_clients(merged_df)
        return merged_df

    def ensure_id_column(self, df):
        if self.id_column in df.columns:
            return df
        # best-effort: if there is any column that looks like legacy client id
        for c in df.columns:
            if str(c).strip().lower() in {"legacy client id", "legacy_client_id", "legacy id"}:
                df = df.rename(columns={c: self.id_column})
                return df
        return df

    def apply_light_normalizations(self, df):
        df = df.copy()

        if "Type of Victimization" in df.columns:
            self.log("Normalizing 'Type of Victimization' (comma/ampersand variants)...")
            df["Type of Victimization"] = df["Type of Victimization"].apply(self.normalize_victimization)
            time.sleep(0.2)

        if "Country of Citizenship" in df.columns:
            self.log("Standardizing 'Country of Citizenship' (typos)...")
            df["Country of Citizenship"] = df["Country of Citizenship"].apply(self.normalize_citizenship)
            time.sleep(0.2)

        # For demo: keep raw “jumbled” exports but normalize blanks to ""
        for col in ["Race/Ethnicity", "Disability", "Victim Type", "Homelessness", "Age at Time of Trafficking"]:
            if col in df.columns:
                df[col] = df[col].apply(lambda v: "" if v is None or str(v).strip().lower() == "nan" else str(v).strip())
        return df

    def dedupe_clients(self, df):
        if self.id_column in df.columns:
            before_rows = len(df)
            before_ids = df[self.id_column].nunique(dropna=True)
            df = df.drop_duplicates(subset=[self.id_column])
            after_rows = len(df)
            after_ids = df[self.id_column].nunique(dropna=True)
            self.log(f"Deduped clients: rows {before_rows} → {after_rows}, unique IDs {before_ids} → {after_ids}")
            return df
        # If no ID column, do a full-row dedupe
        before_rows = len(df)
        df = df.drop_duplicates()
        self.log(f"Deduped rows (no ID column): {before_rows} → {len(df)}")
        return df

    def normalize_victimization(self, value):
        val_str = str(value).lower()
        if 'sex trafficking' in val_str and 'labor trafficking' in val_str:
            return 'Both Sex & Labor Trafficking'
        elif 'sex trafficking' in val_str:
            return 'Sex Trafficking'
        elif 'labor trafficking' in val_str:
            return 'Labor Trafficking'
        if value is None or str(value).strip() == "" or str(value).strip().lower() == "nan":
            return ""
        # preserve as-is but bucket it
        return 'Other/Exploitation'

    def normalize_citizenship(self, value):
        val_str = str(value).strip()
        if val_str.lower() in ['nicaragua', 'nicaraugua', 'niceragua', 'nicaragua ', 'nicaragua.']:
            return 'Nicaragua'
        if val_str.lower() == "nicaragua":
            return "Nicaragua"
        return val_str

    def populate_view_selector(self):
        if not self.views:
            self.combo_view["values"] = []
            return
        names = list(self.views.keys())
        # Prefer putting key views at top if present
        preferred = [
            "New Client Demographics (combined tab)",
            "Merged Report View (Fake Report-style)",
        ]
        ordered = [n for n in preferred if n in names] + [n for n in names if n not in preferred]
        self.combo_view["values"] = ordered
        if self.current_view_name in ordered:
            self.combo_view.set(self.current_view_name)
        else:
            self.combo_view.current(0)

    def on_view_selected(self, event=None):
        name = self.combo_view.get()
        self.set_view(name)

    def enable_grant_filters(self):
        for _field, combo in self.filter_combos.items():
            combo.configure(state="readonly")
        self.btn_clear_grant_filters.config(state="normal")

    def refresh_filters_for_view(self):
        """Populate each grant filter dropdown with values from the current view."""
        if self.df_current is None:
            return

        for field in self.filter_fields:
            combo = self.filter_combos[field]
            var = self.filter_vars[field]

            if field not in self.df_current.columns:
                combo.configure(state="disabled")
                combo["values"] = ["(Not available)"]
                var.set("(Not available)")
                continue

            combo.configure(state="readonly")
            series = self.df_current[field].dropna().astype(str)
            if series.empty:
                combo["values"] = ["All"]
                var.set("All")
                continue

            # For multi-select fields, split on separators so each choice is clean/consistent
            if field in self.multi_select_fields:
                tokens = []
                for v in series:
                    parts = re.split(r"[|,&]", v)
                    for p in parts:
                        p = p.strip()
                        if p:
                            tokens.append(p)
                if tokens:
                    vals = pd.Series(tokens)
                else:
                    vals = series
            else:
                vals = series

            top_vals = vals.value_counts().head(100).index.tolist()
            combo["values"] = ["All"] + top_vals
            var.set("All")

        self.lbl_active_filters.config(text="Active filters: (none)")

    def clear_grant_filters(self):
        for field in self.filter_fields:
            if field in self.filter_vars:
                if self.filter_vars[field].get().startswith("("):
                    continue
                self.filter_vars[field].set("All")
        self.apply_grant_filters()

    def apply_grant_filters(self, event=None):
        if self.df_current is None:
            return

        df = self.df_current.copy()
        active = []

        for field in self.filter_fields:
            if field not in df.columns:
                continue
            selected = self.filter_vars[field].get()
            if not selected or selected == "All" or selected.startswith("("):
                continue
            selected_str = str(selected)
            col_series = df[field].astype(str)

            if field in self.multi_select_fields:
                # Match if any token in the jumbled text equals the selected value
                mask = col_series.apply(
                    lambda v: selected_str in [p.strip() for p in re.split(r"[|,&]", v) if p.strip()]
                )
                df = df[mask]
            else:
                df = df[col_series == selected_str]
            active.append(f"{field}={selected}")

        self.df_display = df
        self.refresh_table()
        self.update_counts_footer()
        self.lbl_active_filters.config(text="Active filters: " + (", ".join(active) if active else "(none)"))

    def update_counts_footer(self):
        if self.df_current is None or self.df_display is None:
            self.lbl_counts.config(text="Filtered clients: 0 of 0")
            self.lbl_rows.config(text="Rows: 0 of 0")
            return

        total_rows = len(self.df_current)
        shown_rows = len(self.df_display)

        total_clients = (
            self.df_current[self.id_column].nunique(dropna=True)
            if self.id_column in self.df_current.columns
            else total_rows
        )
        shown_clients = (
            self.df_display[self.id_column].nunique(dropna=True)
            if self.id_column in self.df_display.columns
            else shown_rows
        )

        self.lbl_counts.config(text=f"Filtered clients: {shown_clients} of {total_clients}")
        if shown_rows > self.max_preview_rows:
            self.lbl_rows.config(text=f"Rows: {shown_rows} of {total_rows} (preview shows first {self.max_preview_rows})")
        else:
            self.lbl_rows.config(text=f"Rows: {shown_rows} of {total_rows}")

    def set_view(self, view_name):
        if not view_name or view_name not in self.views:
            return
        self.current_view_name = view_name
        self.df_current = self.views[view_name].copy()
        self.df_current = self.ensure_id_column(self.df_current)
        self.df_current = self.apply_light_normalizations(self.df_current)
        self.df_current = self.dedupe_clients(self.df_current)

        self.df_display = self.df_current.copy()

        self.refresh_filters_for_view()
        self.apply_grant_filters()
        self.autosize_treeview_columns()

    # (Removed: generic “Column/Value” stackable filters and long overview text.
    # This demo uses grant-focused dropdown filters and an Excel-like preview grid.)

    def refresh_table(self):
        if self.df_display is None:
            return

        # Clear existing
        for i in self.tree.get_children():
            self.tree.delete(i)
            
        # Set columns (grant-style)
        cols = [c for c in self.preview_columns if c in self.df_display.columns]
        if not cols:
            cols = list(self.df_display.columns)
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140, anchor="w")

        # Add rows
        df_to_show = self.df_display.head(self.max_preview_rows)
        for _index, row in df_to_show.iterrows():
            self.tree.insert("", "end", values=[row.get(c, "") for c in cols])

    def autosize_treeview_columns(self):
        """Best-effort column sizing so the grid looks like an Excel export (readable)."""
        if self.df_display is None:
            return
        cols = self.tree["columns"]
        if not cols:
            return
        sample = self.df_display.head(40)
        for c in cols:
            max_len = max(10, len(str(c)))
            if c in sample.columns:
                for v in sample[c].astype(str).tolist():
                    max_len = max(max_len, min(len(v), 55))
            px = min(420, max(120, max_len * 7))
            self.tree.column(c, width=px)

if __name__ == "__main__":
    root = tk.Tk()
    app = GrantReportDemoApp(root)
    root.mainloop()

