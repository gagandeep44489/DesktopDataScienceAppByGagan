import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
from sqlalchemy import create_engine, Integer, Float, String, DateTime, Boolean, Text
from sqlalchemy.types import VARCHAR
import os
import math
import re
from datetime import datetime

# ---------- Utilities ----------
def infer_sqlalchemy_type(series: pd.Series):
    """
    Infer a SQLAlchemy column type for a pandas Series.
    Returns a SQLAlchemy type class (not instance).
    """
    # Drop NA for type checks
    s = series.dropna()
    if s.empty:
        # Default to TEXT if no data
        return Text

    # If dtype already numeric/integer/float/bool/datetime -> map
    if pd.api.types.is_integer_dtype(s):
        return Integer
    if pd.api.types.is_float_dtype(s):
        return Float
    if pd.api.types.is_bool_dtype(s):
        return Boolean
    if pd.api.types.is_datetime64_any_dtype(s):
        return DateTime

    # Heuristic checks on string content
    sample = s.astype(str).head(100).values
    # Check for datelike strings (ISO, common formats)
    date_like = 0
    int_like = 0
    float_like = 0
    bool_like = 0
    total = len(sample)

    date_re = re.compile(r'^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}:\d{2})?$')
    int_re = re.compile(r'^[-+]?\d+$')
    float_re = re.compile(r'^[-+]?\d*\.\d+$')
    bool_values = set(['true','false','yes','no','y','n','0','1'])

    for v in sample:
        vs = str(v).strip()
        l = vs.lower()
        if vs == '' or vs.lower() in ('nan','none','null'):
            total -= 1
            continue
        if date_re.match(vs):
            date_like += 1
        if int_re.match(vs):
            int_like += 1
        if float_re.match(vs):
            float_like += 1
        if l in bool_values:
            bool_like += 1

    # Decide type by counts
    # Prefer DateTime if many matches
    if date_like > 0 and date_like / max(1,len(sample)) > 0.6:
        return DateTime
    if int_like > 0 and int_like / max(1,len(sample)) > 0.8 and float_like == 0:
        return Integer
    if (int_like + float_like) / max(1,len(sample)) > 0.6:
        return Float
    if bool_like / max(1,len(sample)) > 0.6:
        return Boolean

    # Fallback to a bounded VARCHAR length based on max string len, otherwise Text
    max_len = s.astype(str).map(len).max()
    if max_len is None:
        return Text
    if max_len <= 255:
        return VARCHAR(int(max_len))
    return Text

def map_dtype_dict(df: pd.DataFrame):
    """Return dict column_name -> SQLAlchemy type for use in pandas.to_sql dtype parameter."""
    dtype_map = {}
    for col in df.columns:
        dtype_map[col] = infer_sqlalchemy_type(df[col])
    return dtype_map

def safe_table_name(name: str):
    """Simple safe table name (remove spaces/special chars)."""
    return re.sub(r'\W+', '_', name).strip('_')

# ---------- GUI App ----------
class ExcelToDBApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel → Database Importer (Auto Type Detection)")
        self.geometry("1100x700")
        self.minsize(900, 600)

        self.filepath = None
        self.workbook = None  # pandas ExcelFile
        self.sheet_names = []
        self.df_preview = None
        self.selected_sheet = None
        self.engine = None

        self._build_ui()

    def _build_ui(self):
        # Top toolbar
        toolbar = ttk.Frame(self, padding=8)
        toolbar.pack(fill='x')

        ttk.Button(toolbar, text="Open Excel File", command=self.open_excel).pack(side='left')
        ttk.Button(toolbar, text="Connect DB...", command=self.connect_db).pack(side='left', padx=(6,0))
        ttk.Button(toolbar, text="Import Sheet → DB", command=self.import_sheet).pack(side='left', padx=(6,0))

        # Connection display
        self.conn_label = ttk.Label(toolbar, text="DB: Not connected", foreground='blue')
        self.conn_label.pack(side='right')

        # Main paned: left operations, right preview + mapping
        main_pane = ttk.Panedwindow(self, orient='horizontal')
        main_pane.pack(fill='both', expand=True, padx=8, pady=6)

        left_frame = ttk.Frame(main_pane, width=320)
        right_frame = ttk.Frame(main_pane)
        main_pane.add(left_frame, weight=1)
        main_pane.add(right_frame, weight=3)

        # Left: file + sheet list + options
        lf = left_frame
        ttk.Label(lf, text="Excel File:").pack(anchor='w')
        self.file_label = ttk.Label(lf, text="No file loaded")
        self.file_label.pack(anchor='w', pady=(0,8))

        ttk.Label(lf, text="Sheets:").pack(anchor='w')
        self.sheets_list = tk.Listbox(lf, height=8)
        self.sheets_list.pack(fill='x', pady=(0,6))
        self.sheets_list.bind('<<ListboxSelect>>', self.on_sheet_select)

        ttk.Button(lf, text="Reload File", command=self.reload_excel).pack(fill='x', pady=(4,8))

        # Import options
        ops = ttk.LabelFrame(lf, text="Import Options", padding=8)
        ops.pack(fill='x', pady=(8,8))

        ttk.Label(ops, text="Target table name:").pack(anchor='w')
        self.table_name_var = tk.StringVar()
        ttk.Entry(ops, textvariable=self.table_name_var).pack(fill='x', pady=(2,6))

        ttk.Label(ops, text="If table exists:").pack(anchor='w')
        self.if_exists = tk.StringVar(value='fail')
        for val, txt in [('fail','Fail'), ('replace','Replace'), ('append','Append')]:
            ttk.Radiobutton(ops, text=txt, variable=self.if_exists, value=val).pack(anchor='w')

        ttk.Label(ops, text="Rows preview limit:").pack(anchor='w', pady=(6,0))
        self.preview_limit = tk.IntVar(value=200)
        ttk.Entry(ops, textvariable=self.preview_limit).pack(fill='x')

        ttk.Label(ops, text="Batch size (for inserts):").pack(anchor='w', pady=(6,0))
        self.batch_size = tk.IntVar(value=1000)
        ttk.Entry(ops, textvariable=self.batch_size).pack(fill='x')

        ttk.Button(lf, text="Auto-detect column types", command=self.autodetect_types).pack(fill='x', pady=(8,4))
        ttk.Button(lf, text="Open mapping editor", command=self.open_mapping_editor).pack(fill='x')

        # Right: preview and mapping table
        rf = right_frame
        top_rf = ttk.Frame(rf)
        top_rf.pack(fill='x')

        self.preview_title = ttk.Label(top_rf, text="Preview: no sheet selected", font=('Segoe UI', 11))
        self.preview_title.pack(anchor='w')

        # Treeview preview
        tree_frame = ttk.Frame(rf)
        tree_frame.pack(fill='both', expand=True, pady=(6,0))

        self.tree_container = tree_frame
        self.tree = None

        # Bottom: detected schema & actions
        bottom_rf = ttk.Frame(rf, padding=(6,6))
        bottom_rf.pack(fill='x')

        self.schema_text = tk.Text(bottom_rf, height=6, wrap='none')
        self.schema_text.pack(fill='x', expand=False)
        self.schema_text.insert('1.0', 'Detected schema will appear here after autodetection.\n')
        self.schema_text.configure(state='disabled')

        action_frame = ttk.Frame(bottom_rf)
        action_frame.pack(fill='x', pady=(6,0))
        ttk.Button(action_frame, text="Preview SQL DDL", command=self.preview_ddl).pack(side='left')
        ttk.Button(action_frame, text="Validate & Clean Types", command=self.validate_and_clean).pack(side='left', padx=(6,0))
        ttk.Button(action_frame, text="Clear Schema", command=self.clear_schema).pack(side='left', padx=(6,0))

        # status bar
        status = ttk.Frame(self, padding=(6,6))
        status.pack(fill='x')
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status, textvariable=self.status_var).pack(side='left')

    # ---------- Excel file handling ----------
    def open_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx;*.xls"), ("All files","*.*")])
        if not path:
            return
        try:
            self.filepath = path
            self.workbook = pd.ExcelFile(path, engine=None)  # let pandas auto-detect engine
            self.sheet_names = self.workbook.sheet_names
            self.file_label.config(text=os.path.basename(path) + f" — {len(self.sheet_names)} sheets")
            self.sheets_list.delete(0, tk.END)
            for s in self.sheet_names:
                self.sheets_list.insert(tk.END, s)
            self.status_var.set(f"Loaded workbook: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Open Excel", str(e))
            self.status_var.set("Failed to open workbook")

    def reload_excel(self):
        if not self.filepath:
            messagebox.showinfo("No file", "Open an Excel file first.")
            return
        try:
            self.workbook = pd.ExcelFile(self.filepath, engine=None)
            self.sheet_names = self.workbook.sheet_names
            self.sheets_list.delete(0, tk.END)
            for s in self.sheet_names:
                self.sheets_list.insert(tk.END, s)
            self.status_var.set("Reloaded workbook")
        except Exception as e:
            messagebox.showerror("Reload Excel", str(e))

    def on_sheet_select(self, event=None):
        sel = self.sheets_list.curselection()
        if not sel:
            return
        name = self.sheets_list.get(sel[0])
        self.selected_sheet = name
        try:
            n = max(1, int(self.preview_limit.get()))
        except Exception:
            n = 200
        try:
            df = pd.read_excel(self.filepath, sheet_name=name, nrows=n, engine=None)
            self.df_preview = df
            self.update_preview_tree()
            self.preview_title.config(text=f"Preview — Sheet: {name} ({len(df)} rows shown)")
            self.status_var.set(f"Loaded sheet preview: {name}")
            # clear any previous schema display
            self.clear_schema()
        except Exception as e:
            messagebox.showerror("Read sheet", str(e))
            self.status_var.set("Failed to read sheet")

    def update_preview_tree(self):
        # Clear previous tree
        if self.tree:
            self.tree.destroy()
            self.tree = None
        if self.df_preview is None:
            return
        cols = list(self.df_preview.columns)
        rows = self.df_preview.fillna('').astype(str).values.tolist()

        container = self.tree_container
        self.tree = ttk.Treeview(container, columns=cols, show='headings')
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=140, anchor='w')

        for r in rows:
            self.tree.insert('', 'end', values=r)

    # ---------- DB Connection ----------
    def connect_db(self):
        # Ask user for connection string
        current = getattr(self, 'conn_string', '')
        prompt = ("Enter SQLAlchemy connection string.\n"
                  "Examples:\n"
                  "  sqlite:///mydb.sqlite\n"
                  "  mysql+pymysql://user:pass@host:3306/dbname\n"
                  "  postgresql+psycopg2://user:pass@host:5432/dbname\n\n"
                  "Leave blank to use an in-memory SQLite DB for quick testing.")
        conn = simpledialog.askstring("Connect to DB", prompt, initialvalue=current, parent=self)
        if conn is None:
            return
        if conn.strip() == '':
            conn = "sqlite:///:memory:"
        try:
            engine = create_engine(conn, future=True)
            # Try a quick connect
            with engine.connect() as conn_test:
                pass
            self.engine = engine
            self.conn_string = conn
            self.conn_label.config(text=f"DB: {conn}")
            self.status_var.set("Connected to database")
        except Exception as e:
            messagebox.showerror("DB Connection Error", str(e))
            self.status_var.set("DB connection failed")

    # ---------- Auto-detect / mapping ----------
    def autodetect_types(self):
        if self.df_preview is None:
            messagebox.showinfo("No preview", "Select a sheet to generate preview first.")
            return
        # Use a larger sample by reading more rows (but not whole file unless small)
        try:
            df_full = pd.read_excel(self.filepath, sheet_name=self.selected_sheet, engine=None)
        except Exception:
            df_full = self.df_preview.copy()
        dtype_map = map_dtype_dict(df_full)
        # Display mapping in schema_text
        self.schema = dtype_map
        self._display_schema(dtype_map)
        self.status_var.set("Auto-detected column types")

    def _display_schema(self, dtype_map):
        self.schema_text.configure(state='normal')
        self.schema_text.delete('1.0', 'end')
        for col, typ in dtype_map.items():
            tname = getattr(typ, '__name__', str(typ))
            # If instance e.g. VARCHAR(50), display repr
            if isinstance(typ, str):
                display = typ
            else:
                try:
                    display = str(typ)
                except Exception:
                    display = tname
            self.schema_text.insert('end', f"{col} -> {display}\n")
        self.schema_text.configure(state='disabled')

    def open_mapping_editor(self):
        if self.df_preview is None:
            messagebox.showinfo("No preview", "Select a sheet to edit mapping.")
            return
        editor = MappingEditor(self, self.df_preview, getattr(self, 'schema', None))
        self.wait_window(editor)
        if getattr(editor, 'result', None):
            # editor.result is a dict column -> sqlalchemy type (class or instance)
            self.schema = editor.result
            self._display_schema(self.schema)
            self.status_var.set("Mapping updated")

    # ---------- Validation & cleaning ----------
    def validate_and_clean(self):
        """
        Attempt to coerce columns in the preview to the detected types where feasible.
        Shows stats about conversions and non-convertible rows.
        """
        if self.df_preview is None or not hasattr(self, 'schema'):
            messagebox.showinfo("No schema", "Run autodetect or open mapping editor first.")
            return
        df = self.df_preview.copy()
        report = []
        for col, typ in self.schema.items():
            series = df[col]
            target = typ
            convertible = True
            failed = 0
            if target in (Integer, Float):
                # attempt numeric coercion
                coerced = pd.to_numeric(series, errors='coerce')
                failed = int(coerced.isna().sum() - series.isna().sum())
                df[col] = coerced
            elif target is Boolean:
                def to_bool(v):
                    try:
                        s = str(v).strip().lower()
                        if s in ('true','1','t','y','yes'):
                            return True
                        if s in ('false','0','f','n','no'):
                            return False
                        return pd.NA
                    except:
                        return pd.NA
                coerced = series.map(to_bool)
                failed = int(coerced.isna().sum() - series.isna().sum())
                df[col] = coerced
            elif target is DateTime:
                coerced = pd.to_datetime(series, errors='coerce', infer_datetime_format=True)
                failed = int(coerced.isna().sum() - series.isna().sum())
                df[col] = coerced
            else:
                # strings: no coercion needed
                failed = 0
            report.append((col, getattr(target, '__name__', str(target)), failed))
        # Show small report
        msg = "Validation Report (non-convertible counts shown)\n\n"
        for col, tname, failed in report:
            msg += f"{col} -> {tname}: {failed} problematic rows\n"
        messagebox.showinfo("Validation Report", msg)
        self.status_var.set("Validation complete (see report)")

    def clear_schema(self):
        self.schema = {}
        self.schema_text.configure(state='normal')
        self.schema_text.delete('1.0', 'end')
        self.schema_text.insert('1.0','Detected schema will appear here after autodetection.\n')
        self.schema_text.configure(state='disabled')

    def preview_ddl(self):
        if not hasattr(self, 'schema') or not self.schema:
            messagebox.showinfo("No schema", "Autodetect or open mapping editor first.")
            return
        # Build a simple CREATE TABLE DDL preview (SQLite-ish)
        table = safe_table_name(self.table_name_var.get() or self.selected_sheet or 'import_table')
        ddl = f"CREATE TABLE {table} (\n"
        lines = []
        for col, typ in self.schema.items():
            # Map to SQL types for display
            if typ in (Integer,):
                st = "INTEGER"
            elif typ in (Float,):
                st = "REAL"
            elif typ in (DateTime,):
                st = "DATETIME"
            elif typ in (Boolean,):
                st = "BOOLEAN"
            elif isinstance(typ, VARCHAR):
                st = f"VARCHAR({typ.length})"
            else:
                st = "TEXT"
            lines.append(f"  {col} {st}")
        ddl += ",\n".join(lines) + "\n);"
        DDLWindow(self, ddl)

    # ---------- Import ----------
    def import_sheet(self):
        if self.engine is None:
            messagebox.showinfo("No DB", "Connect to a database first.")
            return
        if self.selected_sheet is None:
            messagebox.showinfo("No sheet", "Select a sheet to import.")
            return
        table = self.table_name_var.get().strip() or self.selected_sheet
        table = safe_table_name(table)
        if_exists = self.if_exists.get()
        batch = max(1, int(self.batch_size.get() or 1000))

        try:
            # Read full sheet
            df = pd.read_excel(self.filepath, sheet_name=self.selected_sheet, engine=None)
            # If schema not set, auto detect
            if not hasattr(self, 'schema') or not self.schema:
                dtype_map = map_dtype_dict(df)
            else:
                dtype_map = self.schema

            # Build dtype mapping for pandas.to_sql: must be SQLAlchemy types (instances/classes)
            # pandas accepts either SQLAlchemy types or dict of column->type
            dtype_for_pandas = {}
            for col, typ in dtype_map.items():
                # If typ is a class like Integer, pass SQLAlchemy Integer()
                try:
                    if isinstance(typ, str):
                        # ignore strings here
                        dtype_for_pandas[col] = String()
                    elif isinstance(typ, type):
                        dtype_for_pandas[col] = typ()
                    else:
                        # Already an instance e.g. VARCHAR(n)
                        dtype_for_pandas[col] = typ
                except Exception:
                    dtype_for_pandas[col] = String()

            # Attempt to coerce columns to aligned dtypes where appropriate
            # (coercion similar to validation)
            for col, sqlt in dtype_map.items():
                if col not in df.columns:
                    continue
                if sqlt in (Integer,):
                    df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')
                elif sqlt in (Float,):
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                elif sqlt is Boolean:
                    df[col] = df[col].map(lambda v: True if str(v).strip().lower() in ('true','1','t','y','yes') else (False if str(v).strip().lower() in ('false','0','f','n','no') else pd.NA)).astype('boolean')
                elif sqlt is DateTime:
                    df[col] = pd.to_datetime(df[col], errors='coerce', infer_datetime_format=True)
                else:
                    df[col] = df[col].astype(str)

            # Use pandas to_sql with dtype and method for batch inserts
            # method='multi' helps with many DBs, but for very large files consider chunking
            chunksize = batch
            with self.engine.begin() as conn:
                df.to_sql(name=table, con=conn, if_exists=if_exists, index=False, dtype=dtype_for_pandas, chunksize=chunksize, method='multi')
            messagebox.showinfo("Import complete", f"Imported sheet '{self.selected_sheet}' into table '{table}'.")
            self.status_var.set("Import successful")
        except Exception as e:
            messagebox.showerror("Import error", str(e))
            self.status_var.set("Import failed")

# ---------- Helper windows ----------
class MappingEditor(tk.Toplevel):
    def __init__(self, parent, df_sample: pd.DataFrame, existing_schema: dict = None):
        super().__init__(parent)
        self.title("Column Mapping Editor")
        self.geometry("800x500")
        self.parent = parent
        self.df = df_sample
        self.existing = existing_schema or {}
        self.result = None

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill='both', expand=True)

        ttk.Label(top, text="Column").grid(row=0, column=0, sticky='w')
        ttk.Label(top, text="Detected Type").grid(row=0, column=1, sticky='w')
        ttk.Label(top, text="Override").grid(row=0, column=2, sticky='w')
        ttk.Label(top, text="Sample Value").grid(row=0, column=3, sticky='w')

        self.rows = []
        available_types = ['Auto-Detect', 'Integer', 'Float', 'Boolean', 'DateTime', 'String', 'Text', 'VARCHAR(50)']
        for i, col in enumerate(self.df.columns):
            ttk.Label(top, text=col).grid(row=i+1, column=0, sticky='w', padx=4, pady=2)
            detected = infer_sqlalchemy_type(self.df[col])
            det_name = getattr(detected, '__name__', str(detected))
            ttk.Label(top, text=det_name).grid(row=i+1, column=1, sticky='w', padx=4)

            var = tk.StringVar(value='Auto-Detect')
            cb = ttk.Combobox(top, values=available_types, textvariable=var, width=18)
            cb.grid(row=i+1, column=2, padx=4, pady=2)
            sample = str(self.df[col].dropna().astype(str).head(3).tolist())
            ttk.Label(top, text=sample).grid(row=i+1, column=3, sticky='w', padx=4)
            self.rows.append((col, var))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill='x', pady=8)
        ttk.Button(btn_frame, text="Apply and Close", command=self.apply).pack(side='right', padx=6)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side='right')

    def apply(self):
        mapping = {}
        for col, var in self.rows:
            choice = var.get()
            if choice == 'Auto-Detect':
                mapping[col] = infer_sqlalchemy_type(self.df[col])
            elif choice == 'Integer':
                mapping[col] = Integer
            elif choice == 'Float':
                mapping[col] = Float
            elif choice == 'Boolean':
                mapping[col] = Boolean
            elif choice == 'DateTime':
                mapping[col] = DateTime
            elif choice.startswith('VARCHAR'):
                # parse length
                m = re.search(r'\((\d+)\)', choice)
                if m:
                    mapping[col] = VARCHAR(int(m.group(1)))
                else:
                    mapping[col] = VARCHAR(255)
            elif choice == 'Text':
                mapping[col] = Text
            else:
                mapping[col] = String
        self.result = mapping
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()

class DDLWindow(tk.Toplevel):
    def __init__(self, parent, ddl_text: str):
        super().__init__(parent)
        self.title("DDL Preview")
        self.geometry("700x400")
        txt = tk.Text(self, wrap='none')
        txt.pack(fill='both', expand=True, padx=6, pady=6)
        txt.insert('1.0', ddl_text)
        txt.configure(state='disabled')

# ---------- Main ----------
if __name__ == "__main__":
    app = ExcelToDBApp()
    app.mainloop()
