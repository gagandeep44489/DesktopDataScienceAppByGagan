import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

# ---------- Helper functions ----------
def read_csv(path, delimiter=',', encoding='utf-8'):
    try:
        df = pd.read_csv(path, delimiter=delimiter, encoding=encoding)
        return df
    except Exception as e:
        raise

def preview_dataframe(df, max_rows=25):
    # Return a list of columns and rows suitable for treeview insertion
    cols = list(df.columns)
    rows = df.head(max_rows).fillna('').astype(str).values.tolist()
    return cols, rows

# ---------- GUI App ----------
class CSVCleanerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSV Data Cleaner & Formatter")
        self.geometry("1000x650")
        self.minsize(900, 500)

        self.df = None
        self.filepath = None

        self._build_ui()

    def _build_ui(self):
        # Top frame: file controls and delimiter
        top_frame = ttk.Frame(self, padding=(10, 8))
        top_frame.pack(fill='x')

        ttk.Button(top_frame, text="Load CSV", command=self.load_csv).pack(side='left')
        ttk.Button(top_frame, text="Reload", command=self.reload_csv).pack(side='left', padx=(6,0))
        ttk.Button(top_frame, text="Save CSV", command=self.save_csv).pack(side='left', padx=(6,0))

        ttk.Label(top_frame, text="Delimiter:").pack(side='left', padx=(12,2))
        self.delim_var = tk.StringVar(value=',')
        ttk.Entry(top_frame, width=4, textvariable=self.delim_var).pack(side='left')

        ttk.Label(top_frame, text="Encoding:").pack(side='left', padx=(12,2))
        self.enc_var = tk.StringVar(value='utf-8')
        ttk.Entry(top_frame, width=12, textvariable=self.enc_var).pack(side='left')

        # Middle frame: operations (left) and preview (right)
        middle = ttk.Frame(self)
        middle.pack(fill='both', expand=True, padx=10, pady=8)

        ops_frame = ttk.LabelFrame(middle, text="Cleaning Operations", width=320, padding=(10,10))
        ops_frame.pack(side='left', fill='y')

        # Strip whitespace
        ttk.Label(ops_frame, text="Trim whitespace:").pack(anchor='w')
        self.trim_cols = tk.StringVar()
        ttk.Entry(ops_frame, textvariable=self.trim_cols).pack(fill='x', pady=(0,6))
        ttk.Label(ops_frame, text="(comma-separated columns; leave blank for all)").pack(anchor='w')

        ttk.Button(ops_frame, text="Apply Trim", command=self.apply_trim).pack(fill='x', pady=(6,8))

        # Change case
        ttk.Label(ops_frame, text="Change case:").pack(anchor='w', pady=(8,0))
        self.case_choice = tk.StringVar(value='none')
        case_frame = ttk.Frame(ops_frame)
        case_frame.pack(fill='x')
        ttk.Radiobutton(case_frame, text="None", variable=self.case_choice, value='none').pack(side='left')
        ttk.Radiobutton(case_frame, text="lower", variable=self.case_choice, value='lower').pack(side='left', padx=6)
        ttk.Radiobutton(case_frame, text="UPPER", variable=self.case_choice, value='upper').pack(side='left', padx=6)
        ttk.Radiobutton(case_frame, text="Title", variable=self.case_choice, value='title').pack(side='left', padx=6)
        self.case_cols = tk.StringVar()
        ttk.Entry(ops_frame, textvariable=self.case_cols).pack(fill='x', pady=(6,4))
        ttk.Label(ops_frame, text="(columns for case change; leave blank for all)").pack(anchor='w')
        ttk.Button(ops_frame, text="Apply Case", command=self.apply_case).pack(fill='x', pady=(6,8))

        # Remove duplicates
        ttk.Label(ops_frame, text="Remove duplicates by columns:").pack(anchor='w', pady=(8,0))
        self.dup_cols = tk.StringVar()
        ttk.Entry(ops_frame, textvariable=self.dup_cols).pack(fill='x', pady=(0,6))
        ttk.Button(ops_frame, text="Remove Duplicates", command=self.remove_duplicates).pack(fill='x', pady=(6,8))

        # Drop columns
        ttk.Label(ops_frame, text="Drop columns (comma-separated):").pack(anchor='w', pady=(8,0))
        self.drop_cols = tk.StringVar()
        ttk.Entry(ops_frame, textvariable=self.drop_cols).pack(fill='x', pady=(0,6))
        ttk.Button(ops_frame, text="Drop Columns", command=self.drop_columns).pack(fill='x', pady=(6,8))

        # Fill missing
        ttk.Label(ops_frame, text="Fill missing values:").pack(anchor='w', pady=(8,0))
        fill_frame = ttk.Frame(ops_frame)
        fill_frame.pack(fill='x')
        self.fill_value = tk.StringVar()
        ttk.Entry(fill_frame, textvariable=self.fill_value).pack(side='left', fill='x', expand=True)
        ttk.Button(fill_frame, text="Fill All", command=self.fill_missing).pack(side='left', padx=6)

        # Quick stats
        ttk.Separator(ops_frame).pack(fill='x', pady=8)
        ttk.Button(ops_frame, text="Show Basic Stats", command=self.show_stats).pack(fill='x', pady=(4,8))

        # Reset / Undo: keep original copy
        ttk.Button(ops_frame, text="Reset to Original", command=self.reset_to_original).pack(fill='x', pady=(4,8))

        # Right: preview area
        preview_frame = ttk.Frame(middle)
        preview_frame.pack(side='left', fill='both', expand=True, padx=(12,0))

        self.preview_label = ttk.Label(preview_frame, text="No CSV loaded", font=('Segoe UI', 11))
        self.preview_label.pack(anchor='w')

        # Treeview for table preview
        self.tree_frame = ttk.Frame(preview_frame)
        self.tree_frame.pack(fill='both', expand=True, pady=(6,0))

        self.tree = None  # will be a ttk.Treeview

        # Bottom: status bar / log
        bottom = ttk.Frame(self, padding=(10,6))
        bottom.pack(fill='x')
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(bottom, textvariable=self.status_var).pack(side='left')

    # ---------- File operations ----------
    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files","*.csv"),("All files","*.*")])
        if not path:
            return
        try:
            delim = self.delim_var.get() or ','
            enc = self.enc_var.get() or 'utf-8'
            df = read_csv(path, delimiter=delim, encoding=enc)
            self.df_original = df.copy()
            self.df = df
            self.filepath = path
            self.status_var.set(f"Loaded: {os.path.basename(path)} ({len(self.df)} rows, {len(self.df.columns)} cols)")
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Error loading CSV", str(e))
            self.status_var.set("Failed to load CSV")

    def reload_csv(self):
        if not self.filepath:
            messagebox.showinfo("No file", "Load a CSV first.")
            return
        try:
            delim = self.delim_var.get() or ','
            enc = self.enc_var.get() or 'utf-8'
            df = read_csv(self.filepath, delimiter=delim, encoding=enc)
            self.df_original = df.copy()
            self.df = df
            self.status_var.set(f"Reloaded: {os.path.basename(self.filepath)}")
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Error reloading CSV", str(e))
            self.status_var.set("Failed to reload CSV")

    def save_csv(self):
        if self.df is None:
            messagebox.showinfo("No data", "No dataframe to save.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV files","*.csv"),("All files","*.*")])
        if not path:
            return
        try:
            delim = self.delim_var.get() or ','
            enc = self.enc_var.get() or 'utf-8'
            self.df.to_csv(path, index=False, sep=delim, encoding=enc)
            self.status_var.set(f"Saved: {os.path.basename(path)}")
            messagebox.showinfo("Saved", f"Saved cleaned CSV to:\n{path}")
        except Exception as e:
            messagebox.showerror("Save error", str(e))
            self.status_var.set("Failed to save CSV")

    # ---------- Preview ----------
    def clear_tree(self):
        if self.tree:
            self.tree.destroy()
            self.tree = None

    def update_preview(self):
        if self.df is None:
            self.preview_label.config(text="No CSV loaded")
            self.clear_tree()
            return

        cols, rows = preview_dataframe(self.df, max_rows=50)
        self.preview_label.config(text=f"Preview — {len(self.df)} rows, {len(cols)} columns")

        # Recreate tree
        self.clear_tree()
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show='headings', selectmode='extended')
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        for c in cols:
            # limit heading length for readability in GUI
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120, anchor='w')

        for r in rows:
            # convert any cell to string (already done upstream)
            self.tree.insert('', 'end', values=r)

    # ---------- Cleaning operations ----------
    def _parse_cols_input(self, text):
        if not text or str(text).strip() == '':
            return None
        items = [c.strip() for c in str(text).split(',') if c.strip()!='']
        return items if items else None

    def apply_trim(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a CSV first.")
            return
        cols = self._parse_cols_input(self.trim_cols.get())
        try:
            if cols is None:
                # apply to all object/string columns
                for c in self.df.select_dtypes(include=['object','string']).columns:
                    self.df[c] = self.df[c].astype(str).str.strip()
            else:
                for c in cols:
                    if c in self.df.columns:
                        self.df[c] = self.df[c].astype(str).str.strip()
            self.status_var.set("Applied trim whitespace")
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Trim error", str(e))

    def apply_case(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a CSV first.")
            return
        choice = self.case_choice.get()
        if choice == 'none':
            messagebox.showinfo("No-op", "Case change set to 'None'.")
            return
        cols = self._parse_cols_input(self.case_cols.get())
        try:
            target_cols = cols if cols is not None else list(self.df.select_dtypes(include=['object','string']).columns)
            for c in target_cols:
                if c in self.df.columns:
                    if choice == 'lower':
                        self.df[c] = self.df[c].astype(str).str.lower()
                    elif choice == 'upper':
                        self.df[c] = self.df[c].astype(str).str.upper()
                    elif choice == 'title':
                        self.df[c] = self.df[c].astype(str).str.title()
            self.status_var.set(f"Applied case: {choice}")
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Case change error", str(e))

    def remove_duplicates(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a CSV first.")
            return
        cols = self._parse_cols_input(self.dup_cols.get())
        try:
            before = len(self.df)
            if cols is None:
                self.df = self.df.drop_duplicates()
            else:
                valid = [c for c in cols if c in self.df.columns]
                if not valid:
                    messagebox.showinfo("Invalid columns", "No valid columns found to deduplicate.")
                    return
                self.df = self.df.drop_duplicates(subset=valid)
            after = len(self.df)
            self.status_var.set(f"Removed duplicates: {before-after} rows removed")
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Deduplicate error", str(e))

    def drop_columns(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a CSV first.")
            return
        cols = self._parse_cols_input(self.drop_cols.get())
        if not cols:
            messagebox.showinfo("No columns", "Specify columns to drop (comma-separated).")
            return
        valid = [c for c in cols if c in self.df.columns]
        if not valid:
            messagebox.showinfo("No valid columns", "None of the specified columns exist in the data.")
            return
        try:
            self.df = self.df.drop(columns=valid)
            self.status_var.set(f"Dropped columns: {', '.join(valid)}")
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Drop columns error", str(e))

    def fill_missing(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a CSV first.")
            return
        val = self.fill_value.get()
        if val is None:
            messagebox.showinfo("No value", "Enter a value to fill missing cells.")
            return
        try:
            self.df = self.df.fillna(val)
            self.status_var.set(f"Filled missing values with: {val}")
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Fill error", str(e))

    def show_stats(self):
        if self.df is None:
            messagebox.showinfo("No data", "Load a CSV first.")
            return
        numeric_count = len(self.df.select_dtypes(include=['number']).columns)
        total_rows = len(self.df)
        total_cols = len(self.df.columns)
        missing = int(self.df.isnull().sum().sum())
        info = (f"Rows: {total_rows}\nColumns: {total_cols}\nNumeric columns: {numeric_count}\nTotal missing values: {missing}")
        messagebox.showinfo("Basic Stats", info)
        self.status_var.set("Displayed basic stats")

    def reset_to_original(self):
        if not hasattr(self, 'df_original') or self.df_original is None:
            messagebox.showinfo("No original", "No original data available to reset.")
            return
        self.df = self.df_original.copy()
        self.status_var.set("Reset to original loaded data")
        self.update_preview()

if __name__ == "__main__":
    app = CSVCleanerApp()
    app.mainloop()
