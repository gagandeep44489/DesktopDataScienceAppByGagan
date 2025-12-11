"""
Data Imputation & Outlier Detection Tool
Single-file desktop app using tkinter + pandas + scikit-learn + matplotlib.

Features:
- Load CSV / Excel (.xlsx/.xls)
- Preview dataset (first N rows)
- Column selection UI
- Imputation methods: Mean, Median, Mode, Constant, Interpolation (linear), KNN Imputer (scikit-learn)
- Outlier detection: Z-score, IQR, IsolationForest (scikit-learn)
- Visualize column distribution with detected outliers overlay
- Apply imputation and optionally remove or flag outliers
- Save cleaned dataset (CSV)
- Reset to original data

Requirements:
pip install pandas scikit-learn matplotlib openpyxl scipy
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import numpy as np
import os
import math
import matplotlib
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from sklearn.impute import KNNImputer
from sklearn.ensemble import IsolationForest
from scipy import stats

# ---------- Helper functions ----------
def read_file(path, nrows=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xls', '.xlsx'):
        return pd.read_excel(path, nrows=nrows, engine=None)
    else:
        return pd.read_csv(path, nrows=nrows)

def infer_numeric_columns(df):
    return list(df.select_dtypes(include=[np.number]).columns)

def zscore_outliers(series, thresh=3.0):
    """Return boolean mask of outliers based on z-score (abs(z) > thresh)."""
    try:
        z = np.abs(stats.zscore(series.dropna()))
        mask = pd.Series(False, index=series.index)
        mask.loc[series.dropna().index] = z > thresh
        return mask
    except Exception:
        # fallback: no outliers
        return pd.Series(False, index=series.index)

def iqr_outliers(series, k=1.5):
    """Return boolean mask of outliers based on IQR rule."""
    s = series.dropna()
    q1 = s.quantile(0.25)
    q3 = s.quantile(0.75)
    iqr = q3 - q1
    lower = q1 - k * iqr
    upper = q3 + k * iqr
    mask = (series < lower) | (series > upper)
    mask = mask.fillna(False)
    return mask

def isolation_forest_outliers(df_numeric, contamination=0.05, random_state=0):
    """Return boolean mask (for df index) where True indicates outlier. Uses isolation forest on numeric features."""
    if df_numeric.shape[0] == 0 or df_numeric.shape[1] == 0:
        return pd.Series(False, index=df_numeric.index)
    try:
        iso = IsolationForest(contamination=contamination, random_state=random_state)
        # fill NaNs temporarily with median for fitting
        X = df_numeric.copy()
        X = X.fillna(X.median())
        preds = iso.fit_predict(X)  # -1 outlier, 1 inlier
        mask = pd.Series(preds == -1, index=df_numeric.index)
        return mask
    except Exception:
        return pd.Series(False, index=df_numeric.index)

# ---------- GUI App ----------
class ImputeOutlierApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Data Imputation & Outlier Detection Tool")
        self.geometry("1100x720")
        self.minsize(1000, 650)

        self.filepath = None
        self.df_original = None
        self.df = None
        self.preview_limit = 200

        # UI state
        self.selected_columns = []
        self.impute_settings = {}  # column -> dict(method=..., value=..., knn_k=...)
        self.outlier_mask = pd.Series(dtype=bool)

        self._build_ui()

    def _build_ui(self):
        # Top toolbar
        toolbar = ttk.Frame(self, padding=8)
        toolbar.pack(fill='x')

        ttk.Button(toolbar, text="Open File", command=self.open_file).pack(side='left')
        ttk.Button(toolbar, text="Reload", command=self.reload_file).pack(side='left', padx=(6,0))
        ttk.Button(toolbar, text="Save Cleaned CSV", command=self.save_file).pack(side='left', padx=(6,0))

        ttk.Label(toolbar, text="Preview rows:").pack(side='left', padx=(20,4))
        self.preview_var = tk.IntVar(value=self.preview_limit)
        ttk.Entry(toolbar, width=6, textvariable=self.preview_var).pack(side='left')

        ttk.Label(toolbar, text="Status:").pack(side='right', padx=(6,2))
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(toolbar, textvariable=self.status_var, foreground='blue').pack(side='right')

        # Main pane
        main = ttk.Panedwindow(self, orient='horizontal')
        main.pack(fill='both', expand=True, padx=8, pady=8)

        left = ttk.Frame(main, width=340)
        right = ttk.Frame(main)
        main.add(left, weight=1)
        main.add(right, weight=3)

        # Left panel: controls
        left_scroll = ttk.Frame(left)
        left_scroll.pack(fill='both', expand=True)

        # File info
        ttk.Label(left_scroll, text="File:").pack(anchor='w')
        self.file_label = ttk.Label(left_scroll, text="No file loaded")
        self.file_label.pack(anchor='w', pady=(0,8))

        # Column list
        ttk.Label(left_scroll, text="Columns: (select for operations)").pack(anchor='w')
        self.col_listbox = tk.Listbox(left_scroll, selectmode='extended', height=12)
        self.col_listbox.pack(fill='x', pady=(2,8))
        self.col_listbox.bind('<<ListboxSelect>>', self.on_column_select)

        # Imputation options
        impute_frame = ttk.LabelFrame(left_scroll, text="Imputation", padding=8)
        impute_frame.pack(fill='x', pady=(6,8))

        ttk.Label(impute_frame, text="Method:").pack(anchor='w')
        self.impute_method = tk.StringVar(value='mean')
        methods = [('Mean','mean'), ('Median','median'), ('Mode','mode'), ('Constant','constant'),
                   ('Interpolation (linear)','interpolate'), ('KNN Imputer','knn')]
        for txt, val in methods:
            ttk.Radiobutton(impute_frame, text=txt, variable=self.impute_method, value=val).pack(anchor='w')

        ttk.Label(impute_frame, text="Constant value (for Constant):").pack(anchor='w', pady=(6,0))
        self.constant_value = tk.StringVar()
        ttk.Entry(impute_frame, textvariable=self.constant_value).pack(fill='x')

        ttk.Label(impute_frame, text="K (neighbors for KNN):").pack(anchor='w', pady=(6,0))
        self.knn_k = tk.IntVar(value=5)
        ttk.Entry(impute_frame, textvariable=self.knn_k).pack(fill='x')

        ttk.Button(impute_frame, text="Apply Imputation to Selected", command=self.apply_imputation_to_selected).pack(fill='x', pady=(8,4))
        ttk.Button(impute_frame, text="Apply Imputation to All", command=self.apply_imputation_to_all).pack(fill='x')

        # Outlier detection options
        outlier_frame = ttk.LabelFrame(left_scroll, text="Outlier Detection", padding=8)
        outlier_frame.pack(fill='x', pady=(8,8))

        ttk.Label(outlier_frame, text="Method:").pack(anchor='w')
        self.outlier_method = tk.StringVar(value='zscore')
        out_methods = [('Z-score','zscore'), ('IQR','iqr'), ('Isolation Forest (multi-col)','isoforest')]
        for txt, val in out_methods:
            ttk.Radiobutton(outlier_frame, text=txt, variable=self.outlier_method, value=val).pack(anchor='w')

        ttk.Label(outlier_frame, text="Z-score threshold (abs z):").pack(anchor='w', pady=(6,0))
        self.z_thresh = tk.DoubleVar(value=3.0)
        ttk.Entry(outlier_frame, textvariable=self.z_thresh).pack(fill='x')

        ttk.Label(outlier_frame, text="IQR multiplier (k):").pack(anchor='w', pady=(6,0))
        self.iqr_k = tk.DoubleVar(value=1.5)
        ttk.Entry(outlier_frame, textvariable=self.iqr_k).pack(fill='x')

        ttk.Label(outlier_frame, text="IsolationForest contamination (0-0.5):").pack(anchor='w', pady=(6,0))
        self.iso_cont = tk.DoubleVar(value=0.05)
        ttk.Entry(outlier_frame, textvariable=self.iso_cont).pack(fill='x')

        ttk.Button(outlier_frame, text="Detect Outliers (selected cols)", command=self.detect_outliers_selected).pack(fill='x', pady=(8,4))
        ttk.Button(outlier_frame, text="Detect Outliers (all numeric)", command=self.detect_outliers_all).pack(fill='x')

        ttk.Button(left_scroll, text="Flag Outliers Column", command=self.flag_outliers_column).pack(fill='x', pady=(12,4))
        ttk.Button(left_scroll, text="Remove Outlier Rows", command=self.remove_outlier_rows).pack(fill='x')

        ttk.Button(left_scroll, text="Reset to Original", command=self.reset_to_original).pack(fill='x', pady=(12,4))

        # Right panel: preview and visualization
        top_right = ttk.Frame(right)
        top_right.pack(fill='x')
        ttk.Label(top_right, text="Preview:", font=('Segoe UI', 11)).pack(anchor='w')
        self.preview_label = ttk.Label(top_right, text="No data loaded")
        self.preview_label.pack(anchor='w', pady=(0,6))

        # Treeview preview
        tree_container = ttk.Frame(right)
        tree_container.pack(fill='both', expand=True)
        self.tree_container = tree_container
        self.tree = None

        # Bottom-right: visualization area
        viz_frame = ttk.LabelFrame(right, text="Visualization", padding=6)
        viz_frame.pack(fill='both', expand=False, pady=(8,0))
        self.fig = Figure(figsize=(6,3))
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=viz_frame)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)

        viz_controls = ttk.Frame(viz_frame)
        viz_controls.pack(fill='x', pady=(6,0))
        ttk.Button(viz_controls, text="Plot Selected Column Distribution", command=self.plot_selected_column).pack(side='left')
        ttk.Button(viz_controls, text="Clear Plot", command=self.clear_plot).pack(side='left', padx=(6,0))

    # ---------- File handling ----------
    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV","*.csv"), ("Excel","*.xlsx;*.xls"), ("All files","*.*")])
        if not path:
            return
        try:
            limit = int(self.preview_var.get() or self.preview_limit)
        except Exception:
            limit = self.preview_limit
        try:
            df = read_file(path, nrows=limit if limit>0 else None)
            # For preview we load limited rows; but keep full dataset separately by re-reading full on demand
            full_df = read_file(path, nrows=None)
            self.filepath = path
            self.df_original = full_df.copy()
            self.df = full_df.copy()
            self.file_label.config(text=os.path.basename(path))
            self.status_var.set(f"Loaded file: {os.path.basename(path)} ({len(self.df)} rows, {len(self.df.columns)} cols)")
            self.update_columns_list()
            self.update_preview()
            self.outlier_mask = pd.Series(False, index=self.df.index)
        except Exception as e:
            messagebox.showerror("Open file error", str(e))
            self.status_var.set("Failed to open file")

    def reload_file(self):
        if not self.filepath:
            messagebox.showinfo("No file", "Open a file first.")
            return
        self.open_file()

    def save_file(self):
        if self.df is None:
            messagebox.showinfo("No data", "No data to save.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv"), ("All files","*.*")])
        if not path:
            return
        try:
            self.df.to_csv(path, index=False)
            messagebox.showinfo("Saved", f"Saved cleaned CSV to:\n{path}")
            self.status_var.set(f"Saved file: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Save error", str(e))
            self.status_var.set("Failed to save file")

    def reset_to_original(self):
        if self.df_original is None:
            return
        self.df = self.df_original.copy()
        self.update_preview()
        self.update_columns_list()
        self.outlier_mask = pd.Series(False, index=self.df.index)
        self.status_var.set("Reset to original data")

    # ---------- UI updates ----------
    def update_columns_list(self):
        self.col_listbox.delete(0, tk.END)
        if self.df is None:
            return
        for c in self.df.columns:
            self.col_listbox.insert(tk.END, c)

    def on_column_select(self, event=None):
        sel = self.col_listbox.curselection()
        self.selected_columns = [self.col_listbox.get(i) for i in sel]
        # Update status
        if self.selected_columns:
            self.status_var.set(f"Selected columns: {', '.join(self.selected_columns)}")
        else:
            self.status_var.set("No column selected")

    def update_preview(self):
        # Destroy previous tree
        if self.tree:
            self.tree.destroy()
            self.tree = None
        if self.df is None:
            self.preview_label.config(text="No data loaded")
            return
        limit = int(self.preview_var.get() or self.preview_limit)
        dfp = self.df.head(limit)
        self.preview_label.config(text=f"Preview — {len(dfp)} rows shown")
        cols = list(dfp.columns)
        rows = dfp.fillna('').astype(str).values.tolist()

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

    # ---------- Imputation ----------
    def apply_imputation_to_selected(self):
        if self.df is None:
            messagebox.showinfo("No data", "Open a file first.")
            return
        if not self.selected_columns:
            messagebox.showinfo("No columns", "Select one or more columns to impute.")
            return
        method = self.impute_method.get()
        k = max(1, int(self.knn_k.get() or 5))
        const = self.constant_value.get()
        try:
            if method == 'knn':
                # use KNN imputer on numeric subset including selected columns
                numeric_cols = infer_numeric_columns(self.df)
                if not numeric_cols:
                    messagebox.showinfo("No numeric columns", "KNN imputer requires numeric columns.")
                    return
                imputer = KNNImputer(n_neighbors=k)
                arr = imputer.fit_transform(self.df[numeric_cols])
                self.df[numeric_cols] = pd.DataFrame(arr, columns=numeric_cols, index=self.df.index)
            else:
                for col in self.selected_columns:
                    if method == 'mean':
                        if self.df[col].dtype.kind in 'biufc':
                            val = self.df[col].mean()
                        else:
                            # non-numeric: fallback to mode
                            val = self.df[col].mode().iloc[0] if not self.df[col].mode().empty else ''
                        self.df[col] = self.df[col].fillna(val)
                    elif method == 'median':
                        if self.df[col].dtype.kind in 'biufc':
                            val = self.df[col].median()
                            self.df[col] = self.df[col].fillna(val)
                        else:
                            val = self.df[col].mode().iloc[0] if not self.df[col].mode().empty else ''
                            self.df[col] = self.df[col].fillna(val)
                    elif method == 'mode':
                        val = self.df[col].mode().iloc[0] if not self.df[col].mode().empty else ''
                        self.df[col] = self.df[col].fillna(val)
                    elif method == 'constant':
                        # try to convert constant to numeric if column is numeric
                        if self.df[col].dtype.kind in 'biufc':
                            try:
                                cval = float(const)
                            except Exception:
                                cval = np.nan
                            self.df[col] = self.df[col].fillna(cval)
                        else:
                            self.df[col] = self.df[col].fillna(const)
                    elif method == 'interpolate':
                        # for numeric columns only
                        if self.df[col].dtype.kind in 'biufc':
                            self.df[col] = self.df[col].interpolate(method='linear', limit_direction='both')
                        else:
                            # non-numeric: fallback to forward-fill then back-fill
                            self.df[col] = self.df[col].fillna(method='ffill').fillna(method='bfill')
            self.update_preview()
            self.status_var.set(f"Imputation applied ({method})")
        except Exception as e:
            messagebox.showerror("Imputation error", str(e))
            self.status_var.set("Imputation failed")

    def apply_imputation_to_all(self):
        if self.df is None:
            messagebox.showinfo("No data", "Open a file first.")
            return
        all_cols = list(self.df.columns)
        self.selected_columns = all_cols
        self.apply_imputation_to_selected()

    # ---------- Outlier detection ----------
    def detect_outliers_selected(self):
        if self.df is None:
            messagebox.showinfo("No data", "Open a file first.")
            return
        if not self.selected_columns:
            messagebox.showinfo("No columns", "Select one or more columns for outlier detection.")
            return
        method = self.outlier_method.get()
        mask = pd.Series(False, index=self.df.index)
        try:
            if method == 'zscore':
                thresh = float(self.z_thresh.get() or 3.0)
                for col in self.selected_columns:
                    if col in self.df.columns and self.df[col].dtype.kind in 'biufc':
                        mask = mask | zscore_outliers(self.df[col], thresh=thresh)
            elif method == 'iqr':
                k = float(self.iqr_k.get() or 1.5)
                for col in self.selected_columns:
                    if col in self.df.columns and self.df[col].dtype.kind in 'biufc':
                        mask = mask | iqr_outliers(self.df[col], k=k)
            elif method == 'isoforest':
                # run isolation forest on numeric subset of selected columns
                num_cols = [c for c in self.selected_columns if c in self.df.columns and self.df[c].dtype.kind in 'biufc']
                if not num_cols:
                    messagebox.showinfo("No numeric columns", "IsolationForest requires numeric columns.")
                    return
                cont = float(self.iso_cont.get() or 0.05)
                mask = mask | isolation_forest_outliers(self.df[num_cols], contamination=cont)
            self.outlier_mask = mask
            n = int(mask.sum())
            messagebox.showinfo("Outlier Detection", f"Detected {n} outlier rows (flagged).")
            self.status_var.set(f"Detected {n} outliers")
            # Optionally highlight in preview? For simplicity we just show counts and allow flag/remove
        except Exception as e:
            messagebox.showerror("Outlier detection error", str(e))
            self.status_var.set("Outlier detection failed")

    def detect_outliers_all(self):
        if self.df is None:
            messagebox.showinfo("No data", "Open a file first.")
            return
        method = self.outlier_method.get()
        mask = pd.Series(False, index=self.df.index)
        try:
            numeric_cols = infer_numeric_columns(self.df)
            if method == 'zscore':
                thresh = float(self.z_thresh.get() or 3.0)
                for col in numeric_cols:
                    mask = mask | zscore_outliers(self.df[col], thresh=thresh)
            elif method == 'iqr':
                k = float(self.iqr_k.get() or 1.5)
                for col in numeric_cols:
                    mask = mask | iqr_outliers(self.df[col], k=k)
            elif method == 'isoforest':
                cont = float(self.iso_cont.get() or 0.05)
                mask = isolation_forest_outliers(self.df[numeric_cols], contamination=cont)
            self.outlier_mask = mask
            n = int(mask.sum())
            messagebox.showinfo("Outlier Detection", f"Detected {n} outlier rows (flagged).")
            self.status_var.set(f"Detected {n} outliers")
        except Exception as e:
            messagebox.showerror("Outlier detection error", str(e))
            self.status_var.set("Outlier detection failed")

    def flag_outliers_column(self):
        """Create a boolean column 'is_outlier' in dataframe to flag detected outliers."""
        if self.df is None:
            return
        if self.outlier_mask is None or self.outlier_mask.shape[0] == 0:
            messagebox.showinfo("No outliers", "No outliers detected. Run detection first.")
            return
        colname = 'is_outlier'
        # ensure unique column name
        base = colname
        i = 1
        while colname in self.df.columns:
            colname = f"{base}_{i}"
            i += 1
        self.df[colname] = self.outlier_mask.astype(bool).values
        self.update_columns_list()
        self.update_preview()
        self.status_var.set(f"Flagged outliers in column '{colname}'")

    def remove_outlier_rows(self):
        if self.df is None:
            return
        if self.outlier_mask is None or self.outlier_mask.sum() == 0:
            messagebox.showinfo("No outliers", "No outliers detected. Run detection first.")
            return
        before = len(self.df)
        self.df = self.df.loc[~self.outlier_mask].reset_index(drop=True)
        removed = before - len(self.df)
        self.outlier_mask = pd.Series(False, index=self.df.index)
        self.update_columns_list()
        self.update_preview()
        messagebox.showinfo("Removed outliers", f"Removed {removed} rows.")
        self.status_var.set(f"Removed {removed} outlier rows")

    # ---------- Visualization ----------
    def plot_selected_column(self):
        if self.df is None:
            messagebox.showinfo("No data", "Open a file first.")
            return
        if not self.selected_columns or len(self.selected_columns) > 1:
            messagebox.showinfo("Select column", "Select exactly one numeric column to plot.")
            return
        col = self.selected_columns[0]
        if col not in self.df.columns:
            return
        series = self.df[col]
        if series.dtype.kind not in 'biufc':
            messagebox.showinfo("Non-numeric", "Selected column is not numeric. Plotting categorical frequency instead.")
            # plot value counts
            self.ax.clear()
            vc = series.fillna('<<NA>>').astype(str).value_counts().head(30)
            vc.plot(kind='bar', ax=self.ax)
            self.ax.set_title(f"Value counts: {col}")
            self.canvas.draw()
            return

        # numeric distribution
        self.ax.clear()
        s = series.dropna()
        if s.empty:
            messagebox.showinfo("Empty column", "No numeric values to plot in selected column.")
            return
        # histogram
        self.ax.hist(s, bins='auto', alpha=0.7, edgecolor='black')
        # mark outliers if available
        if self.outlier_mask is not None and self.outlier_mask.any():
            out_vals = self.df.loc[self.outlier_mask, col].dropna()
            if not out_vals.empty:
                ymin, ymax = self.ax.get_ylim()
                # plot outlier points as red crosses at top
                self.ax.plot(out_vals, np.full_like(out_vals, ymax*0.95), 'x', color='red', label='Outliers')
                self.ax.legend()
        self.ax.set_title(f"Distribution: {col}")
        self.canvas.draw()

    def clear_plot(self):
        self.ax.clear()
        self.canvas.draw()

# ---------- Main ----------
if __name__ == "__main__":
    app = ImputeOutlierApp()
    app.mainloop()
