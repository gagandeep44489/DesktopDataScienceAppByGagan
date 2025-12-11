"""
Interactive Data Dashboard Creator
----------------------------------

A standalone desktop app that allows users to:
- Import CSV or Excel files
- Preview data
- Select chart types (Bar, Line, Scatter, Histogram, Pie)
- Choose X/Y axes dynamically
- Apply quick filters
- Add multiple charts to a dashboard-like layout
- Save dashboard charts as images

Developed with tkinter + pandas + matplotlib.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import os

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


class DashboardCreator(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Interactive Data Dashboard Creator")
        self.geometry("1350x800")
        self.minsize(1100, 700)

        self.df = None
        self.file_path = None

        self.build_ui()

    def build_ui(self):

        # ---------------- Toolbar ----------------
        toolbar = ttk.Frame(self, padding=8)
        toolbar.pack(fill='x')

        ttk.Button(toolbar, text="Open File", command=self.open_file).pack(side="left")
        ttk.Button(toolbar, text="Save Chart", command=self.save_chart).pack(side="left", padx=6)

        ttk.Label(toolbar, text="Status:").pack(side="right", padx=(10, 4))
        self.status_var = tk.StringVar(value="Waiting for data...")
        ttk.Label(toolbar, textvariable=self.status_var, foreground="blue").pack(side="right")

        # ---------------- Main Split ----------------
        main_split = ttk.Panedwindow(self, orient='horizontal')
        main_split.pack(fill='both', expand=True)

        # ---------- Left Panel (Config) ----------
        left_panel = ttk.Frame(main_split, width=350, padding=10)
        main_split.add(left_panel, weight=1)

        # File label
        ttk.Label(left_panel, text="Loaded File:").pack(anchor="w")
        self.file_label = ttk.Label(left_panel, text="None")
        self.file_label.pack(anchor="w", pady=(0, 10))

        # Column selectors
        ttk.Label(left_panel, text="X-axis Column:").pack(anchor="w")
        self.x_var = tk.StringVar()
        self.x_menu = ttk.Combobox(left_panel, textvariable=self.x_var, state='readonly')
        self.x_menu.pack(fill='x', pady=3)

        ttk.Label(left_panel, text="Y-axis Column:").pack(anchor="w")
        self.y_var = tk.StringVar()
        self.y_menu = ttk.Combobox(left_panel, textvariable=self.y_var, state='readonly')
        self.y_menu.pack(fill='x', pady=3)

        # Chart types
        ttk.Label(left_panel, text="Chart Type:").pack(anchor="w", pady=(10, 0))
        self.chart_type = tk.StringVar(value="line")
        chart_options = ["line", "bar", "scatter", "histogram", "pie"]
        self.chart_menu = ttk.Combobox(left_panel, textvariable=self.chart_type, values=chart_options, state='readonly')
        self.chart_menu.pack(fill='x')

        # Filter options
        ttk.Label(left_panel, text="Filter (optional):").pack(anchor="w", pady=(10, 0))
        self.filter_column = tk.StringVar()
        self.filter_value = tk.StringVar()

        self.filter_col_menu = ttk.Combobox(left_panel, textvariable=self.filter_column, state='readonly')
        self.filter_col_menu.pack(fill='x', pady=3)

        self.filter_entry = ttk.Entry(left_panel, textvariable=self.filter_value)
        self.filter_entry.pack(fill='x', pady=(3, 10))

        ttk.Button(left_panel, text="Generate Chart", command=self.generate_chart).pack(fill='x', pady=(6, 0))

        # Dashboard instructions
        ttk.Label(left_panel, text="Tip: Add multiple charts\nto build a dashboard.",
                  foreground="gray").pack(anchor="w", pady=10)

        # ---------- Right Panel (Dashboard) ----------
        right_panel = ttk.Frame(main_split)
        main_split.add(right_panel, weight=3)

        ttk.Label(right_panel, text="Dashboard View", font=("Segoe UI", 12, "bold")).pack(anchor='w', padx=10, pady=5)

        # Canvas container for charts
        self.chart_area = ttk.Frame(right_panel)
        self.chart_area.pack(fill='both', expand=True)

    # ----------------------------------------------------------
    # Load File
    # ----------------------------------------------------------
    def open_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx;*.xls"), ("All Files", "*.*")]
        )

        if not path:
            return

        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in [".xlsx", ".xls"]:
                self.df = pd.read_excel(path)
            else:
                self.df = pd.read_csv(path)

            self.file_path = path
            self.file_label.config(text=os.path.basename(path))
            self.status_var.set(f"Loaded: {os.path.basename(path)}")

            cols = list(self.df.columns)
            self.x_menu['values'] = cols
            self.y_menu['values'] = cols
            self.filter_col_menu['values'] = cols

        except Exception as e:
            messagebox.showerror("File Error", str(e))

    # ----------------------------------------------------------
    # Generate Chart
    # ----------------------------------------------------------
    def generate_chart(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Please load a dataset first.")
            return

        chart_type = self.chart_type.get()
        x = self.x_var.get()
        y = self.y_var.get()

        df = self.df.copy()

        # Optional filter
        col = self.filter_column.get()
        val = self.filter_value.get()

        if col and val:
            df = df[df[col].astype(str) == str(val)]

        fig = Figure(figsize=(5, 3))
        ax = fig.add_subplot(111)

        try:
            if chart_type == "line":
                ax.plot(df[x], df[y])

            elif chart_type == "bar":
                ax.bar(df[x], df[y])

            elif chart_type == "scatter":
                ax.scatter(df[x], df[y])

            elif chart_type == "histogram":
                ax.hist(df[x].dropna(), bins=20)

            elif chart_type == "pie":
                counts = df[x].value_counts()
                ax.pie(counts.values, labels=counts.index, autopct="%1.1f%%")

            ax.set_title(f"{chart_type.capitalize()} Chart")
            ax.set_xlabel(x)
            if chart_type != "pie":
                ax.set_ylabel(y)

        except Exception as e:
            messagebox.showerror("Chart Error", str(e))
            return

        frame = ttk.Frame(self.chart_area, padding=5)
        frame.pack(fill='both', expand=True)

        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)

        ttk.Label(frame, text=f"{chart_type.capitalize()} Chart ({x} vs {y})", foreground="gray").pack()

        self.status_var.set("Chart added to dashboard")

    # ----------------------------------------------------------
    # Save Current Chart
    # ----------------------------------------------------------
    def save_chart(self):
        path = filedialog.asksaveasfilename(defaultextension=".png",
                                            filetypes=[("PNG Image", "*.png")])
        if not path:
            return

        try:
            children = self.chart_area.winfo_children()
            if not children:
                messagebox.showinfo("No Chart", "No chart available to save.")
                return

            # Save last chart
            last_frame = children[-1]
            canvas_widget = None

            for widget in last_frame.winfo_children():
                if isinstance(widget, tk.Canvas):
                    canvas_widget = widget

            if hasattr(last_frame, "children"):
                # But matplotlib canvas is nested
                for child in last_frame.children.values():
                    if hasattr(child, "print_png"):
                        child.print_png(path)

            messagebox.showinfo("Saved", f"Chart saved to:\n{path}")

        except Exception as e:
            messagebox.showerror("Save Error", str(e))


# ---------------------------- RUN APP ------------------------------
if __name__ == "__main__":
    app = DashboardCreator()
    app.mainloop()
