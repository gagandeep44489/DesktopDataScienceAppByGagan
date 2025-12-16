import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from tkinter import ttk

class DataFilterQueryBuilder:
    def __init__(self, root):
        self.root = root
        self.root.title("Custom Data Filter & Query Builder")
        self.root.geometry("1000x600")

        self.df = None

        # Top controls
        top_frame = tk.Frame(root)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        load_btn = tk.Button(top_frame, text="Load CSV", command=self.load_csv)
        load_btn.pack(side=tk.LEFT, padx=5)

        tk.Label(top_frame, text="Column").pack(side=tk.LEFT, padx=5)
        self.column_var = tk.StringVar()
        self.column_menu = tk.OptionMenu(top_frame, self.column_var, "")
        self.column_menu.pack(side=tk.LEFT)

        tk.Label(top_frame, text="Operator").pack(side=tk.LEFT, padx=5)
        self.operator_var = tk.StringVar(value="==")
        operator_menu = tk.OptionMenu(top_frame, self.operator_var, "==", "!=", ">", "<", ">=", "<=", "contains")
        operator_menu.pack(side=tk.LEFT)

        tk.Label(top_frame, text="Value").pack(side=tk.LEFT, padx=5)
        self.value_entry = tk.Entry(top_frame)
        self.value_entry.pack(side=tk.LEFT, padx=5)

        apply_btn = tk.Button(top_frame, text="Apply Filter", command=self.apply_filter)
        apply_btn.pack(side=tk.LEFT, padx=10)

        reset_btn = tk.Button(top_frame, text="Reset", command=self.reset_table)
        reset_btn.pack(side=tk.LEFT)

        # Table frame
        table_frame = tk.Frame(root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(table_frame, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return
        try:
            self.df = pd.read_csv(file_path)
            self.populate_table(self.df)
            self.update_column_menu(self.df.columns)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def update_column_menu(self, columns):
        self.column_menu['menu'].delete(0, 'end')
        for col in columns:
            self.column_menu['menu'].add_command(label=col, command=tk._setit(self.column_var, col))
        self.column_var.set(columns[0])

    def populate_table(self, dataframe):
        self.tree.delete(*self.tree.get_children())
        self.tree['columns'] = list(dataframe.columns)
        for col in dataframe.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)
        for _, row in dataframe.iterrows():
            self.tree.insert("", tk.END, values=list(row))

    def apply_filter(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Load a CSV first")
            return

        col = self.column_var.get()
        op = self.operator_var.get()
        val = self.value_entry.get()

        try:
            if op == "==":
                filtered = self.df[self.df[col] == self.cast_value(col, val)]
            elif op == "!=":
                filtered = self.df[self.df[col] != self.cast_value(col, val)]
            elif op == ">":
                filtered = self.df[self.df[col] > self.cast_value(col, val)]
            elif op == "<":
                filtered = self.df[self.df[col] < self.cast_value(col, val)]
            elif op == ">=":
                filtered = self.df[self.df[col] >= self.cast_value(col, val)]
            elif op == "<=":
                filtered = self.df[self.df[col] <= self.cast_value(col, val)]
            elif op == "contains":
                filtered = self.df[self.df[col].astype(str).str.contains(val, na=False)]
            else:
                return

            self.populate_table(filtered)
        except Exception as e:
            messagebox.showerror("Filter Error", str(e))

    def cast_value(self, col, val):
        if pd.api.types.is_numeric_dtype(self.df[col]):
            return float(val)
        return val

    def reset_table(self):
        if self.df is not None:
            self.populate_table(self.df)

if __name__ == "__main__":
    root = tk.Tk()
    app = DataFilterQueryBuilder(root)
    root.mainloop()