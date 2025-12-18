"""
Multi-Dataset Comparison Tool
Python Desktop Application using Tkinter

Features:
- Load multiple CSV datasets
- Compare rows, columns, and schema
- Compare missing values
- Display summary comparison

Suitable for data analysts, students, and professionals
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class DatasetComparisonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Dataset Comparison Tool")
        self.root.geometry("900x600")

        self.datasets = {}

        title = tk.Label(root, text="Multi-Dataset Comparison Tool", font=("Arial", 16, "bold"))
        title.pack(pady=10)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        load_btn = tk.Button(btn_frame, text="Load CSV Files", width=25, command=self.load_files)
        load_btn.grid(row=0, column=0, padx=10)

        compare_btn = tk.Button(btn_frame, text="Compare Datasets", width=25, command=self.compare_datasets)
        compare_btn.grid(row=0, column=1, padx=10)

        self.text_area = tk.Text(root, wrap=tk.WORD)
        self.text_area.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

    def load_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("CSV Files", "*.csv")])
        if file_paths:
            self.datasets.clear()
            self.text_area.delete(1.0, tk.END)

            for path in file_paths:
                try:
                    df = pd.read_csv(path)
                    self.datasets[os.path.basename(path)] = df
                    self.text_area.insert(tk.END, f"Loaded: {os.path.basename(path)} | Rows: {df.shape[0]} | Columns: {df.shape[1]}\n")
                except Exception as e:
                    messagebox.showerror("Error", str(e))

    def compare_datasets(self):
        if len(self.datasets) < 2:
            messagebox.showwarning("Warning", "Please load at least two datasets")
            return

        self.text_area.insert(tk.END, "\n--- Dataset Comparison Report ---\n")

        columns_sets = {name: set(df.columns) for name, df in self.datasets.items()}
        all_columns = set.union(*columns_sets.values())

        self.text_area.insert(tk.END, f"\nTotal Unique Columns Across Datasets: {len(all_columns)}\n")

        for name, df in self.datasets.items():
            self.text_area.insert(tk.END, f"\nDataset: {name}\n")
            self.text_area.insert(tk.END, f"Rows: {df.shape[0]} | Columns: {df.shape[1]}\n")
            self.text_area.insert(tk.END, "Missing Values:\n")
            self.text_area.insert(tk.END, str(df.isnull().sum()) + "\n")

        self.text_area.insert(tk.END, "\nColumn Presence Comparison:\n")
        for col in sorted(all_columns):
            present_in = [name for name, cols in columns_sets.items() if col in cols]
            self.text_area.insert(tk.END, f"{col}: Present in {', '.join(present_in)}\n")


if __name__ == "__main__":
    root = tk.Tk()
    app = DatasetComparisonApp(root)
    root.mainloop()
