"""
Automatic Data Summary & Report Generator
Desktop Application using Python (Tkinter)

Features:
- Load CSV file
- Display basic dataset info
- Generate statistical summary
- Export report as text file

Author: ChatGPT
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class DataSummaryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatic Data Summary & Report Generator")
        self.root.geometry("800x600")

        self.data = None

        title = tk.Label(root, text="Automatic Data Summary & Report Generator", font=("Arial", 16, "bold"))
        title.pack(pady=10)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        load_btn = tk.Button(btn_frame, text="Load CSV File", width=20, command=self.load_file)
        load_btn.grid(row=0, column=0, padx=10)

        summary_btn = tk.Button(btn_frame, text="Generate Summary", width=20, command=self.generate_summary)
        summary_btn.grid(row=0, column=1, padx=10)

        export_btn = tk.Button(btn_frame, text="Export Report", width=20, command=self.export_report)
        export_btn.grid(row=0, column=2, padx=10)

        self.text_area = tk.Text(root, wrap=tk.WORD)
        self.text_area.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            try:
                self.data = pd.read_csv(file_path)
                self.text_area.delete(1.0, tk.END)
                self.text_area.insert(tk.END, f"File Loaded Successfully: {os.path.basename(file_path)}\n\n")
                self.text_area.insert(tk.END, f"Rows: {self.data.shape[0]}\n")
                self.text_area.insert(tk.END, f"Columns: {self.data.shape[1]}\n\n")
                self.text_area.insert(tk.END, "Column Names:\n")
                for col in self.data.columns:
                    self.text_area.insert(tk.END, f"- {col}\n")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def generate_summary(self):
        if self.data is None:
            messagebox.showwarning("Warning", "Please load a dataset first")
            return

        self.text_area.insert(tk.END, "\nStatistical Summary:\n")
        self.text_area.insert(tk.END, str(self.data.describe(include='all')))
        self.text_area.insert(tk.END, "\n\nMissing Values:\n")
        self.text_area.insert(tk.END, str(self.data.isnull().sum()))

    def export_report(self):
        if self.data is None:
            messagebox.showwarning("Warning", "No report to export")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".txt")
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(self.text_area.get(1.0, tk.END))
            messagebox.showinfo("Success", "Report exported successfully")


if __name__ == "__main__":
    root = tk.Tk()
    app = DataSummaryApp(root)
    root.mainloop()
