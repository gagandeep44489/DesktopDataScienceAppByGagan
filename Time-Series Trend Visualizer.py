import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class TimeSeriesTrendVisualizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Time-Series Trend Visualizer")
        self.root.geometry("900x600")

        self.df = None

        # Top frame for controls
        control_frame = tk.Frame(root)
        control_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

        load_btn = tk.Button(control_frame, text="Load CSV", command=self.load_csv)
        load_btn.pack(side=tk.LEFT, padx=5)

        tk.Label(control_frame, text="Date Column:").pack(side=tk.LEFT, padx=5)
        self.date_col_var = tk.StringVar()
        self.date_col_menu = tk.OptionMenu(control_frame, self.date_col_var, "")
        self.date_col_menu.pack(side=tk.LEFT, padx=5)

        tk.Label(control_frame, text="Value Column:").pack(side=tk.LEFT, padx=5)
        self.value_col_var = tk.StringVar()
        self.value_col_menu = tk.OptionMenu(control_frame, self.value_col_var, "")
        self.value_col_menu.pack(side=tk.LEFT, padx=5)

        plot_btn = tk.Button(control_frame, text="Plot Trend", command=self.plot_trend)
        plot_btn.pack(side=tk.LEFT, padx=10)

        # Frame for plot
        self.plot_frame = tk.Frame(root)
        self.plot_frame.pack(fill=tk.BOTH, expand=True)

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return
        try:
            self.df = pd.read_csv(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV:\n{e}")
            return

        cols = list(self.df.columns)
        if len(cols) < 2:
            messagebox.showwarning("Warning", "CSV must contain at least two columns.")
            return

        self.update_option_menu(self.date_col_menu, self.date_col_var, cols)
        self.update_option_menu(self.value_col_menu, self.value_col_var, cols)

        self.date_col_var.set(cols[0])
        self.value_col_var.set(cols[1])

    def update_option_menu(self, menu, var, options):
        menu['menu'].delete(0, 'end')
        for opt in options:
            menu['menu'].add_command(label=opt, command=tk._setit(var, opt))

    def plot_trend(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Please load a CSV file first.")
            return

        date_col = self.date_col_var.get()
        value_col = self.value_col_var.get()

        if date_col == "" or value_col == "":
            messagebox.showwarning("Warning", "Please select columns.")
            return

        try:
            data = self.df.copy()
            data[date_col] = pd.to_datetime(data[date_col])
            data = data.sort_values(by=date_col)
        except Exception as e:
            messagebox.showerror("Error", f"Date parsing failed:\n{e}")
            return

        for widget in self.plot_frame.winfo_children():
            widget.destroy()

        fig, ax = plt.subplots(figsize=(8, 5))
        ax.plot(data[date_col], data[value_col])
        ax.set_title("Time-Series Trend")
        ax.set_xlabel("Time")
        ax.set_ylabel(value_col)
        ax.grid(True)

        canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

if __name__ == "__main__":
    root = tk.Tk()
    app = TimeSeriesTrendVisualizer(root)
    root.mainloop()
