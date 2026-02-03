import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt

class WeatherTrendVisualizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Weather Trend Visualizer")
        self.root.geometry("750x520")

        self.data = None
        self.metric = tk.StringVar()

        self.create_ui()

    def create_ui(self):
        tk.Button(self.root, text="Load Weather CSV", command=self.load_csv, width=25).pack(pady=10)

        tk.Label(self.root, text="Select Weather Metric").pack()
        self.metric_dropdown = tk.OptionMenu(self.root, self.metric, "")
        self.metric_dropdown.pack(pady=5)

        tk.Button(self.root, text="Show Trend", command=self.plot_trend, width=25).pack(pady=10)

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return

        self.data = pd.read_csv(file_path)

        if "Date" not in self.data.columns:
            messagebox.showerror("Error", "CSV must contain 'Date' column")
            return

        numeric_cols = self.data.select_dtypes(include="number").columns.tolist()
        if not numeric_cols:
            messagebox.showerror("Error", "No numeric weather columns found")
            return

        self.metric.set(numeric_cols[0])
        menu = self.metric_dropdown["menu"]
        menu.delete(0, "end")
        for col in numeric_cols:
            menu.add_command(label=col, command=lambda v=col: self.metric.set(v))

        messagebox.showinfo("Success", "Weather data loaded successfully")

    def plot_trend(self):
        if self.data is None:
            messagebox.showerror("Error", "Load weather data first")
            return

        metric = self.metric.get()
        if metric == "":
            messagebox.showerror("Error", "Select a weather metric")
            return

        plt.figure()
        plt.plot(self.data["Date"], self.data[metric])
        plt.title(f"{metric} Trend Over Time")
        plt.xlabel("Date")
        plt.ylabel(metric)
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = WeatherTrendVisualizer(root)
    root.mainloop()
