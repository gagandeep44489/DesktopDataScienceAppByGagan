import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class WearableAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Wearable Data Analyzer")
        self.root.geometry("1000x650")

        self.data = None
        self.create_widgets()
        self.create_plot()

    def create_widgets(self):
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10)

        tk.Button(top_frame, text="Load CSV", command=self.load_csv).grid(row=0, column=0, padx=10)
        tk.Button(top_frame, text="Show Stats", command=self.show_stats).grid(row=0, column=1, padx=10)
        tk.Button(top_frame, text="Clear Graph", command=self.clear_graph).grid(row=0, column=2, padx=10)

        self.metric_var = tk.StringVar(value="HeartRate")
        tk.OptionMenu(top_frame, self.metric_var, "HeartRate", "Steps", "Calories").grid(row=0, column=3, padx=10)

        tk.Button(top_frame, text="Plot Metric", command=self.plot_metric).grid(row=0, column=4, padx=10)

    def create_plot(self):
        self.figure, self.ax = plt.subplots(figsize=(9,5))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.root)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.ax.set_title("Wearable Data Visualization")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Value")

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            try:
                self.data = pd.read_csv(file_path)
                messagebox.showinfo("Success", "CSV Loaded Successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file\n{e}")

    def plot_metric(self):
        if self.data is None:
            messagebox.showwarning("Warning", "Please load a CSV file first.")
            return

        metric = self.metric_var.get()

        if metric not in self.data.columns:
            messagebox.showerror("Error", f"{metric} column not found in CSV.")
            return

        self.ax.clear()
        self.ax.plot(self.data["Time"], self.data[metric], marker='o')
        self.ax.set_title(f"{metric} Over Time")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel(metric)
        self.ax.grid(True)
        self.canvas.draw()

    def show_stats(self):
        if self.data is None:
            messagebox.showwarning("Warning", "Please load a CSV file first.")
            return

        metric = self.metric_var.get()

        if metric not in self.data.columns:
            messagebox.showerror("Error", f"{metric} column not found in CSV.")
            return

        avg = self.data[metric].mean()
        max_val = self.data[metric].max()
        min_val = self.data[metric].min()

        messagebox.showinfo(
            "Statistics",
            f"{metric} Statistics:\n\n"
            f"Average: {avg:.2f}\n"
            f"Maximum: {max_val}\n"
            f"Minimum: {min_val}"
        )

    def clear_graph(self):
        self.ax.clear()
        self.ax.set_title("Wearable Data Visualization")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Value")
        self.canvas.draw()


if __name__ == "__main__":
    root = tk.Tk()
    app = WearableAnalyzer(root)
    root.mainloop()