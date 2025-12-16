import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class CorrelationHeatmapApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Correlation Heatmap Generator")
        self.root.geometry("900x600")

        self.df = None

        # Control frame
        control_frame = tk.Frame(root)
        control_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

        load_btn = tk.Button(control_frame, text="Load CSV", command=self.load_csv)
        load_btn.pack(side=tk.LEFT, padx=5)

        plot_btn = tk.Button(control_frame, text="Generate Heatmap", command=self.plot_heatmap)
        plot_btn.pack(side=tk.LEFT, padx=10)

        info_lbl = tk.Label(control_frame, text="Numeric columns only will be used")
        info_lbl.pack(side=tk.LEFT, padx=10)

        # Plot frame
        self.plot_frame = tk.Frame(root)
        self.plot_frame.pack(fill=tk.BOTH, expand=True)

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return
        try:
            self.df = pd.read_csv(file_path)
            messagebox.showinfo("Success", "CSV file loaded successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV:\n{e}")

    def plot_heatmap(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Please load a CSV file first")
            return

        numeric_df = self.df.select_dtypes(include=['number'])
        if numeric_df.shape[1] < 2:
            messagebox.showwarning("Warning", "Need at least two numeric columns")
            return

        corr_matrix = numeric_df.corr()

        for widget in self.plot_frame.winfo_children():
            widget.destroy()

        fig, ax = plt.subplots(figsize=(8, 6))
        cax = ax.imshow(corr_matrix)
        ax.set_title("Correlation Heatmap")

        ax.set_xticks(range(len(corr_matrix.columns)))
        ax.set_yticks(range(len(corr_matrix.columns)))
        ax.set_xticklabels(corr_matrix.columns, rotation=45, ha='right')
        ax.set_yticklabels(corr_matrix.columns)

        fig.colorbar(cax)

        canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

if __name__ == "__main__":
    root = tk.Tk()
    app = CorrelationHeatmapApp(root)
    root.mainloop()