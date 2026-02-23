import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import threading
import time

class IoTDataAggregatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("IoT Data Aggregator")
        self.root.geometry("700x550")

        self.data = None
        self.streaming = False

        tk.Label(root, text="IoT Data Aggregator",
                 font=("Arial", 18, "bold")).pack(pady=10)

        tk.Button(root, text="Load CSV Data", command=self.load_data).pack(pady=5)
        tk.Button(root, text="Start Streaming Simulation", command=self.start_stream).pack(pady=5)
        tk.Button(root, text="Stop Streaming", command=self.stop_stream).pack(pady=5)
        tk.Button(root, text="Aggregate Data", command=self.aggregate_data).pack(pady=5)
        tk.Button(root, text="Visualize Temperature", command=self.visualize).pack(pady=5)
        tk.Button(root, text="Export Aggregated Report", command=self.export_report).pack(pady=5)

        self.result_label = tk.Label(root, text="", font=("Arial", 12))
        self.result_label.pack(pady=20)

    def load_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.data = pd.read_csv(file_path)
            messagebox.showinfo("Success", "Data Loaded Successfully")

    def stream_simulation(self):
        for i in range(len(self.data)):
            if not self.streaming:
                break
            current_row = self.data.iloc[i]
            self.result_label.config(text=f"Streaming Data:\n{current_row}")
            time.sleep(1)

    def start_stream(self):
        if self.data is None:
            messagebox.showerror("Error", "Load data first")
            return
        self.streaming = True
        threading.Thread(target=self.stream_simulation).start()

    def stop_stream(self):
        self.streaming = False
        self.result_label.config(text="Streaming Stopped")

    def aggregate_data(self):
        if self.data is None:
            messagebox.showerror("Error", "Load data first")
            return

        summary = self.data.describe().loc[["mean", "min", "max", "std"]]
        self.summary = summary
        self.result_label.config(text=f"Aggregation Complete\n\n{summary}")

    def visualize(self):
        if self.data is None:
            messagebox.showerror("Error", "Load data first")
            return

        plt.figure()
        plt.plot(self.data["temperature"])
        plt.title("Temperature Trend")
        plt.xlabel("Index")
        plt.ylabel("Temperature")
        plt.show()

    def export_report(self):
        if not hasattr(self, "summary"):
            messagebox.showerror("Error", "Aggregate data first")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".csv")
        if file_path:
            self.summary.to_csv(file_path)
            messagebox.showinfo("Success", "Report Exported Successfully")

if __name__ == "__main__":
    root = tk.Tk()
    app = IoTDataAggregatorApp(root)
    root.mainloop()