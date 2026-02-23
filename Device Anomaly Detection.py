import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.ensemble import IsolationForest
from sklearn.preprocessing import StandardScaler

class DeviceAnomalyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Device Anomaly Detection Tool")
        self.root.geometry("600x500")

        self.data = None
        self.model = None
        self.scaler = StandardScaler()

        tk.Label(root, text="Device Anomaly Detection",
                 font=("Arial", 18, "bold")).pack(pady=10)

        tk.Button(root, text="Load CSV Data", command=self.load_data).pack(pady=5)
        tk.Button(root, text="Train Anomaly Model", command=self.train_model).pack(pady=5)
        tk.Button(root, text="Detect Anomalies", command=self.detect_anomalies).pack(pady=5)
        tk.Button(root, text="Visualize Anomalies", command=self.visualize).pack(pady=5)

        self.result_label = tk.Label(root, text="", font=("Arial", 12))
        self.result_label.pack(pady=20)

    def load_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.data = pd.read_csv(file_path)
            messagebox.showinfo("Success", "Data Loaded Successfully")

    def train_model(self):
        if self.data is None:
            messagebox.showerror("Error", "Load data first")
            return

        scaled_data = self.scaler.fit_transform(self.data)

        self.model = IsolationForest(contamination=0.05, random_state=42)
        self.model.fit(scaled_data)

        self.result_label.config(text="Anomaly Detection Model Trained Successfully")

    def detect_anomalies(self):
        if self.model is None:
            messagebox.showerror("Error", "Train model first")
            return

        scaled_data = self.scaler.transform(self.data)
        predictions = self.model.predict(scaled_data)

        self.data["Anomaly"] = predictions
        anomaly_count = sum(predictions == -1)
        total = len(predictions)

        anomaly_percent = (anomaly_count / total) * 100

        self.result_label.config(
            text=f"Total Records: {total}\nAnomalies Detected: {anomaly_count}\nAnomaly %: {anomaly_percent:.2f}%"
        )

    def visualize(self):
        if "Anomaly" not in self.data.columns:
            messagebox.showerror("Error", "Detect anomalies first")
            return

        plt.figure()
        plt.scatter(range(len(self.data)), self.data.iloc[:, 0],
                    c=self.data["Anomaly"])
        plt.title("Anomaly Detection Visualization")
        plt.xlabel("Index")
        plt.ylabel(self.data.columns[0])
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = DeviceAnomalyApp(root)
    root.mainloop()