import tkinter as tk
from tkinter import messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from sklearn.ensemble import RandomForestClassifier

class SmartAgriApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Agriculture Predictor")
        self.root.geometry("1000x700")

        self.create_widgets()
        self.create_plot()
        self.train_model()

    def create_widgets(self):
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=10)

        labels = ["Nitrogen", "Phosphorus", "Potassium", 
                  "Temperature (°C)", "Humidity (%)", 
                  "Rainfall (mm)", "pH"]

        self.entries = {}

        for i, label in enumerate(labels):
            tk.Label(input_frame, text=label).grid(row=i, column=0, padx=5, pady=5)
            entry = tk.Entry(input_frame)
            entry.grid(row=i, column=1, padx=5, pady=5)
            self.entries[label] = entry

        tk.Button(input_frame, text="Predict Crop", command=self.predict_crop).grid(row=8, column=0, pady=10)
        tk.Button(input_frame, text="Clear", command=self.clear_inputs).grid(row=8, column=1, pady=10)

        self.result_label = tk.Label(self.root, text="", font=("Arial", 14))
        self.result_label.pack(pady=10)

    def create_plot(self):
        self.figure, self.ax = plt.subplots(figsize=(8,4))
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.root)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.ax.set_title("Environmental Parameter Overview")
        self.ax.set_ylabel("Values")

    def train_model(self):
        # Dummy dataset (for demonstration)
        data = {
            "N": [90, 85, 60, 40],
            "P": [42, 58, 55, 35],
            "K": [43, 41, 44, 30],
            "Temperature": [20, 25, 30, 35],
            "Humidity": [80, 70, 60, 50],
            "Rainfall": [200, 150, 100, 50],
            "pH": [6.5, 6.0, 5.5, 7.0],
            "Crop": ["Rice", "Wheat", "Maize", "Cotton"]
        }

        df = pd.DataFrame(data)

        X = df[["N", "P", "K", "Temperature", "Humidity", "Rainfall", "pH"]]
        y = df["Crop"]

        self.model = RandomForestClassifier()
        self.model.fit(X, y)

    def predict_crop(self):
        try:
            values = [float(self.entries[label].get()) for label in self.entries]
            input_data = np.array(values).reshape(1, -1)

            prediction = self.model.predict(input_data)[0]
            self.result_label.config(text=f"Recommended Crop: {prediction}")

            self.update_plot(values)

        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numeric values.")

    def update_plot(self, values):
        self.ax.clear()
        parameters = list(self.entries.keys())
        self.ax.bar(parameters, values)
        self.ax.set_title("Environmental Parameter Overview")
        self.ax.set_ylabel("Values")
        self.ax.tick_params(axis='x', rotation=45)
        self.canvas.draw()

    def clear_inputs(self):
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        self.result_label.config(text="")
        self.ax.clear()
        self.canvas.draw()


if __name__ == "__main__":
    root = tk.Tk()
    app = SmartAgriApp(root)
    root.mainloop()