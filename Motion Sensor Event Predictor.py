import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score
from sklearn.preprocessing import StandardScaler

class MotionSensorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Motion Sensor Event Predictor")
        self.root.geometry("650x550")

        self.data = None
        self.model = None
        self.scaler = StandardScaler()

        tk.Label(root, text="Motion Sensor Event Predictor",
                 font=("Arial", 18, "bold")).pack(pady=10)

        tk.Button(root, text="Load CSV Data", command=self.load_data).pack(pady=5)
        tk.Button(root, text="Train Model", command=self.train_model).pack(pady=5)
        tk.Button(root, text="Predict New Motion", command=self.predict_motion).pack(pady=5)
        tk.Button(root, text="Visualize Acceleration", command=self.visualize).pack(pady=5)

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

        X = self.data.drop("event", axis=1)
        y = self.data["event"]

        X_scaled = self.scaler.fit_transform(X)

        X_train, X_test, y_train, y_test = train_test_split(
            X_scaled, y, test_size=0.2, random_state=42
        )

        self.model = RandomForestClassifier(n_estimators=100)
        self.model.fit(X_train, y_train)

        y_pred = self.model.predict(X_test)
        accuracy = accuracy_score(y_test, y_pred)

        self.result_label.config(text=f"Model Trained\nAccuracy: {accuracy:.2f}")

    def predict_motion(self):
        if self.model is None:
            messagebox.showerror("Error", "Train model first")
            return

        input_window = tk.Toplevel(self.root)
        input_window.title("Enter Motion Values")

        entries = {}
        features = ["acc_x", "acc_y", "acc_z", "gyro_x", "gyro_y", "gyro_z"]

        for feature in features:
            tk.Label(input_window, text=feature).pack()
            entry = tk.Entry(input_window)
            entry.pack()
            entries[feature] = entry

        def submit():
            values = [[float(entries[f].get()) for f in features]]
            values_scaled = self.scaler.transform(values)

            prediction = self.model.predict(values_scaled)[0]
            probability = self.model.predict_proba(values_scaled)[0][1]

            status = "⚠ Motion Event Detected" if prediction == 1 else "✅ Normal Motion"
            self.result_label.config(
                text=f"Prediction: {status}\nEvent Probability: {probability:.2f}"
            )
            input_window.destroy()

        tk.Button(input_window, text="Predict", command=submit).pack(pady=10)

    def visualize(self):
        if self.data is None:
            messagebox.showerror("Error", "Load data first")
            return

        plt.figure()
        plt.plot(self.data["acc_x"], label="acc_x")
        plt.plot(self.data["acc_y"], label="acc_y")
        plt.plot(self.data["acc_z"], label="acc_z")
        plt.title("Acceleration Sensor Trends")
        plt.legend()
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = MotionSensorApp(root)
    root.mainloop()