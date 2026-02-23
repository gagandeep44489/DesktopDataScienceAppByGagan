import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score
import joblib

class PredictiveMaintenanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Predictive Maintenance Tool")
        self.root.geometry("600x500")

        self.model = None
        self.data = None

        tk.Label(root, text="Predictive Maintenance Tool", 
                 font=("Arial", 18, "bold")).pack(pady=10)

        tk.Button(root, text="Load CSV Data", command=self.load_data).pack(pady=5)
        tk.Button(root, text="Train Model", command=self.train_model).pack(pady=5)
        tk.Button(root, text="Predict New Data", command=self.predict_data).pack(pady=5)
        tk.Button(root, text="Show Feature Importance", command=self.show_importance).pack(pady=5)

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

        X = self.data.drop("failure", axis=1)
        y = self.data["failure"]

        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        self.model = RandomForestClassifier(n_estimators=100)
        self.model.fit(X_train, y_train)

        y_pred = self.model.predict(X_test)
        accuracy = accuracy_score(y_test, y_pred)

        self.result_label.config(text=f"Model Trained\nAccuracy: {accuracy:.2f}")

    def predict_data(self):
        if self.model is None:
            messagebox.showerror("Error", "Train model first")
            return

        input_window = tk.Toplevel(self.root)
        input_window.title("Enter Sensor Values")

        entries = {}
        features = ["temperature", "vibration", "pressure", "humidity", "rpm"]

        for feature in features:
            tk.Label(input_window, text=feature).pack()
            entry = tk.Entry(input_window)
            entry.pack()
            entries[feature] = entry

        def submit():
            values = [[float(entries[f].get()) for f in features]]
            prediction = self.model.predict(values)[0]
            probability = self.model.predict_proba(values)[0][1]

            status = "⚠ High Failure Risk" if prediction == 1 else "✅ Normal"
            self.result_label.config(
                text=f"Prediction: {status}\nFailure Probability: {probability:.2f}"
            )
            input_window.destroy()

        tk.Button(input_window, text="Predict", command=submit).pack(pady=10)

    def show_importance(self):
        if self.model is None:
            messagebox.showerror("Error", "Train model first")
            return

        importances = self.model.feature_importances_
        features = self.data.drop("failure", axis=1).columns

        plt.figure()
        plt.bar(features, importances)
        plt.title("Feature Importance")
        plt.xticks(rotation=45)
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = PredictiveMaintenanceApp(root)
    root.mainloop()