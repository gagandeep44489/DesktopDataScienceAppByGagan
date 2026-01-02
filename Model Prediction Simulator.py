import tkinter as tk
from tkinter import ttk, messagebox
import numpy as np
import joblib
from sklearn.datasets import load_iris
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier

# Train demo models (one-time)
def train_and_save_models():
    iris = load_iris()
    X, y = iris.data, iris.target

    lr = LogisticRegression(max_iter=200)
    rf = RandomForestClassifier()

    lr.fit(X, y)
    rf.fit(X, y)

    joblib.dump(lr, "logistic_model.pkl")
    joblib.dump(rf, "rf_model.pkl")

train_and_save_models()

# GUI App
class PredictionSimulator:
    def __init__(self, root):
        self.root = root
        self.root.title("Model Prediction Simulator")
        self.root.geometry("500x450")

        self.models = {
            "Logistic Regression": joblib.load("logistic_model.pkl"),
            "Random Forest": joblib.load("rf_model.pkl")
        }

        self.features = load_iris().feature_names
        self.create_widgets()

    def create_widgets(self):
        title = tk.Label(self.root, text="Model Prediction Simulator",
                         font=("Arial", 14, "bold"))
        title.pack(pady=10)

        self.entries = []
        for feature in self.features:
            frame = tk.Frame(self.root)
            frame.pack(pady=3)

            tk.Label(frame, text=feature, width=20, anchor="w").pack(side="left")
            entry = tk.Entry(frame, width=10)
            entry.pack(side="left")
            self.entries.append(entry)

        model_frame = tk.Frame(self.root)
        model_frame.pack(pady=10)

        tk.Label(model_frame, text="Select Model:").pack(side="left")
        self.model_var = tk.StringVar()
        self.model_box = ttk.Combobox(
            model_frame,
            textvariable=self.model_var,
            values=list(self.models.keys()),
            state="readonly"
        )
        self.model_box.current(0)
        self.model_box.pack(side="left")

        btn = tk.Button(self.root, text="Predict",
                        command=self.predict)
        btn.pack(pady=15)

        self.result_label = tk.Label(self.root, text="Prediction: ",
                                     font=("Arial", 12, "bold"))
        self.result_label.pack(pady=10)

    def predict(self):
        try:
            values = [float(e.get()) for e in self.entries]
            X = np.array(values).reshape(1, -1)

            model = self.models[self.model_var.get()]
            pred = model.predict(X)[0]

            if hasattr(model, "predict_proba"):
                prob = model.predict_proba(X).max()
                result = f"Class: {pred} | Confidence: {prob:.2f}"
            else:
                result = f"Prediction: {pred}"

            self.result_label.config(text=result)

        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = PredictionSimulator(root)
    root.mainloop()
