import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report, confusion_matrix
import matplotlib.pyplot as plt
import seaborn as sns

class FraudDetectionApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Fraud Detection Simulator")
        self.root.geometry("600x500")

        self.data = None
        self.model = None

        tk.Button(root, text="Load Transaction Data", command=self.load_data).pack(pady=10)
        tk.Button(root, text="Train Model", command=self.train_model).pack(pady=10)
        tk.Button(root, text="Evaluate Model", command=self.evaluate_model).pack(pady=10)
        tk.Button(root, text="Simulate Transaction", command=self.simulate_transaction).pack(pady=10)

        self.result_label = tk.Label(root, text="", font=("Arial", 11))
        self.result_label.pack(pady=20)

    def load_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.data = pd.read_csv(file_path)
            messagebox.showinfo("Success", "Transaction Data Loaded")

    def preprocess(self):
        df = self.data.copy()
        df = pd.get_dummies(df, drop_first=True)
        return df

    def train_model(self):
        if self.data is None:
            messagebox.showerror("Error", "Load data first")
            return

        df = self.preprocess()
        X = df.drop("Is_Fraud", axis=1)
        y = df["Is_Fraud"]

        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=0.2, random_state=42
        )

        self.model = LogisticRegression(max_iter=1000)
        self.model.fit(X_train, y_train)

        self.X_test = X_test
        self.y_test = y_test

        messagebox.showinfo("Success", "Fraud Model Trained")

    def evaluate_model(self):
        if self.model is None:
            messagebox.showerror("Error", "Train model first")
            return

        predictions = self.model.predict(self.X_test)
        cm = confusion_matrix(self.y_test, predictions)

        plt.figure()
        sns.heatmap(cm, annot=True, fmt='d')
        plt.title("Confusion Matrix")
        plt.xlabel("Predicted")
        plt.ylabel("Actual")
        plt.show()

        report = classification_report(self.y_test, predictions)
        self.result_label.config(text=report)

    def simulate_transaction(self):
        if self.model is None:
            messagebox.showerror("Error", "Train model first")
            return

        sample = self.X_test.iloc[0].values.reshape(1, -1)
        prob = self.model.predict_proba(sample)[0][1]

        result = "FRAUD" if prob > 0.5 else "LEGITIMATE"

        self.result_label.config(
            text=f"Fraud Probability: {round(prob,2)}\nPrediction: {result}"
        )

if __name__ == "__main__":
    root = tk.Tk()
    app = FraudDetectionApp(root)
    root.mainloop()
