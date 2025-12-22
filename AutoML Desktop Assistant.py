import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt

from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import accuracy_score, r2_score

from sklearn.linear_model import LogisticRegression, LinearRegression
from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
from sklearn.svm import SVC, SVR
from sklearn.neighbors import KNeighborsClassifier, KNeighborsRegressor

class AutoMLAssistant:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoML Desktop Assistant")
        self.root.geometry("800x600")

        self.data = None
        self.results = []

        tk.Button(root, text="Load CSV Dataset", command=self.load_data).pack(pady=10)

        tk.Label(root, text="Select Target Column").pack()
        self.target_combo = ttk.Combobox(root)
        self.target_combo.pack(pady=5)

        tk.Button(root, text="Run AutoML", command=self.run_automl).pack(pady=10)

        self.output = tk.Text(root, height=14)
        self.output.pack(pady=10)

        tk.Button(root, text="Visualize Model Comparison", command=self.visualize).pack(pady=10)

    def load_data(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path:
            self.data = pd.read_csv(path)
            self.target_combo["values"] = list(self.data.columns)
            messagebox.showinfo("Success", "Dataset Loaded")

    def run_automl(self):
        if self.data is None:
            messagebox.showerror("Error", "Load dataset first")
            return

        target = self.target_combo.get()
        if not target:
            messagebox.showerror("Error", "Select target column")
            return

        X = self.data.drop(columns=[target]).select_dtypes(include="number")
        y = self.data[target]

        scaler = StandardScaler()
        X = scaler.fit_transform(X)

        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        self.results.clear()
        self.output.delete("1.0", tk.END)

        is_classification = y.nunique() <= 10

        if is_classification:
            models = {
                "Logistic Regression": LogisticRegression(max_iter=500),
                "Random Forest": RandomForestClassifier(),
                "SVM": SVC(),
                "KNN": KNeighborsClassifier()
            }
            metric_name = "Accuracy"
        else:
            models = {
                "Linear Regression": LinearRegression(),
                "Random Forest": RandomForestRegressor(),
                "SVR": SVR(),
                "KNN": KNeighborsRegressor()
            }
            metric_name = "R² Score"

        for name, model in models.items():
            model.fit(X_train, y_train)
            preds = model.predict(X_test)

            if is_classification:
                score = accuracy_score(y_test, preds)
            else:
                score = r2_score(y_test, preds)

            self.results.append((name, score))
            self.output.insert(tk.END, f"{name} → {metric_name}: {score:.4f}\n")

        best = max(self.results, key=lambda x: x[1])
        self.output.insert(tk.END, f"\nBest Model: {best[0]} ({metric_name}: {best[1]:.4f})")

    def visualize(self):
        if not self.results:
            messagebox.showerror("Error", "Run AutoML first")
            return

        names, scores = zip(*self.results)

        plt.figure()
        plt.bar(names, scores)
        plt.xlabel("Model")
        plt.ylabel("Score")
        plt.title("AutoML Model Performance Comparison")
        plt.xticks(rotation=30)
        plt.tight_layout()
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoMLAssistant(root)
    root.mainloop()
