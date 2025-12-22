import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt

from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import accuracy_score
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier
from sklearn.svm import SVC

class HyperparameterTuner:
    def __init__(self, root):
        self.root = root
        self.root.title("Hyperparameter Tuning Interface")
        self.root.geometry("780x580")

        self.data = None
        self.results = []

        tk.Button(root, text="Load CSV Dataset", command=self.load_data).pack(pady=10)

        tk.Label(root, text="Select Target Column").pack()
        self.target_combo = ttk.Combobox(root)
        self.target_combo.pack(pady=5)

        tk.Label(root, text="Select Model").pack()
        self.model_combo = ttk.Combobox(
            root,
            values=["Logistic Regression", "Random Forest", "SVM"]
        )
        self.model_combo.pack(pady=5)

        tk.Label(root, text="Hyperparameter Values (comma-separated)").pack()
        self.param_entry = tk.Entry(root, width=50)
        self.param_entry.pack(pady=5)
        self.param_entry.insert(0, "Example: 10,50,100")

        tk.Button(root, text="Run Hyperparameter Tuning", command=self.run_tuning).pack(pady=10)

        self.output = tk.Text(root, height=12)
        self.output.pack(pady=10)

        tk.Button(root, text="Visualize Results", command=self.visualize_results).pack(pady=10)

    def load_data(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path:
            self.data = pd.read_csv(path)
            self.target_combo["values"] = list(self.data.columns)
            messagebox.showinfo("Success", "Dataset Loaded Successfully")

    def run_tuning(self):
        if self.data is None:
            messagebox.showerror("Error", "Load dataset first")
            return

        target = self.target_combo.get()
        model_name = self.model_combo.get()

        if not target or not model_name:
            messagebox.showerror("Error", "Select target and model")
            return

        param_values = self.param_entry.get().split(",")
        param_values = [int(p.strip()) for p in param_values]

        X = self.data.drop(columns=[target]).select_dtypes(include="number")
        y = self.data[target]

        scaler = StandardScaler()
        X = scaler.fit_transform(X)

        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        self.results.clear()
        self.output.delete("1.0", tk.END)

        for val in param_values:
            if model_name == "Logistic Regression":
                model = LogisticRegression(C=val, max_iter=500)
                label = f"C={val}"

            elif model_name == "Random Forest":
                model = RandomForestClassifier(n_estimators=val)
                label = f"Trees={val}"

            else:
                model = SVC(C=val)
                label = f"C={val}"

            model.fit(X_train, y_train)
            preds = model.predict(X_test)
            acc = accuracy_score(y_test, preds)

            self.results.append((label, acc))
            self.output.insert(tk.END, f"{label} → Accuracy: {acc:.4f}\n")

        best = max(self.results, key=lambda x: x[1])
        self.output.insert(tk.END, f"\nBest Configuration: {best[0]} ({best[1]:.4f})")

    def visualize_results(self):
        if not self.results:
            messagebox.showerror("Error", "Run tuning first")
            return

        labels, scores = zip(*self.results)

        plt.figure()
        plt.plot(labels, scores, marker='o')
        plt.xlabel("Hyperparameter Setting")
        plt.ylabel("Accuracy")
        plt.title("Hyperparameter Tuning Performance")
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = HyperparameterTuner(root)
    root.mainloop()
