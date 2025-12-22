import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, r2_score
import matplotlib.pyplot as plt

class RandomForestExplorer:
    def __init__(self, root):
        self.root = root
        self.root.title("Random Forest Explorer & Feature Ranking Tool")
        self.root.geometry("700x500")

        self.data = None

        tk.Button(root, text="Load CSV Dataset", command=self.load_data).pack(pady=10)

        self.target_label = tk.Label(root, text="Select Target Column")
        self.target_label.pack()

        self.target_combo = ttk.Combobox(root)
        self.target_combo.pack(pady=5)

        tk.Button(root, text="Train Random Forest", command=self.train_model).pack(pady=10)

        self.result_text = tk.Text(root, height=10)
        self.result_text.pack(pady=10)

        tk.Button(root, text="Show Feature Importance", command=self.show_feature_importance).pack(pady=10)

    def load_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.data = pd.read_csv(file_path)
            self.target_combo['values'] = list(self.data.columns)
            messagebox.showinfo("Success", "Dataset Loaded Successfully")

    def train_model(self):
        target = self.target_combo.get()
        if not target:
            messagebox.showerror("Error", "Please select target column")
            return

        X = self.data.drop(columns=[target])
        y = self.data[target]

        X = X.select_dtypes(include=['number'])

        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        if y.nunique() <= 10:
            self.model = RandomForestClassifier(n_estimators=100)
            self.model.fit(X_train, y_train)
            preds = self.model.predict(X_test)
            score = accuracy_score(y_test, preds)
            metric = f"Accuracy: {score:.4f}"
        else:
            self.model = RandomForestRegressor(n_estimators=100)
            self.model.fit(X_train, y_train)
            preds = self.model.predict(X_test)
            score = r2_score(y_test, preds)
            metric = f"R² Score: {score:.4f}"

        self.features = X.columns
        self.importances = self.model.feature_importances_

        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, metric + "\n\nFeature Ranking:\n")

        for f, imp in sorted(zip(self.features, self.importances), key=lambda x: x[1], reverse=True):
            self.result_text.insert(tk.END, f"{f}: {imp:.4f}\n")

    def show_feature_importance(self):
        if not hasattr(self, 'importances'):
            messagebox.showerror("Error", "Train model first")
            return

        plt.figure()
        plt.barh(self.features, self.importances)
        plt.xlabel("Importance Score")
        plt.title("Feature Importance Ranking")
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = RandomForestExplorer(root)
    root.mainloop()
