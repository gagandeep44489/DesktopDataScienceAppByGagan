"""
Classification Model Trainer with Metrics Dashboard
Python Desktop Application using Tkinter, Pandas, Scikit-learn, Matplotlib

Features:
- Load CSV dataset
- Select target column
- Train Classification model (Logistic Regression)
- Display accuracy, precision, recall, F1-score
- Show confusion matrix visualization
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score, confusion_matrix

class ClassificationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Classification Model Trainer & Metrics Dashboard")
        self.root.geometry("950x700")

        self.data = None
        self.model = LogisticRegression(max_iter=1000)

        title = tk.Label(root, text="Classification Model Trainer & Metrics Dashboard", font=("Arial", 16, "bold"))
        title.pack(pady=10)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        load_btn = tk.Button(btn_frame, text="Load CSV Dataset", width=22, command=self.load_data)
        load_btn.grid(row=0, column=0, padx=10)

        train_btn = tk.Button(btn_frame, text="Train Model", width=22, command=self.train_model)
        train_btn.grid(row=0, column=1, padx=10)

        plot_btn = tk.Button(btn_frame, text="Show Confusion Matrix", width=22, command=self.plot_confusion)
        plot_btn.grid(row=0, column=2, padx=10)

        self.target_label = tk.Label(root, text="Select Target Column:")
        self.target_label.pack()

        self.target_var = tk.StringVar()
        self.target_menu = tk.OptionMenu(root, self.target_var, "")
        self.target_menu.pack(pady=5)

        self.text_area = tk.Text(root, wrap=tk.WORD)
        self.text_area.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

    def load_data(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path:
            try:
                self.data = pd.read_csv(path)
                self.text_area.delete(1.0, tk.END)
                self.text_area.insert(tk.END, f"Dataset Loaded Successfully\nRows: {self.data.shape[0]} | Columns: {self.data.shape[1]}\n")

                menu = self.target_menu["menu"]
                menu.delete(0, "end")
                for col in self.data.columns:
                    menu.add_command(label=col, command=lambda value=col: self.target_var.set(value))

                self.target_var.set(self.data.columns[-1])
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def train_model(self):
        if self.data is None:
            messagebox.showwarning("Warning", "Load dataset first")
            return

        target = self.target_var.get()
        X = self.data.drop(columns=[target])
        y = self.data[target]

        X = X.select_dtypes(include='number')

        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=0.2, random_state=42, stratify=y
        )

        self.model.fit(X_train, y_train)
        predictions = self.model.predict(X_test)

        acc = accuracy_score(y_test, predictions)
        prec = precision_score(y_test, predictions, average='weighted', zero_division=0)
        rec = recall_score(y_test, predictions, average='weighted', zero_division=0)
        f1 = f1_score(y_test, predictions, average='weighted', zero_division=0)

        self.cm = confusion_matrix(y_test, predictions)

        self.text_area.insert(tk.END, "\n--- Classification Metrics ---\n")
        self.text_area.insert(tk.END, f"Accuracy: {acc:.4f}\n")
        self.text_area.insert(tk.END, f"Precision: {prec:.4f}\n")
        self.text_area.insert(tk.END, f"Recall: {rec:.4f}\n")
        self.text_area.insert(tk.END, f"F1 Score: {f1:.4f}\n")

    def plot_confusion(self):
        if not hasattr(self, 'cm'):
            messagebox.showinfo("Info", "Train the model first")
            return

        plt.figure()
        plt.imshow(self.cm)
        plt.title("Confusion Matrix")
        plt.xlabel("Predicted")
        plt.ylabel("Actual")
        plt.colorbar()

        for i in range(self.cm.shape[0]):
            for j in range(self.cm.shape[1]):
                plt.text(j, i, self.cm[i, j], ha="center", va="center")

        plt.show()


if __name__ == "__main__":
    root = tk.Tk()
    app = ClassificationApp(root)
    root.mainloop()
