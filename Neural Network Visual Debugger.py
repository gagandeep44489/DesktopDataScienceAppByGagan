import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.neural_network import MLPClassifier
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import accuracy_score

class NeuralNetworkDebugger:
    def __init__(self, root):
        self.root = root
        self.root.title("Neural Network Visual Debugger")
        self.root.geometry("750x550")

        self.data = None

        tk.Button(root, text="Load CSV Dataset", command=self.load_data).pack(pady=10)

        tk.Label(root, text="Select Target Column").pack()
        self.target_combo = ttk.Combobox(root)
        self.target_combo.pack(pady=5)

        tk.Button(root, text="Train Neural Network", command=self.train_model).pack(pady=10)

        self.output = tk.Text(root, height=12)
        self.output.pack(pady=10)

        tk.Button(root, text="Visualize Training Metrics", command=self.visualize_metrics).pack(pady=10)

    def load_data(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path:
            self.data = pd.read_csv(path)
            self.target_combo["values"] = list(self.data.columns)
            messagebox.showinfo("Success", "Dataset Loaded")

    def train_model(self):
        target = self.target_combo.get()
        if not target:
            messagebox.showerror("Error", "Select target column")
            return

        X = self.data.drop(columns=[target]).select_dtypes(include="number")
        y = self.data[target]

        scaler = StandardScaler()
        X = scaler.fit_transform(X)

        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        self.model = MLPClassifier(
            hidden_layer_sizes=(50, 30),
            max_iter=1,
            warm_start=True,
            random_state=42
        )

        self.train_loss = []
        self.train_acc = []

        for epoch in range(50):
            self.model.fit(X_train, y_train)
            self.train_loss.append(self.model.loss_)

            preds = self.model.predict(X_train)
            acc = accuracy_score(y_train, preds)
            self.train_acc.append(acc)

        test_preds = self.model.predict(X_test)
        test_acc = accuracy_score(y_test, test_preds)

        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, f"Training Complete\n\n")
        self.output.insert(tk.END, f"Final Training Accuracy: {self.train_acc[-1]:.4f}\n")
        self.output.insert(tk.END, f"Test Accuracy: {test_acc:.4f}\n")

        if self.train_acc[-1] > test_acc + 0.1:
            self.output.insert(tk.END, "\n⚠ Possible Overfitting Detected")
        elif self.train_acc[-1] < 0.6:
            self.output.insert(tk.END, "\n⚠ Possible Underfitting Detected")
        else:
            self.output.insert(tk.END, "\n✔ Model Training Looks Healthy")

    def visualize_metrics(self):
        if not hasattr(self, "train_loss"):
            messagebox.showerror("Error", "Train model first")
            return

        plt.figure()
        plt.plot(self.train_loss, label="Loss")
        plt.plot(self.train_acc, label="Accuracy")
        plt.xlabel("Epoch")
        plt.title("Neural Network Training Debug View")
        plt.legend()
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = NeuralNetworkDebugger(root)
    root.mainloop()
