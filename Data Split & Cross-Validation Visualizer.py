import tkinter as tk
from tkinter import ttk, messagebox
import numpy as np
import matplotlib.pyplot as plt
from sklearn.model_selection import KFold

class DataSplitVisualizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Split & Cross-Validation Visualizer")
        self.root.geometry("600x450")

        self.create_widgets()

    def create_widgets(self):
        title = tk.Label(self.root,
                         text="Data Split & Cross-Validation Visualizer",
                         font=("Arial", 14, "bold"))
        title.pack(pady=10)

        frame = tk.Frame(self.root)
        frame.pack(pady=10)

        tk.Label(frame, text="Dataset Size:").grid(row=0, column=0, padx=5)
        self.size_entry = tk.Entry(frame, width=10)
        self.size_entry.insert(0, "100")
        self.size_entry.grid(row=0, column=1)

        tk.Label(frame, text="Train %:").grid(row=1, column=0)
        self.train_entry = tk.Entry(frame, width=10)
        self.train_entry.insert(0, "70")
        self.train_entry.grid(row=1, column=1)

        tk.Label(frame, text="Validation %:").grid(row=2, column=0)
        self.val_entry = tk.Entry(frame, width=10)
        self.val_entry.insert(0, "15")
        self.val_entry.grid(row=2, column=1)

        tk.Label(frame, text="Test %:").grid(row=3, column=0)
        self.test_entry = tk.Entry(frame, width=10)
        self.test_entry.insert(0, "15")
        self.test_entry.grid(row=3, column=1)

        tk.Label(frame, text="K-Folds:").grid(row=4, column=0)
        self.k_entry = tk.Entry(frame, width=10)
        self.k_entry.insert(0, "5")
        self.k_entry.grid(row=4, column=1)

        btn = tk.Button(self.root, text="Visualize",
                        command=self.visualize)
        btn.pack(pady=15)

    def visualize(self):
        try:
            n = int(self.size_entry.get())
            train_p = int(self.train_entry.get())
            val_p = int(self.val_entry.get())
            test_p = int(self.test_entry.get())
            k = int(self.k_entry.get())

            if train_p + val_p + test_p != 100:
                raise ValueError("Split percentages must total 100")

            indices = np.arange(n)

            # Train/Val/Test split visualization
            split_points = [
                int(n * train_p / 100),
                int(n * (train_p + val_p) / 100)
            ]

            plt.figure(figsize=(10, 5))

            plt.subplot(2, 1, 1)
            plt.title("Train / Validation / Test Split")
            plt.scatter(indices[:split_points[0]], np.zeros(split_points[0]),
                        label="Train")
            plt.scatter(indices[split_points[0]:split_points[1]],
                        np.zeros(split_points[1] - split_points[0]),
                        label="Validation")
            plt.scatter(indices[split_points[1]:],
                        np.zeros(n - split_points[1]),
                        label="Test")
            plt.yticks([])
            plt.legend()

            # K-Fold visualization
            plt.subplot(2, 1, 2)
            plt.title(f"{k}-Fold Cross Validation")

            kf = KFold(n_splits=k, shuffle=True, random_state=42)
            y_pos = 0

            for fold, (train_idx, test_idx) in enumerate(kf.split(indices)):
                plt.scatter(train_idx, np.ones(len(train_idx)) * y_pos,
                            label=f"Train Fold {fold+1}")
                plt.scatter(test_idx, np.ones(len(test_idx)) * y_pos,
                            label=f"Test Fold {fold+1}")
                y_pos += 1

            plt.yticks([])
            plt.tight_layout()
            plt.show()

        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = DataSplitVisualizer(root)
    root.mainloop()
