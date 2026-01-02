import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt

class ModelComparatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Model Performance Comparator")
        self.root.geometry("700x500")

        self.data = []

        self.create_widgets()

    def create_widgets(self):
        title = tk.Label(self.root, text="Model Performance Comparator",
                         font=("Arial", 16, "bold"))
        title.pack(pady=10)

        frame = tk.Frame(self.root)
        frame.pack(pady=10)

        labels = ["Model Name", "Accuracy", "Precision", "Recall", "F1 Score"]

        self.entries = {}
        for i, label in enumerate(labels):
            tk.Label(frame, text=label).grid(row=0, column=i, padx=5)
            entry = tk.Entry(frame, width=15)
            entry.grid(row=1, column=i, padx=5)
            self.entries[label] = entry

        add_btn = tk.Button(self.root, text="Add Model",
                            command=self.add_model)
        add_btn.pack(pady=10)

        self.tree = ttk.Treeview(self.root,
                                 columns=("Model", "Acc", "Prec", "Rec", "F1"),
                                 show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
        self.tree.pack(expand=True, fill="both", pady=10)

        compare_btn = tk.Button(self.root, text="Compare Models",
                                command=self.compare_models)
        compare_btn.pack(pady=10)

    def add_model(self):
        try:
            model_data = {
                "Model": self.entries["Model Name"].get(),
                "Accuracy": float(self.entries["Accuracy"].get()),
                "Precision": float(self.entries["Precision"].get()),
                "Recall": float(self.entries["Recall"].get()),
                "F1 Score": float(self.entries["F1 Score"].get())
            }

            self.data.append(model_data)
            self.tree.insert("", "end", values=list(model_data.values()))

            for entry in self.entries.values():
                entry.delete(0, tk.END)

        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numeric values")

    def compare_models(self):
        if not self.data:
            messagebox.showwarning("No Data", "Add at least one model")
            return

        df = pd.DataFrame(self.data)
        df.set_index("Model", inplace=True)

        df.plot(kind="bar", figsize=(10, 5))
        plt.title("Model Performance Comparison")
        plt.ylabel("Score")
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()

        best_model = df["F1 Score"].idxmax()
        messagebox.showinfo("Best Model",
                            f"Best Model based on F1 Score:\n{best_model}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ModelComparatorApp(root)
    root.mainloop()
