"""
Regression Model Trainer & Visualizer
Python Desktop Application using Tkinter, Pandas, Scikit-learn, Matplotlib

Features:
- Load CSV dataset
- Select target variable
- Train Linear Regression model
- Display model metrics
- Visualize regression results
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score, mean_squared_error

class RegressionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Regression Model Trainer & Visualizer")
        self.root.geometry("900x650")

        self.data = None
        self.model = LinearRegression()

        title = tk.Label(root, text="Regression Model Trainer & Visualizer", font=("Arial", 16, "bold"))
        title.pack(pady=10)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        load_btn = tk.Button(btn_frame, text="Load CSV Dataset", width=20, command=self.load_data)
        load_btn.grid(row=0, column=0, padx=10)

        train_btn = tk.Button(btn_frame, text="Train Model", width=20, command=self.train_model)
        train_btn.grid(row=0, column=1, padx=10)

        plot_btn = tk.Button(btn_frame, text="Visualize Regression", width=20, command=self.plot_regression)
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

        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
        self.model.fit(X_train, y_train)

        predictions = self.model.predict(X_test)
        r2 = r2_score(y_test, predictions)
        mse = mean_squared_error(y_test, predictions)

        self.text_area.insert(tk.END, "\n--- Model Training Results ---\n")
        self.text_area.insert(tk.END, f"R² Score: {r2:.4f}\n")
        self.text_area.insert(tk.END, f"Mean Squared Error: {mse:.4f}\n")

    def plot_regression(self):
        if self.data is None:
            return

        target = self.target_var.get()
        X = self.data.drop(columns=[target]).select_dtypes(include='number')
        y = self.data[target]

        if X.shape[1] != 1:
            messagebox.showinfo("Info", "Visualization works best with one numeric feature")
            return

        plt.figure()
        plt.scatter(X, y)
        plt.plot(X, self.model.predict(X))
        plt.xlabel(X.columns[0])
        plt.ylabel(target)
        plt.title("Regression Visualization")
        plt.show()


if __name__ == "__main__":
    root = tk.Tk()
    app = RegressionApp(root)
    root.mainloop()
