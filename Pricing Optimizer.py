import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression

class PricingOptimizer:

    def __init__(self, root):
        self.root = root
        self.root.title("Pricing Optimizer Tool")
        self.root.geometry("500x400")

        self.data = None
        self.model = None

        tk.Button(root, text="Load CSV Data", command=self.load_data).pack(pady=10)
        tk.Button(root, text="Train Model", command=self.train_model).pack(pady=10)
        tk.Button(root, text="Optimize Price", command=self.optimize_price).pack(pady=10)
        tk.Button(root, text="Plot Revenue Curve", command=self.plot_curve).pack(pady=10)

        self.result_label = tk.Label(root, text="", font=("Arial", 12))
        self.result_label.pack(pady=20)

    def load_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.data = pd.read_csv(file_path)
            messagebox.showinfo("Success", "Data Loaded Successfully")

    def train_model(self):
        if self.data is None:
            messagebox.showerror("Error", "Load data first")
            return

        X = self.data[['Price']]
        y = self.data['Quantity']

        self.model = LinearRegression()
        self.model.fit(X, y)

        messagebox.showinfo("Success", "Demand Model Trained")

    def optimize_price(self):
        if self.model is None:
            messagebox.showerror("Error", "Train model first")
            return

        a = self.model.intercept_
        b = -self.model.coef_[0]

        optimal_price = a / (2 * b)

        self.result_label.config(
            text=f"Optimal Revenue-Maximizing Price: {round(optimal_price,2)}"
        )

    def plot_curve(self):
        if self.model is None:
            messagebox.showerror("Error", "Train model first")
            return

        prices = np.linspace(min(self.data['Price']),
                             max(self.data['Price']), 100)

        demand = self.model.predict(prices.reshape(-1,1))
        revenue = prices * demand

        plt.figure()
        plt.plot(prices, revenue)
        plt.xlabel("Price")
        plt.ylabel("Revenue")
        plt.title("Revenue Curve")
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = PricingOptimizer(root)
    root.mainloop()
