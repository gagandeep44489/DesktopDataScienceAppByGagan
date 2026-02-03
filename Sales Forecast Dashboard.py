import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from sklearn.linear_model import LinearRegression

class SalesForecastDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Sales Forecast Dashboard")
        self.root.geometry("750x520")

        self.data = None
        self.model = LinearRegression()

        self.create_ui()

    def create_ui(self):
        tk.Button(self.root, text="Load Sales CSV", command=self.load_csv, width=25).pack(pady=10)
        tk.Button(self.root, text="Train Forecast Model", command=self.train_model, width=25).pack(pady=5)
        tk.Button(self.root, text="Predict Future Sales", command=self.predict_sales, width=25).pack(pady=5)
        tk.Button(self.root, text="Show Dashboard", command=self.show_dashboard, width=25).pack(pady=5)

        self.result_label = tk.Label(self.root, text="", font=("Arial", 12))
        self.result_label.pack(pady=15)

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return

        self.data = pd.read_csv(file_path)

        if "Sales" not in self.data.columns:
            messagebox.showerror("Error", "CSV must contain 'Sales' column")
            return

        self.data["Day"] = np.arange(len(self.data))
        messagebox.showinfo("Success", "Sales data loaded successfully")

    def train_model(self):
        if self.data is None:
            messagebox.showerror("Error", "Load sales data first")
            return

        X = self.data[["Day"]]
        y = self.data["Sales"]

        self.model.fit(X, y)
        messagebox.showinfo("Success", "Forecast model trained")

    def predict_sales(self):
        if self.data is None:
            messagebox.showerror("Error", "Train model first")
            return

        next_day = [[len(self.data)]]
        prediction = self.model.predict(next_day)[0]

        self.result_label.config(
            text=f"Predicted Next Period Sales: {prediction:.2f}"
        )

    def show_dashboard(self):
        if self.data is None:
            messagebox.showerror("Error", "Load data first")
            return

        future_day = len(self.data)
        future_sales = self.model.predict([[future_day]])[0]

        plt.figure()
        plt.plot(self.data["Day"], self.data["Sales"], label="Actual Sales")
        plt.scatter(future_day, future_sales, label="Forecast", marker="o")
        plt.title("Sales Forecast Dashboard")
        plt.xlabel("Time Period")
        plt.ylabel("Sales")
        plt.legend()
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesForecastDashboard(root)
    root.mainloop()
