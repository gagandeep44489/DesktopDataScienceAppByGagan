import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
import numpy as np

class StockPricePredictor:
    def __init__(self, root):
        self.root = root
        self.root.title("Stock Price Predictor")
        self.root.geometry("700x500")

        self.data = None
        self.model = LinearRegression()

        self.create_ui()

    def create_ui(self):
        tk.Button(self.root, text="Load Stock CSV", command=self.load_csv).pack(pady=10)
        tk.Button(self.root, text="Train Model", command=self.train_model).pack(pady=5)
        tk.Button(self.root, text="Predict Next Price", command=self.predict_price).pack(pady=5)

        self.result_label = tk.Label(self.root, text="", font=("Arial", 12))
        self.result_label.pack(pady=10)

        tk.Button(self.root, text="Show Price Chart", command=self.plot_prices).pack(pady=5)

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return

        self.data = pd.read_csv(file_path)
        if "Close" not in self.data.columns:
            messagebox.showerror("Error", "CSV must contain 'Close' column")
            return

        messagebox.showinfo("Success", "Stock data loaded successfully")

    def train_model(self):
        if self.data is None:
            messagebox.showerror("Error", "Load stock data first")
            return

        self.data["Day"] = np.arange(len(self.data))
        X = self.data[["Day"]]
        y = self.data["Close"]

        self.model.fit(X, y)
        messagebox.showinfo("Success", "Model trained successfully")

    def predict_price(self):
        if self.data is None:
            messagebox.showerror("Error", "Train model first")
            return

        next_day = [[len(self.data)]]
        prediction = self.model.predict(next_day)[0]
        self.result_label.config(text=f"Predicted Next Close Price: {prediction:.2f}")

    def plot_prices(self):
        if self.data is None:
            messagebox.showerror("Error", "Load data first")
            return

        plt.figure()
        plt.plot(self.data["Close"])
        plt.title("Historical Closing Prices")
        plt.xlabel("Days")
        plt.ylabel("Price")
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = StockPricePredictor(root)
    root.mainloop()
