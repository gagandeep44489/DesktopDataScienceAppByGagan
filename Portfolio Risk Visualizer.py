import tkinter as tk
from tkinter import messagebox
import numpy as np
import matplotlib.pyplot as plt

def calculate_portfolio():
    try:
        # Get inputs
        weights = np.array([
            float(entry_w1.get()),
            float(entry_w2.get()),
            float(entry_w3.get())
        ])

        returns = np.array([
            float(entry_r1.get()),
            float(entry_r2.get()),
            float(entry_r3.get())
        ])

        risks = np.array([
            float(entry_s1.get()),
            float(entry_s2.get()),
            float(entry_s3.get())
        ])

        # Normalize weights
        weights = weights / np.sum(weights)

        # Portfolio return
        portfolio_return = np.sum(weights * returns)

        # Portfolio risk (simplified)
        portfolio_risk = np.sqrt(np.sum((weights ** 2) * (risks ** 2)))

        result_label.config(
            text=f"Expected Return: {portfolio_return:.2f}%\nPortfolio Risk: {portfolio_risk:.2f}%"
        )

        # Plot risk-return graph
        plt.figure()
        plt.scatter(risks, returns)
        plt.scatter(portfolio_risk, portfolio_return)
        plt.xlabel("Risk (Std Dev %)")
        plt.ylabel("Expected Return (%)")
        plt.title("Portfolio Risk-Return Visualization")
        plt.show()

    except:
        messagebox.showerror("Error", "Please enter valid numeric values.")

# GUI Setup
root = tk.Tk()
root.title("Portfolio Risk Visualizer")
root.geometry("500x650")

tk.Label(root, text="Portfolio Risk Visualizer", font=("Arial", 16)).pack(pady=10)

# Asset 1
tk.Label(root, text="Asset 1 Weight").pack()
entry_w1 = tk.Entry(root)
entry_w1.pack()

tk.Label(root, text="Asset 1 Expected Return (%)").pack()
entry_r1 = tk.Entry(root)
entry_r1.pack()

tk.Label(root, text="Asset 1 Risk (%)").pack()
entry_s1 = tk.Entry(root)
entry_s1.pack()

# Asset 2
tk.Label(root, text="Asset 2 Weight").pack()
entry_w2 = tk.Entry(root)
entry_w2.pack()

tk.Label(root, text="Asset 2 Expected Return (%)").pack()
entry_r2 = tk.Entry(root)
entry_r2.pack()

tk.Label(root, text="Asset 2 Risk (%)").pack()
entry_s2 = tk.Entry(root)
entry_s2.pack()

# Asset 3
tk.Label(root, text="Asset 3 Weight").pack()
entry_w3 = tk.Entry(root)
entry_w3.pack()

tk.Label(root, text="Asset 3 Expected Return (%)").pack()
entry_r3 = tk.Entry(root)
entry_r3.pack()

tk.Label(root, text="Asset 3 Risk (%)").pack()
entry_s3 = tk.Entry(root)
entry_s3.pack()

tk.Button(root, text="Calculate Portfolio", command=calculate_portfolio).pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 14))
result_label.pack(pady=10)

root.mainloop()
