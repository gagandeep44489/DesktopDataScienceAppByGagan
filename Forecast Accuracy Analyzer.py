import tkinter as tk
from tkinter import messagebox
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from sklearn.metrics import mean_absolute_error, mean_squared_error

# -------------------------
# Sample Actual & Forecast Data
# -------------------------
actual = np.array([100, 120, 130, 150, 170, 160, 180])
forecast = np.array([110, 115, 140, 145, 165, 170, 175])

# -------------------------
# Accuracy Calculation
# -------------------------
def analyze_accuracy():
    try:
        mae = mean_absolute_error(actual, forecast)
        rmse = np.sqrt(mean_squared_error(actual, forecast))
        mape = np.mean(np.abs((actual - forecast) / actual)) * 100

        result_label.config(
            text=f"MAE: {mae:.2f}   RMSE: {rmse:.2f}   MAPE: {mape:.2f}%"
        )

        plot_results()

    except Exception as e:
        messagebox.showerror("Error", str(e))

# -------------------------
# Plot Function
# -------------------------
def plot_results():
    fig.clear()
    ax = fig.add_subplot(111)
    ax.plot(actual, marker="o", label="Actual")
    ax.plot(forecast, marker="x", label="Forecast")
    ax.set_title("Actual vs Forecast Comparison")
    ax.set_xlabel("Time")
    ax.set_ylabel("Value")
    ax.legend()
    canvas.draw()

# -------------------------
# GUI Setup
# -------------------------
root = tk.Tk()
root.title("Forecast Accuracy Analyzer")
root.geometry("850x600")
root.resizable(False, False)

title = tk.Label(
    root,
    text="Forecast Accuracy Analyzer",
    font=("Arial", 16, "bold")
)
title.pack(pady=10)

desc = tk.Label(
    root,
    text="Analyze prediction accuracy using MAE, RMSE, and MAPE",
    font=("Arial", 11),
    fg="gray"
)
desc.pack()

btn = tk.Button(
    root,
    text="Analyze Forecast Accuracy",
    font=("Arial", 12),
    bg="#2ecc71",
    fg="white",
    command=analyze_accuracy
)
btn.pack(pady=15)

result_label = tk.Label(
    root,
    text="MAE: --   RMSE: --   MAPE: --",
    font=("Arial", 12, "bold"),
    fg="blue"
)
result_label.pack(pady=10)

fig = plt.Figure(figsize=(8, 4), dpi=100)
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack(pady=15)

footer = tk.Label(
    root,
    text="Model Evaluation | Forecast Accuracy | Python Desktop App",
    font=("Arial", 9),
    fg="gray"
)
footer.pack(side="bottom", pady=5)

root.mainloop()
