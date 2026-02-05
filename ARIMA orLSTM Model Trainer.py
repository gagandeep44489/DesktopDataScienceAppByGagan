import tkinter as tk
from tkinter import messagebox
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from statsmodels.tsa.arima.model import ARIMA
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import LSTM, Dense
from sklearn.preprocessing import MinMaxScaler

# -------------------------
# Generate Sample Time Series
# -------------------------
def generate_series():
    t = np.arange(0, 100)
    return np.sin(t / 10) + np.random.normal(0, 0.1, len(t))

series = generate_series()

# -------------------------
# Train ARIMA
# -------------------------
def train_arima():
    try:
        model = ARIMA(series, order=(2, 1, 2))
        model_fit = model.fit()
        forecast = model_fit.forecast(steps=20)
        plot_results(series, forecast, "ARIMA Forecast")
    except Exception as e:
        messagebox.showerror("ARIMA Error", str(e))

# -------------------------
# Train LSTM
# -------------------------
def train_lstm():
    try:
        data = series.reshape(-1, 1)
        scaler = MinMaxScaler()
        data_scaled = scaler.fit_transform(data)

        X, y = [], []
        for i in range(10, len(data_scaled)):
            X.append(data_scaled[i-10:i])
            y.append(data_scaled[i])

        X, y = np.array(X), np.array(y)

        model = Sequential([
            LSTM(50, activation="relu", input_shape=(X.shape[1], 1)),
            Dense(1)
        ])
        model.compile(optimizer="adam", loss="mse")
        model.fit(X, y, epochs=10, verbose=0)

        last_seq = data_scaled[-10:].reshape(1, 10, 1)
        preds = []
        for _ in range(20):
            p = model.predict(last_seq, verbose=0)
            preds.append(p[0, 0])
            last_seq = np.append(last_seq[:, 1:, :], [[p]], axis=1)

        forecast = scaler.inverse_transform(np.array(preds).reshape(-1, 1)).flatten()
        plot_results(series, forecast, "LSTM Forecast")

    except Exception as e:
        messagebox.showerror("LSTM Error", str(e))

# -------------------------
# Plot Function
# -------------------------
def plot_results(actual, forecast, title):
    fig.clear()
    ax = fig.add_subplot(111)
    ax.plot(actual, label="Actual")
    ax.plot(range(len(actual), len(actual)+len(forecast)), forecast, label="Forecast")
    ax.set_title(title)
    ax.legend()
    canvas.draw()

# -------------------------
# GUI Setup
# -------------------------
root = tk.Tk()
root.title("ARIMA / LSTM Model Trainer")
root.geometry("850x600")
root.resizable(False, False)

title = tk.Label(
    root,
    text="ARIMA / LSTM Time Series Model Trainer",
    font=("Arial", 16, "bold")
)
title.pack(pady=10)

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

tk.Button(
    btn_frame,
    text="Train ARIMA Model",
    font=("Arial", 11),
    bg="#3498db",
    fg="white",
    width=18,
    command=train_arima
).grid(row=0, column=0, padx=10)

tk.Button(
    btn_frame,
    text="Train LSTM Model",
    font=("Arial", 11),
    bg="#2ecc71",
    fg="white",
    width=18,
    command=train_lstm
).grid(row=0, column=1, padx=10)

fig = plt.Figure(figsize=(8, 4), dpi=100)
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack(pady=20)

footer = tk.Label(
    root,
    text="Time Series Forecasting | ARIMA vs LSTM | Python Desktop App",
    font=("Arial", 9),
    fg="gray"
)
footer.pack(side="bottom", pady=5)

root.mainloop()
