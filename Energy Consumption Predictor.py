import tkinter as tk
from tkinter import messagebox
import numpy as np
from sklearn.linear_model import LinearRegression

# -------------------------
# Train Dummy ML Model
# -------------------------
# Features: [Temperature, Occupants, Usage Hours]
X = np.array([
    [20, 2, 5],
    [25, 3, 6],
    [30, 4, 8],
    [35, 5, 10],
    [22, 2, 4],
    [28, 3, 7],
    [32, 4, 9],
    [18, 1, 3]
])

# Target: Energy Consumption (kWh)
y = np.array([12, 18, 25, 35, 10, 22, 30, 8])

model = LinearRegression()
model.fit(X, y)

# -------------------------
# Prediction Function
# -------------------------
def predict_energy():
    try:
        temp = float(entry_temp.get())
        occupants = int(entry_occupants.get())
        hours = float(entry_hours.get())

        prediction = model.predict([[temp, occupants, hours]])
        result_label.config(
            text=f"Predicted Energy Consumption: {prediction[0]:.2f} kWh"
        )
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numeric values")

# -------------------------
# GUI Setup
# -------------------------
root = tk.Tk()
root.title("Energy Consumption Predictor")
root.geometry("420x350")
root.resizable(False, False)

title = tk.Label(
    root,
    text="Energy Consumption Predictor",
    font=("Arial", 16, "bold")
)
title.pack(pady=10)

frame = tk.Frame(root)
frame.pack(pady=10)

tk.Label(frame, text="Temperature (°C):", font=("Arial", 11)).grid(row=0, column=0, pady=8, sticky="w")
entry_temp = tk.Entry(frame, width=20)
entry_temp.grid(row=0, column=1)

tk.Label(frame, text="Number of Occupants:", font=("Arial", 11)).grid(row=1, column=0, pady=8, sticky="w")
entry_occupants = tk.Entry(frame, width=20)
entry_occupants.grid(row=1, column=1)

tk.Label(frame, text="Usage Hours per Day:", font=("Arial", 11)).grid(row=2, column=0, pady=8, sticky="w")
entry_hours = tk.Entry(frame, width=20)
entry_hours.grid(row=2, column=1)

predict_btn = tk.Button(
    root,
    text="Predict Energy Consumption",
    font=("Arial", 12),
    bg="#2ecc71",
    fg="white",
    command=predict_energy
)
predict_btn.pack(pady=15)

result_label = tk.Label(
    root,
    text="Predicted Energy Consumption: -- kWh",
    font=("Arial", 12, "bold"),
    fg="blue"
)
result_label.pack(pady=10)

footer = tk.Label(
    root,
    text="Desktop ML App | Python + Tkinter",
    font=("Arial", 9),
    fg="gray"
)
footer.pack(side="bottom", pady=5)

root.mainloop()
