import tkinter as tk
from tkinter import messagebox
import numpy as np
from sklearn.linear_model import LinearRegression

# Sample Training Data
# Features: [Hours Studied, Attendance %, Previous Score]
X = np.array([
    [2, 60, 50],
    [3, 65, 55],
    [4, 70, 60],
    [5, 75, 65],
    [6, 80, 70],
    [7, 85, 75],
    [8, 90, 80],
    [9, 95, 85],
    [10, 100, 90]
])

# Target: Final Exam Score
y = np.array([52, 58, 63, 68, 74, 79, 85, 90, 96])

# Train Model
model = LinearRegression()
model.fit(X, y)

# Prediction Function
def predict_score():
    try:
        hours = float(entry_hours.get())
        attendance = float(entry_attendance.get())
        previous = float(entry_previous.get())

        input_data = np.array([[hours, attendance, previous]])
        prediction = model.predict(input_data)

        result_label.config(text=f"Predicted Exam Score: {prediction[0]:.2f}")

    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numeric values!")

# GUI Setup
root = tk.Tk()
root.title("Exam Score Predictor")
root.geometry("400x350")
root.resizable(False, False)

tk.Label(root, text="Exam Score Predictor", font=("Arial", 16, "bold")).pack(pady=10)

tk.Label(root, text="Hours Studied:").pack()
entry_hours = tk.Entry(root)
entry_hours.pack(pady=5)

tk.Label(root, text="Attendance (%):").pack()
entry_attendance = tk.Entry(root)
entry_attendance.pack(pady=5)

tk.Label(root, text="Previous Exam Score:").pack()
entry_previous = tk.Entry(root)
entry_previous.pack(pady=5)

tk.Button(root, text="Predict Score", command=predict_score, bg="blue", fg="white").pack(pady=15)

result_label = tk.Label(root, text="", font=("Arial", 12, "bold"))
result_label.pack(pady=10)

root.mainloop()