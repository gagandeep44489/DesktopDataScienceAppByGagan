import tkinter as tk
from tkinter import messagebox
import numpy as np
import pandas as pd
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import train_test_split

# --------------------------
# Sample Dataset (Demo Data)
# --------------------------

data = {
    "age": [45, 50, 37, 62, 41, 56, 48, 52, 60, 39],
    "cholesterol": [230, 250, 180, 270, 210, 260, 240, 255, 280, 190],
    "blood_pressure": [130, 145, 120, 160, 135, 150, 140, 155, 165, 125],
    "max_heart_rate": [150, 140, 170, 130, 160, 135, 145, 138, 125, 175],
    "target": [1, 1, 0, 1, 0, 1, 0, 1, 1, 0]
}

df = pd.DataFrame(data)

X = df.drop("target", axis=1)
y = df["target"]

model = LogisticRegression()
model.fit(X, y)

# --------------------------
# Prediction Function
# --------------------------

def predict_risk():
    try:
        age = float(entry_age.get())
        chol = float(entry_chol.get())
        bp = float(entry_bp.get())
        hr = float(entry_hr.get())

        input_data = np.array([[age, chol, bp, hr]])
        probability = model.predict_proba(input_data)[0][1]

        if probability > 0.6:
            risk_level = "High Risk"
        elif probability > 0.4:
            risk_level = "Moderate Risk"
        else:
            risk_level = "Low Risk"

        result = f"""
Predicted Risk Probability: {probability:.2f}

Risk Category: {risk_level}
        """

        text_result.delete(1.0, tk.END)
        text_result.insert(tk.END, result)

    except:
        messagebox.showerror("Error", "Please enter valid numeric values.")


# --------------------------
# GUI Setup
# --------------------------

root = tk.Tk()
root.title("Heart Disease Risk Predictor")
root.geometry("600x600")

tk.Label(root, text="Heart Disease Risk Predictor",
         font=("Arial", 16, "bold")).pack(pady=10)

tk.Label(root, text="Age:").pack()
entry_age = tk.Entry(root)
entry_age.pack()

tk.Label(root, text="Cholesterol Level:").pack()
entry_chol = tk.Entry(root)
entry_chol.pack()

tk.Label(root, text="Blood Pressure:").pack()
entry_bp = tk.Entry(root)
entry_bp.pack()

tk.Label(root, text="Max Heart Rate:").pack()
entry_hr = tk.Entry(root)
entry_hr.pack()

tk.Button(root, text="Predict Risk",
          command=predict_risk,
          bg="red", fg="white").pack(pady=15)

text_result = tk.Text(root, height=10, width=70)
text_result.pack(pady=10)

root.mainloop()
