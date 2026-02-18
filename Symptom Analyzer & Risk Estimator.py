import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
import tkinter as tk
from tkinter import messagebox
import pickle
import numpy as np

# ------------------------
# 1. Sample Dataset
# ------------------------
# Symptoms encoded as 0 (absent) / 1 (present)
data = pd.DataFrame({
    "fever": [1, 0, 1, 0, 1, 0],
    "cough": [1, 1, 0, 0, 1, 1],
    "headache": [0, 1, 1, 0, 1, 1],
    "fatigue": [1, 1, 1, 0, 1, 0],
    "vomiting": [0, 0, 1, 0, 1, 0],
    "disease": ["Flu", "Cold", "Dengue", "Healthy", "Malaria", "Healthy"]
})

# ------------------------
# 2. Train ML model
# ------------------------
X = data.drop("disease", axis=1)
y = data["disease"]

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
model = RandomForestClassifier()
model.fit(X_train, y_train)

# Save model to file
with open("risk_model.pkl", "wb") as f:
    pickle.dump(model, f)

# ------------------------
# 3. GUI
# ------------------------
root = tk.Tk()
root.title("Symptom Analyzer & Risk Estimator")
root.geometry("450x450")

tk.Label(root, text="Select Symptoms You Have", font=("Arial", 14)).pack(pady=10)

# Symptoms list
symptoms = ["fever", "cough", "headache", "fatigue", "vomiting"]
vars = []

for symptom in symptoms:
    var = tk.IntVar()
    tk.Checkbutton(root, text=symptom.capitalize(), variable=var).pack(anchor="w")
    vars.append(var)

def analyze_risk():
    input_data = [var.get() for var in vars]
    
    # Predict disease
    prediction = model.predict([input_data])[0]
    
    # Calculate risk score: simple % based on number of symptoms present
    total_symptoms = len(input_data)
    symptoms_present = sum(input_data)
    risk_score = (symptoms_present / total_symptoms) * 100
    
    messagebox.showinfo(
        "Analysis Result",
        f"Predicted Disease: {prediction}\nEstimated Risk Score: {risk_score:.1f}%"
    )

tk.Button(root, text="Analyze & Estimate Risk", command=analyze_risk, bg="blue", fg="white").pack(pady=20)

root.mainloop()
