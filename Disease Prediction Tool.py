import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
import tkinter as tk
from tkinter import messagebox
import pickle
import os

# ------------------------
# 1. Dataset (You can replace it with CSV)
# ------------------------
data = pd.DataFrame({
    "fever": [1, 0, 1, 0, 1],
    "cough": [1, 1, 0, 0, 1],
    "headache": [0, 1, 1, 0, 1],
    "fatigue": [1, 1, 1, 0, 1],
    "vomiting": [0, 0, 1, 0, 1],
    "disease": ["Flu", "Cold", "Dengue", "Healthy", "Malaria"]
})

# ------------------------
# 2. Train ML model
# ------------------------
X = data.drop("disease", axis=1)
y = data["disease"]

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
model = RandomForestClassifier()
model.fit(X_train, y_train)

# Save model to disk
model_file = "model.pkl"
with open(model_file, "wb") as f:
    pickle.dump(model, f)

# ------------------------
# 3. Create GUI
# ------------------------
root = tk.Tk()
root.title("Disease Prediction Tool")
root.geometry("400x400")

tk.Label(root, text="Select Symptoms", font=("Arial", 14)).pack(pady=10)

# Symptoms list
symptoms = ["fever", "cough", "headache", "fatigue", "vomiting"]
vars = []

for symptom in symptoms:
    var = tk.IntVar()
    tk.Checkbutton(root, text=symptom.capitalize(), variable=var).pack(anchor="w")
    vars.append(var)

def predict():
    input_data = [var.get() for var in vars]
    prediction = model.predict([input_data])
    messagebox.showinfo("Prediction Result", f"Predicted Disease: {prediction[0]}")

tk.Button(root, text="Predict", command=predict, bg="green", fg="white").pack(pady=20)

# ------------------------
# 4. Run the app
# ------------------------
root.mainloop()
