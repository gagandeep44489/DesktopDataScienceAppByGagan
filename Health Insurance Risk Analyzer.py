import tkinter as tk
from tkinter import messagebox
import matplotlib.pyplot as plt

# ==============================
# RISK CALCULATION FUNCTION
# ==============================

def calculate_risk():
    try:
        age = int(age_var.get())
        bmi = float(bmi_var.get())
        smoker = smoker_var.get()
        exercise = int(exercise_var.get())
        bp = int(bp_var.get())
        cholesterol = int(chol_var.get())
        condition = condition_var.get()

        risk = 0

        # Age risk
        risk += (age / 100) * 20

        # BMI risk
        risk += (bmi / 40) * 20

        # Smoking risk
        if smoker == "Yes":
            risk += 20

        # Pre-existing condition
        if condition == "Yes":
            risk += 15

        # Blood pressure
        risk += (bp / 180) * 10

        # Cholesterol
        risk += (cholesterol / 300) * 10

        # Exercise (less exercise = more risk)
        risk += ((7 - exercise) / 7) * 5

        risk = min(risk, 100)

        if risk < 35:
            category = "Low Risk"
            multiplier = "1.0x Premium"
        elif risk < 65:
            category = "Moderate Risk"
            multiplier = "1.5x Premium"
        else:
            category = "High Risk"
            multiplier = "2.0x Premium"

        result_label.config(
            text=f"Risk Score: {risk:.2f}\nCategory: {category}\nEstimated Premium: {multiplier}"
        )

        show_chart(risk)

    except:
        messagebox.showerror("Error", "Please enter valid numeric inputs")

# ==============================
# VISUALIZATION
# ==============================

def show_chart(risk):
    plt.figure()
    plt.bar(["Risk Score"], [risk])
    plt.ylim(0, 100)
    plt.title("Insurance Risk Score")
    plt.ylabel("Score")
    plt.show()

# ==============================
# GUI SETUP
# ==============================

root = tk.Tk()
root.title("Health Insurance Risk Analyzer")
root.geometry("400x600")

tk.Label(root, text="Age").pack()
age_var = tk.StringVar()
tk.Entry(root, textvariable=age_var).pack()

tk.Label(root, text="BMI").pack()
bmi_var = tk.StringVar()
tk.Entry(root, textvariable=bmi_var).pack()

tk.Label(root, text="Smoker (Yes/No)").pack()
smoker_var = tk.StringVar()
tk.Entry(root, textvariable=smoker_var).pack()

tk.Label(root, text="Exercise Days per Week (0-7)").pack()
exercise_var = tk.StringVar()
tk.Entry(root, textvariable=exercise_var).pack()

tk.Label(root, text="Blood Pressure").pack()
bp_var = tk.StringVar()
tk.Entry(root, textvariable=bp_var).pack()

tk.Label(root, text="Cholesterol").pack()
chol_var = tk.StringVar()
tk.Entry(root, textvariable=chol_var).pack()

tk.Label(root, text="Pre-existing Condition (Yes/No)").pack()
condition_var = tk.StringVar()
tk.Entry(root, textvariable=condition_var).pack()

tk.Button(root, text="Calculate Risk", command=calculate_risk).pack(pady=15)

result_label = tk.Label(root, text="", fg="blue")
result_label.pack()

root.mainloop()