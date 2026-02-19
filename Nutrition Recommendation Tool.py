import tkinter as tk
from tkinter import messagebox

def calculate_nutrition():
    try:
        age = int(entry_age.get())
        weight = float(entry_weight.get())
        height = float(entry_height.get())
        gender = gender_var.get()
        activity = activity_var.get()
        goal = goal_var.get()

        # BMI Calculation
        height_m = height / 100
        bmi = weight / (height_m ** 2)

        # BMR (Mifflin-St Jeor Equation)
        if gender == "Male":
            bmr = 10 * weight + 6.25 * height - 5 * age + 5
        else:
            bmr = 10 * weight + 6.25 * height - 5 * age - 161

        # Activity Multiplier
        activity_multipliers = {
            "Sedentary": 1.2,
            "Light": 1.375,
            "Moderate": 1.55,
            "Active": 1.725,
            "Very Active": 1.9
        }

        tdee = bmr * activity_multipliers[activity]

        # Adjust Calories Based on Goal
        if goal == "Weight Loss":
            calories = tdee - 500
        elif goal == "Weight Gain":
            calories = tdee + 500
        else:
            calories = tdee

        # Macronutrient Distribution (Balanced Diet)
        protein = (0.30 * calories) / 4
        carbs = (0.40 * calories) / 4
        fats = (0.30 * calories) / 9

        result = f"""
BMI: {bmi:.2f}

Basal Metabolic Rate (BMR): {bmr:.2f} kcal/day
Total Daily Energy Expenditure (TDEE): {tdee:.2f} kcal/day

Recommended Daily Calories: {calories:.2f} kcal

Macronutrient Recommendation:
Protein: {protein:.2f} grams
Carbohydrates: {carbs:.2f} grams
Fats: {fats:.2f} grams
        """

        text_result.delete(1.0, tk.END)
        text_result.insert(tk.END, result)

    except Exception:
        messagebox.showerror("Error", "Please enter valid input values.")


# GUI Setup
root = tk.Tk()
root.title("Nutrition Recommendation Tool")
root.geometry("600x650")

tk.Label(root, text="Nutrition Recommendation Tool",
         font=("Arial", 16, "bold")).pack(pady=10)

tk.Label(root, text="Age:").pack()
entry_age = tk.Entry(root)
entry_age.pack()

tk.Label(root, text="Weight (kg):").pack()
entry_weight = tk.Entry(root)
entry_weight.pack()

tk.Label(root, text="Height (cm):").pack()
entry_height = tk.Entry(root)
entry_height.pack()

tk.Label(root, text="Gender:").pack()
gender_var = tk.StringVar(value="Male")
tk.OptionMenu(root, gender_var, "Male", "Female").pack()

tk.Label(root, text="Activity Level:").pack()
activity_var = tk.StringVar(value="Sedentary")
tk.OptionMenu(root, activity_var,
              "Sedentary", "Light", "Moderate", "Active", "Very Active").pack()

tk.Label(root, text="Goal:").pack()
goal_var = tk.StringVar(value="Maintain Weight")
tk.OptionMenu(root, goal_var,
              "Weight Loss", "Maintain Weight", "Weight Gain").pack()

tk.Button(root, text="Generate Recommendation",
          command=calculate_nutrition,
          bg="green", fg="white").pack(pady=15)

text_result = tk.Text(root, height=15, width=70)
text_result.pack(pady=10)

root.mainloop()
