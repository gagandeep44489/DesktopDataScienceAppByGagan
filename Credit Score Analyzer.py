import tkinter as tk
from tkinter import messagebox

def calculate_score():
    try:
        age = int(entry_age.get())
        income = float(entry_income.get())
        loans = int(entry_loans.get())
        utilization = float(entry_utilization.get())
        payment = payment_var.get()
        cards = int(entry_cards.get())

        score = 300  # Base score

        # Income contribution
        score += min(income / 100000 * 100, 100)

        # Credit utilization (lower is better)
        if utilization < 30:
            score += 150
        elif utilization < 50:
            score += 100
        else:
            score += 50

        # Payment history
        if payment == "Good":
            score += 200
        elif payment == "Average":
            score += 120
        else:
            score += 50

        # Existing loans (less is better)
        score += max(100 - loans * 10, 0)

        # Credit cards
        if 2 <= cards <= 5:
            score += 80
        else:
            score += 40

        score = min(int(score), 900)

        # Risk category
        if score >= 750:
            category = "Excellent"
        elif score >= 650:
            category = "Good"
        elif score >= 550:
            category = "Fair"
        else:
            category = "Poor"

        result_label.config(text=f"Credit Score: {score}\nCategory: {category}")

    except:
        messagebox.showerror("Error", "Please enter valid data.")

# GUI Setup
root = tk.Tk()
root.title("Credit Score Analyzer")
root.geometry("400x500")

tk.Label(root, text="Credit Score Analyzer", font=("Arial", 16)).pack(pady=10)

tk.Label(root, text="Age").pack()
entry_age = tk.Entry(root)
entry_age.pack()

tk.Label(root, text="Annual Income").pack()
entry_income = tk.Entry(root)
entry_income.pack()

tk.Label(root, text="Existing Loans").pack()
entry_loans = tk.Entry(root)
entry_loans.pack()

tk.Label(root, text="Credit Utilization (%)").pack()
entry_utilization = tk.Entry(root)
entry_utilization.pack()

tk.Label(root, text="Payment History").pack()
payment_var = tk.StringVar(value="Good")
tk.OptionMenu(root, payment_var, "Good", "Average", "Poor").pack()

tk.Label(root, text="Number of Credit Cards").pack()
entry_cards = tk.Entry(root)
entry_cards.pack()

tk.Button(root, text="Calculate Score", command=calculate_score).pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 14))
result_label.pack(pady=10)

root.mainloop()
