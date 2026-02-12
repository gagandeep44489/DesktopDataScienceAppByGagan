import tkinter as tk
from tkinter import messagebox
import numpy as np
from sklearn.linear_model import LinearRegression

# Sample training data
# [Income, Rent, Utilities, Groceries, Transport, Entertainment]
X = np.array([
    [50000, 10000, 3000, 5000, 2000, 3000],
    [60000, 15000, 4000, 6000, 3000, 4000],
    [40000, 8000, 2500, 4000, 1500, 2000],
    [70000, 20000, 5000, 7000, 4000, 5000],
    [55000, 12000, 3500, 5500, 2500, 3500]
])

# Total expense (target)
y = np.array([23000, 32000, 18000, 41000, 27000])

# Train model
model = LinearRegression()
model.fit(X, y)

def predict_expense():
    try:
        income = float(entry_income.get())
        rent = float(entry_rent.get())
        utilities = float(entry_utilities.get())
        groceries = float(entry_groceries.get())
        transport = float(entry_transport.get())
        entertainment = float(entry_entertainment.get())

        input_data = np.array([[income, rent, utilities, groceries, transport, entertainment]])
        prediction = model.predict(input_data)

        result_label.config(text=f"Predicted Monthly Expense: ₹{int(prediction[0])}")

    except:
        messagebox.showerror("Error", "Please enter valid numeric values.")

# GUI Setup
root = tk.Tk()
root.title("Expense Prediction Tool")
root.geometry("450x550")

tk.Label(root, text="Expense Prediction Tool", font=("Arial", 16)).pack(pady=10)

tk.Label(root, text="Monthly Income").pack()
entry_income = tk.Entry(root)
entry_income.pack()

tk.Label(root, text="Rent").pack()
entry_rent = tk.Entry(root)
entry_rent.pack()

tk.Label(root, text="Utilities").pack()
entry_utilities = tk.Entry(root)
entry_utilities.pack()

tk.Label(root, text="Groceries").pack()
entry_groceries = tk.Entry(root)
entry_groceries.pack()

tk.Label(root, text="Transport").pack()
entry_transport = tk.Entry(root)
entry_transport.pack()

tk.Label(root, text="Entertainment").pack()
entry_entertainment = tk.Entry(root)
entry_entertainment.pack()

tk.Button(root, text="Predict Expense", command=predict_expense).pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 14))
result_label.pack(pady=10)

root.mainloop()
