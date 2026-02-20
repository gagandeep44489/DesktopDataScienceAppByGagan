import tkinter as tk
from tkinter import messagebox
import pandas as pd
import matplotlib.pyplot as plt

appliances = []

# ==============================
# ADD APPLIANCE FUNCTION
# ==============================

def add_appliance():
    try:
        name = name_var.get()
        watts = float(watts_var.get())
        hours = float(hours_var.get())

        daily_kwh = (watts * hours) / 1000
        monthly_kwh = daily_kwh * 30

        appliances.append({
            "Name": name,
            "Watts": watts,
            "Hours": hours,
            "Monthly_kWh": monthly_kwh
        })

        messagebox.showinfo("Added", f"{name} added successfully")

        name_var.set("")
        watts_var.set("")
        hours_var.set("")

    except:
        messagebox.showerror("Error", "Enter valid numeric values")

# ==============================
# SHOW ANALYSIS
# ==============================

def analyze():
    if not appliances:
        messagebox.showwarning("No Data", "Add appliances first")
        return

    rate = float(rate_var.get())
    df = pd.DataFrame(appliances)

    total_kwh = df["Monthly_kWh"].sum()
    total_cost = total_kwh * rate

    result_label.config(
        text=f"Total Monthly Consumption: {total_kwh:.2f} kWh\nEstimated Bill: ₹{total_cost:.2f}"
    )

    show_chart(df)

def show_chart(df):
    plt.figure()
    plt.bar(df["Name"], df["Monthly_kWh"])
    plt.xticks(rotation=45)
    plt.ylabel("Monthly kWh")
    plt.title("Appliance-wise Energy Consumption")
    plt.tight_layout()
    plt.show()

# ==============================
# GUI SETUP
# ==============================

root = tk.Tk()
root.title("Smart Home Energy Analyzer")
root.geometry("450x600")

tk.Label(root, text="Appliance Name").pack()
name_var = tk.StringVar()
tk.Entry(root, textvariable=name_var).pack()

tk.Label(root, text="Power Rating (Watts)").pack()
watts_var = tk.StringVar()
tk.Entry(root, textvariable=watts_var).pack()

tk.Label(root, text="Usage Hours per Day").pack()
hours_var = tk.StringVar()
tk.Entry(root, textvariable=hours_var).pack()

tk.Button(root, text="Add Appliance", command=add_appliance).pack(pady=10)

tk.Label(root, text="Electricity Rate (₹ per kWh)").pack()
rate_var = tk.StringVar()
tk.Entry(root, textvariable=rate_var).pack()

tk.Button(root, text="Analyze Consumption", command=analyze).pack(pady=15)

result_label = tk.Label(root, text="", fg="blue")
result_label.pack()

root.mainloop()