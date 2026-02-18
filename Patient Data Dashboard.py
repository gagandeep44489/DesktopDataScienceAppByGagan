import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox

# ------------------------
# 1. Sample Patient Data
# ------------------------
data = pd.DataFrame({
    "PatientID": [101, 102, 103, 104, 105],
    "Name": ["Alice", "Bob", "Charlie", "David", "Eve"],
    "Age": [29, 45, 34, 50, 41],
    "Gender": ["F", "M", "M", "M", "F"],
    "Diagnosis": ["Flu", "Diabetes", "Hypertension", "Flu", "Cold"],
    "RiskScore": [20, 70, 60, 25, 15]
})

# ------------------------
# 2. GUI
# ------------------------
root = tk.Tk()
root.title("Patient Data Dashboard")
root.geometry("700x400")

tk.Label(root, text="Patient Data Dashboard", font=("Arial", 16)).pack(pady=10)

# Frame for table
frame = tk.Frame(root)
frame.pack(pady=10)

# Scrollbar
scrollbar = tk.Scrollbar(frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Treeview (Table)
columns = list(data.columns)
tree = ttk.Treeview(frame, columns=columns, show='headings', yscrollcommand=scrollbar.set)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=100)
tree.pack()

scrollbar.config(command=tree.yview)

# Load data into table
def load_data(df):
    for row in tree.get_children():
        tree.delete(row)
    for _, row in df.iterrows():
        tree.insert("", tk.END, values=list(row))

load_data(data)

# ------------------------
# 3. Filter / Analytics
# ------------------------
filter_frame = tk.Frame(root)
filter_frame.pack(pady=10)

tk.Label(filter_frame, text="Filter by Diagnosis:").grid(row=0, column=0, padx=5)
diagnosis_var = tk.StringVar()
diagnosis_entry = tk.Entry(filter_frame, textvariable=diagnosis_var)
diagnosis_entry.grid(row=0, column=1, padx=5)

def filter_data():
    diagnosis = diagnosis_var.get().strip()
    if diagnosis:
        filtered = data[data["Diagnosis"].str.contains(diagnosis, case=False)]
        if filtered.empty:
            messagebox.showinfo("No Data", f"No patients found with diagnosis '{diagnosis}'")
        load_data(filtered)
    else:
        load_data(data)

tk.Button(filter_frame, text="Filter", command=filter_data, bg="blue", fg="white").grid(row=0, column=2, padx=5)
tk.Button(filter_frame, text="Reset", command=lambda: load_data(data)).grid(row=0, column=3, padx=5)

# ------------------------
# 4. Analytics: Average Risk Score
# ------------------------
def show_avg_risk():
    avg = data["RiskScore"].mean()
    messagebox.showinfo("Analytics", f"Average Risk Score of Patients: {avg:.2f}%")

tk.Button(root, text="Show Average Risk Score", command=show_avg_risk, bg="green", fg="white").pack(pady=10)

# ------------------------
# 5. Run the App
# ------------------------
root.mainloop()
