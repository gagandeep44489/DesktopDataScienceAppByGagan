import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox

# -------------------------
# 1. Sample Drug Dataset
# -------------------------
# Effectiveness Score: 0-100 (post-treatment improvement)
data = pd.DataFrame({
    "PatientID": [201, 202, 203, 204, 205, 206, 207],
    "DrugName": ["DrugA", "DrugB", "DrugA", "DrugC", "DrugB", "DrugA", "DrugC"],
    "Dosage_mg": [50, 100, 50, 75, 100, 50, 75],
    "InitialSeverity": [80, 70, 60, 90, 75, 85, 88],
    "PostTreatmentSeverity": [30, 40, 25, 50, 45, 35, 55]
})

# Calculate Effectiveness %
data["ImprovementPercent"] = (
    (data["InitialSeverity"] - data["PostTreatmentSeverity"]) 
    / data["InitialSeverity"]
) * 100


# -------------------------
# 2. GUI Setup
# -------------------------
root = tk.Tk()
root.title("Drug Effectiveness Analyzer")
root.geometry("850x450")

tk.Label(root, text="Drug Effectiveness Analyzer Dashboard",
         font=("Arial", 16)).pack(pady=10)

# -------------------------
# 3. Table View
# -------------------------
frame = tk.Frame(root)
frame.pack()

scrollbar = tk.Scrollbar(frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

columns = list(data.columns)
tree = ttk.Treeview(frame, columns=columns, show="headings",
                    yscrollcommand=scrollbar.set)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=120)

tree.pack()
scrollbar.config(command=tree.yview)

def load_data(df):
    for row in tree.get_children():
        tree.delete(row)
    for _, row in df.iterrows():
        tree.insert("", tk.END, values=list(row))

load_data(data)

# -------------------------
# 4. Drug Analysis Section
# -------------------------
analysis_frame = tk.Frame(root)
analysis_frame.pack(pady=10)

tk.Label(analysis_frame, text="Analyze Drug:").grid(row=0, column=0, padx=5)

drug_var = tk.StringVar()
drug_entry = tk.Entry(analysis_frame, textvariable=drug_var)
drug_entry.grid(row=0, column=1, padx=5)

def analyze_drug():
    drug = drug_var.get().strip()
    if not drug:
        messagebox.showwarning("Input Error", "Please enter a Drug Name")
        return

    filtered = data[data["DrugName"].str.lower() == drug.lower()]

    if filtered.empty:
        messagebox.showinfo("No Data", f"No records found for {drug}")
        return

    avg_effectiveness = filtered["ImprovementPercent"].mean()

    messagebox.showinfo(
        "Drug Analysis Result",
        f"Drug: {drug}\n"
        f"Number of Patients: {len(filtered)}\n"
        f"Average Effectiveness: {avg_effectiveness:.2f}%"
    )

tk.Button(analysis_frame, text="Analyze",
          command=analyze_drug,
          bg="blue", fg="white").grid(row=0, column=2, padx=5)

# -------------------------
# 5. Overall Best Drug
# -------------------------
def best_drug():
    summary = data.groupby("DrugName")["ImprovementPercent"].mean()
    best = summary.idxmax()
    best_score = summary.max()

    messagebox.showinfo(
        "Best Performing Drug",
        f"Best Drug: {best}\n"
        f"Average Effectiveness: {best_score:.2f}%"
    )

tk.Button(root, text="Find Best Performing Drug",
          command=best_drug,
          bg="green", fg="white").pack(pady=10)

root.mainloop()
