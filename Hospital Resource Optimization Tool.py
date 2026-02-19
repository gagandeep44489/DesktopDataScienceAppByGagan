import tkinter as tk
from tkinter import messagebox
import pulp

def optimize_resources():
    try:
        # Get inputs
        patients = int(entry_patients.get())
        doctor_cost = float(entry_doctor_cost.get())
        nurse_cost = float(entry_nurse_cost.get())
        bed_cost = float(entry_bed_cost.get())

        # Create LP Problem
        problem = pulp.LpProblem("Hospital_Optimization", pulp.LpMinimize)

        # Decision Variables
        doctors = pulp.LpVariable("Doctors", lowBound=0, cat='Integer')
        nurses = pulp.LpVariable("Nurses", lowBound=0, cat='Integer')
        beds = pulp.LpVariable("Beds", lowBound=0, cat='Integer')

        # Objective Function (Minimize Cost)
        problem += (doctor_cost * doctors +
                    nurse_cost * nurses +
                    bed_cost * beds), "Total_Cost"

        # Constraints
        problem += doctors * 10 >= patients      # 1 doctor handles 10 patients
        problem += nurses * 5 >= patients        # 1 nurse handles 5 patients
        problem += beds >= patients              # 1 bed per patient

        # Solve
        problem.solve()

        result_text = f"""
Optimization Result:

Doctors Required: {int(doctors.varValue)}
Nurses Required: {int(nurses.varValue)}
Beds Required: {int(beds.varValue)}

Minimum Total Cost: ₹ {pulp.value(problem.objective):,.2f}
        """

        text_result.delete(1.0, tk.END)
        text_result.insert(tk.END, result_text)

    except Exception as e:
        messagebox.showerror("Error", str(e))


# GUI Setup
root = tk.Tk()
root.title("Hospital Resource Optimization Tool")
root.geometry("500x550")

tk.Label(root, text="Hospital Resource Optimization",
         font=("Arial", 16, "bold")).pack(pady=10)

tk.Label(root, text="Number of Patients:").pack()
entry_patients = tk.Entry(root)
entry_patients.pack()

tk.Label(root, text="Cost per Doctor:").pack()
entry_doctor_cost = tk.Entry(root)
entry_doctor_cost.pack()

tk.Label(root, text="Cost per Nurse:").pack()
entry_nurse_cost = tk.Entry(root)
entry_nurse_cost.pack()

tk.Label(root, text="Cost per Bed:").pack()
entry_bed_cost = tk.Entry(root)
entry_bed_cost.pack()

tk.Button(root, text="Optimize Resources",
          command=optimize_resources,
          bg="green", fg="white").pack(pady=15)

text_result = tk.Text(root, height=12, width=55)
text_result.pack()

root.mainloop()
