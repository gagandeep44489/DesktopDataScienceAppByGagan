import tkinter as tk
from tkinter import messagebox
import pandas as pd
import matplotlib.pyplot as plt

# Store student data
students = []

# Grade Calculation Function
def calculate_grade(percentage):
    if percentage >= 90:
        return "A+"
    elif percentage >= 80:
        return "A"
    elif percentage >= 70:
        return "B"
    elif percentage >= 60:
        return "C"
    elif percentage >= 50:
        return "D"
    else:
        return "Fail"

# Analyze Performance
def analyze_performance():
    try:
        name = entry_name.get()
        math = float(entry_math.get())
        science = float(entry_science.get())
        english = float(entry_english.get())

        total = math + science + english
        average = total / 3
        percentage = (total / 300) * 100
        grade = calculate_grade(percentage)

        result_text.set(
            f"Total: {total}\n"
            f"Average: {average:.2f}\n"
            f"Percentage: {percentage:.2f}%\n"
            f"Grade: {grade}"
        )

        students.append({
            "Name": name,
            "Math": math,
            "Science": science,
            "English": english
        })

    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numeric marks!")

# Show Graph
def show_graph():
    if not students:
        messagebox.showwarning("No Data", "No student data available!")
        return

    latest = students[-1]

    subjects = ["Math", "Science", "English"]
    marks = [latest["Math"], latest["Science"], latest["English"]]

    plt.figure()
    plt.bar(subjects, marks)
    plt.xlabel("Subjects")
    plt.ylabel("Marks")
    plt.title(f"Performance of {latest['Name']}")
    plt.ylim(0, 100)
    plt.show()

# GUI Setup
root = tk.Tk()
root.title("Student Performance Analyzer")
root.geometry("450x500")
root.resizable(False, False)

tk.Label(root, text="Student Performance Analyzer", 
         font=("Arial", 16, "bold")).pack(pady=10)

tk.Label(root, text="Student Name").pack()
entry_name = tk.Entry(root)
entry_name.pack(pady=5)

tk.Label(root, text="Math Marks").pack()
entry_math = tk.Entry(root)
entry_math.pack(pady=5)

tk.Label(root, text="Science Marks").pack()
entry_science = tk.Entry(root)
entry_science.pack(pady=5)

tk.Label(root, text="English Marks").pack()
entry_english = tk.Entry(root)
entry_english.pack(pady=5)

tk.Button(root, text="Analyze Performance", 
          command=analyze_performance, bg="blue", fg="white").pack(pady=15)

tk.Button(root, text="Show Performance Graph", 
          command=show_graph, bg="green", fg="white").pack(pady=5)

result_text = tk.StringVar()
tk.Label(root, textvariable=result_text, 
         font=("Arial", 12, "bold")).pack(pady=15)

root.mainloop()