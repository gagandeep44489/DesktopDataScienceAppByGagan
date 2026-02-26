import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt

df = None

def load_file():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if not file_path:
        return
    
    try:
        df = pd.read_csv(file_path)
        messagebox.showinfo("Success", "File Loaded Successfully")
        update_stats()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file:\n{e}")

def update_stats():
    if df is None:
        return

    avg_scores = df.mean(numeric_only=True)

    stats_text.delete("1.0", tk.END)
    stats_text.insert(tk.END, "📊 Average Scores:\n\n")

    for subject, score in avg_scores.items():
        stats_text.insert(tk.END, f"{subject}: {score:.2f}\n")

def plot_bar_chart():
    if df is None:
        messagebox.showwarning("Warning", "Load CSV file first.")
        return

    avg_scores = df.mean(numeric_only=True)

    plt.figure()
    avg_scores.plot(kind='bar')
    plt.title("Average Subject Scores")
    plt.xlabel("Subjects")
    plt.ylabel("Average Score")
    plt.tight_layout()
    plt.show()

def plot_student_performance():
    if df is None:
        messagebox.showwarning("Warning", "Load CSV file first.")
        return

    student = df.iloc[0]  # first student example

    subjects = df.columns[1:]
    scores = student[1:]

    plt.figure()
    plt.plot(subjects, scores, marker='o')
    plt.title(f"Performance of {student['Student']}")
    plt.xlabel("Subjects")
    plt.ylabel("Score")
    plt.tight_layout()
    plt.show()

# GUI Setup
root = tk.Tk()
root.title("Interactive Learning Analytics Dashboard")
root.geometry("750x500")

tk.Label(root, text="Learning Analytics Dashboard", 
         font=("Arial", 16, "bold")).pack(pady=10)

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(button_frame, text="Load CSV Data", 
          command=load_file, bg="green", fg="white").pack(side="left", padx=5)

tk.Button(button_frame, text="Show Average Bar Chart", 
          command=plot_bar_chart, bg="blue", fg="white").pack(side="left", padx=5)

tk.Button(button_frame, text="Show Student Performance", 
          command=plot_student_performance, bg="purple", fg="white").pack(side="left", padx=5)

tk.Label(root, text="Statistics:", font=("Arial", 12, "bold")).pack(pady=5)

stats_text = tk.Text(root, height=15, wrap="word", bg="#f4f4f4")
stats_text.pack(fill="both", padx=10, pady=5)

root.mainloop()