import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd

# Global dataframe
df = None

# -----------------------------
# Function to Load CSV File
# -----------------------------
def load_file():
    global df
    file_path = filedialog.askopenfilename(
        title="Select Course CSV File",
        filetypes=[("CSV Files", "*.csv")]
    )
    
    if file_path:
        try:
            df = pd.read_csv(file_path)
            
            required_columns = {"Course", "Category", "Level", "Keywords"}
            if not required_columns.issubset(df.columns):
                messagebox.showerror(
                    "Invalid File",
                    "CSV must contain columns:\nCourse, Category, Level, Keywords"
                )
                df = None
                return
            
            messagebox.showinfo("Success", "Course file loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")

# -----------------------------
# Recommendation Logic
# -----------------------------
def recommend_courses():
    global df
    
    if df is None:
        messagebox.showerror("Error", "Please load a course CSV file first.")
        return

    interest = interest_entry.get().strip().lower()
    level = level_var.get()

    if not interest:
        messagebox.showwarning("Input Required", "Please enter your interests.")
        return

    # Filter by level
    filtered = df[df["Level"].str.lower() == level.lower()]

    recommendations = []

    for _, row in filtered.iterrows():
        keywords = str(row["Keywords"]).lower()
        score = sum(word in keywords for word in interest.split())
        
        if score > 0:
            recommendations.append((row["Course"], score))

    recommendations.sort(key=lambda x: x[1], reverse=True)

    result_text.delete("1.0", tk.END)

    if recommendations:
        result_text.insert(tk.END, "Recommended Courses:\n\n")
        for course, score in recommendations:
            result_text.insert(
                tk.END,
                f"{course}  |  Match Score: {score}\n"
            )
    else:
        result_text.insert(tk.END, "No matching courses found.")

# -----------------------------
# GUI Setup
# -----------------------------
root = tk.Tk()
root.title("Course Recommendation Engine")
root.geometry("650x500")
root.resizable(False, False)

title_label = tk.Label(
    root,
    text="Course Recommendation Engine",
    font=("Arial", 16, "bold")
)
title_label.pack(pady=10)

# Load File Button
load_button = tk.Button(
    root,
    text="Load Course CSV File",
    command=load_file,
    bg="blue",
    fg="white",
    width=25
)
load_button.pack(pady=5)

# Interest Input
tk.Label(root, text="Enter Your Interests (e.g., python data ai):").pack(pady=5)

interest_entry = tk.Entry(root, width=50)
interest_entry.pack(pady=5)

# Level Selection
tk.Label(root, text="Select Your Skill Level:").pack(pady=5)

level_var = tk.StringVar()
level_var.set("Beginner")

level_menu = tk.OptionMenu(
    root,
    level_var,
    "Beginner",
    "Intermediate",
    "Advanced"
)
level_menu.pack(pady=5)

# Recommend Button
recommend_button = tk.Button(
    root,
    text="Recommend Courses",
    command=recommend_courses,
    bg="green",
    fg="white",
    width=25
)
recommend_button.pack(pady=10)

# Results Box
tk.Label(root, text="Results:", font=("Arial", 12, "bold")).pack(pady=5)

result_text = tk.Text(
    root,
    height=15,
    width=70,
    wrap="word",
    bg="#f4f4f4"
)
result_text.pack(padx=10, pady=5)

root.mainloop()