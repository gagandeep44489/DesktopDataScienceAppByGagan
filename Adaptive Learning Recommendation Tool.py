import tkinter as tk
from tkinter import messagebox

# Recommendation Logic
def generate_recommendations():
    try:
        math = float(entry_math.get())
        science = float(entry_science.get())
        english = float(entry_english.get())

        recommendations = []

        if math < 60:
            recommendations.append("📘 Math: Revise algebra & practice problem-solving daily.")
        elif math < 80:
            recommendations.append("📘 Math: Practice advanced numerical exercises.")

        if science < 60:
            recommendations.append("🔬 Science: Focus on core concepts and diagrams.")
        elif science < 80:
            recommendations.append("🔬 Science: Practice application-based questions.")

        if english < 60:
            recommendations.append("📖 English: Improve grammar and reading comprehension.")
        elif english < 80:
            recommendations.append("📖 English: Work on vocabulary and writing skills.")

        avg = (math + science + english) / 3

        if avg >= 85:
            recommendations.append("🌟 Excellent performance! Start competitive-level preparation.")
        elif avg < 60:
            recommendations.append("⚠ Overall performance needs structured revision planning.")

        if not recommendations:
            recommendations.append("✅ Good performance! Maintain consistent study routine.")

        result_text.set("\n\n".join(recommendations))

    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numeric marks!")

# GUI Setup
root = tk.Tk()
root.title("Adaptive Learning Recommendation Tool")
root.geometry("500x550")
root.resizable(False, False)

tk.Label(root, text="Adaptive Learning Recommendation Tool",
         font=("Arial", 14, "bold")).pack(pady=15)

tk.Label(root, text="Enter Math Marks").pack()
entry_math = tk.Entry(root)
entry_math.pack(pady=5)

tk.Label(root, text="Enter Science Marks").pack()
entry_science = tk.Entry(root)
entry_science.pack(pady=5)

tk.Label(root, text="Enter English Marks").pack()
entry_english = tk.Entry(root)
entry_english.pack(pady=5)

tk.Button(root, text="Generate Recommendations",
          command=generate_recommendations,
          bg="blue", fg="white").pack(pady=15)

result_text = tk.StringVar()
tk.Label(root, textvariable=result_text,
         wraplength=450,
         justify="left",
         font=("Arial", 11)).pack(pady=20)

root.mainloop()