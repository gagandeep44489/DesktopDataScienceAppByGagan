# Student Performance Predictor - Desktop App in Python
# Purpose:
# Predict student final exam performance using manually entered academic indicators.
# Useful for classroom analytics demos and basic educational data science practice.

import tkinter as tk
from tkinter import ttk, messagebox


def clamp(value, low, high):
    return max(low, min(high, value))


def parse_student_rows(raw_text: str):
    """
    Expected row format (comma-separated):
    student_name, hours_studied_per_week, attendance_percent, previous_exam_score,
    assignments_completed_percent, sleep_hours_per_day
    """
    students = []
    errors = []

    for line_no, line in enumerate(raw_text.strip().splitlines(), start=1):
        if not line.strip():
            continue

        parts = [p.strip() for p in line.split(',')]
        if len(parts) != 6:
            errors.append(f"Line {line_no}: expected 6 comma-separated values")
            continue

        name, study_hours, attendance, prev_score, assignment_rate, sleep_hours = parts

        if not name:
            errors.append(f"Line {line_no}: student name cannot be empty")
            continue

        try:
            study_hours = float(study_hours)
            attendance = float(attendance)
            prev_score = float(prev_score)
            assignment_rate = float(assignment_rate)
            sleep_hours = float(sleep_hours)
        except ValueError:
            errors.append(f"Line {line_no}: numeric fields contain invalid values")
            continue

        if study_hours < 0:
            errors.append(f"Line {line_no}: study hours must be >= 0")
            continue
        if not (0 <= attendance <= 100):
            errors.append(f"Line {line_no}: attendance must be in range 0-100")
            continue
        if not (0 <= prev_score <= 100):
            errors.append(f"Line {line_no}: previous score must be in range 0-100")
            continue
        if not (0 <= assignment_rate <= 100):
            errors.append(f"Line {line_no}: assignments completed must be in range 0-100")
            continue
        if not (0 <= sleep_hours <= 24):
            errors.append(f"Line {line_no}: sleep hours must be in range 0-24")
            continue

        students.append(
            {
                "name": name,
                "study_hours": study_hours,
                "attendance": attendance,
                "prev_score": prev_score,
                "assignment_rate": assignment_rate,
                "sleep_hours": sleep_hours,
            }
        )

    return students, errors


def predict_score(student):
    """
    A weighted educational heuristic model.
    Weights sum to ~1 before bonuses/penalties:
      study_hours (20%), attendance (20%), previous score (35%), assignments (20%), sleep quality (5%)
    """
    study_component = clamp((student["study_hours"] / 30) * 100, 0, 100)
    sleep_component = clamp((student["sleep_hours"] / 8) * 100, 0, 100)

    weighted_score = (
        0.20 * study_component
        + 0.20 * student["attendance"]
        + 0.35 * student["prev_score"]
        + 0.20 * student["assignment_rate"]
        + 0.05 * sleep_component
    )

    bonus = 0
    if student["study_hours"] >= 15 and student["attendance"] >= 90:
        bonus += 3
    if student["assignment_rate"] >= 95:
        bonus += 2

    penalty = 0
    if student["sleep_hours"] < 5:
        penalty += 3
    if student["attendance"] < 60:
        penalty += 5

    predicted = clamp(weighted_score + bonus - penalty, 0, 100)

    if predicted >= 85:
        risk = "Low Risk"
        recommendation = "Maintain consistency and attempt advanced practice papers."
    elif predicted >= 70:
        risk = "Moderate Risk"
        recommendation = "Improve weak topics and increase weekly revision frequency."
    elif predicted >= 50:
        risk = "High Risk"
        recommendation = "Start structured daily study plan and seek teacher guidance."
    else:
        risk = "Critical Risk"
        recommendation = "Immediate intervention: tutoring, mentor check-ins, and attendance improvement."

    return round(predicted, 2), risk, recommendation


def analyze_performance():
    raw = input_text.get("1.0", tk.END)
    students, errors = parse_student_rows(raw)

    output_text.delete("1.0", tk.END)
    result_tree.delete(*result_tree.get_children())

    if errors:
        messagebox.showwarning(
            "Input Warnings",
            "Some rows were skipped due to errors:\n\n" + "\n".join(errors[:10])
            + ("\n..." if len(errors) > 10 else ""),
        )

    if not students:
        messagebox.showerror("No Valid Data", "Please enter at least one valid student row.")
        return

    total_predicted = 0
    risk_counts = {"Low Risk": 0, "Moderate Risk": 0, "High Risk": 0, "Critical Risk": 0}
    topper_name = ""
    topper_score = -1

    for student in students:
        score, risk, recommendation = predict_score(student)
        total_predicted += score
        risk_counts[risk] += 1

        if score > topper_score:
            topper_name = student["name"]
            topper_score = score

        result_tree.insert(
            "",
            tk.END,
            values=(
                student["name"],
                f"{score:.2f}",
                risk,
                recommendation,
            ),
        )

    avg_predicted = total_predicted / len(students)

    summary = [
        "Student Performance Prediction Report",
        "=" * 38,
        f"Total Students Processed: {len(students)}",
        f"Average Predicted Score: {avg_predicted:.2f}",
        f"Top Predicted Performer: {topper_name} ({topper_score:.2f})",
        "",
        "Risk Distribution:",
        f"  - Low Risk: {risk_counts['Low Risk']}",
        f"  - Moderate Risk: {risk_counts['Moderate Risk']}",
        f"  - High Risk: {risk_counts['High Risk']}",
        f"  - Critical Risk: {risk_counts['Critical Risk']}",
    ]

    output_text.insert(tk.END, "\n".join(summary))


def load_sample_data():
    sample = """Aanya,16,94,88,97,7.5
Rohan,9,81,72,78,6.3
Kabir,4,59,61,52,4.8
Meera,13,89,84,91,7.1
Ishita,6,68,58,63,5.5
Arjun,18,96,91,99,8.0
"""
    input_text.delete("1.0", tk.END)
    input_text.insert(tk.END, sample)


root = tk.Tk()
root.title("Student Performance Predictor")
root.geometry("1080x740")

style = ttk.Style(root)
style.theme_use("clam")

main = ttk.Frame(root, padding=12)
main.pack(fill=tk.BOTH, expand=True)

header = ttk.Label(main, text="Student Performance Predictor", font=("Segoe UI", 16, "bold"))
header.pack(anchor="w", pady=(0, 8))

instructions = ttk.Label(
    main,
    text=(
        "Enter one student per line: "
        "name, hours_studied_per_week, attendance_percent, previous_exam_score, "
        "assignments_completed_percent, sleep_hours_per_day"
    ),
)
instructions.pack(anchor="w")

input_text = tk.Text(main, height=11, width=140)
input_text.pack(fill=tk.X, pady=8)

buttons = ttk.Frame(main)
buttons.pack(fill=tk.X, pady=(0, 8))

ttk.Button(buttons, text="Load Sample Data", command=load_sample_data).pack(side=tk.LEFT)
ttk.Button(buttons, text="Predict Performance", command=analyze_performance).pack(side=tk.LEFT, padx=8)

output_label = ttk.Label(main, text="Summary")
output_label.pack(anchor="w")

output_text = tk.Text(main, height=12, width=140)
output_text.pack(fill=tk.BOTH, expand=False, pady=6)

tree_label = ttk.Label(main, text="Student-wise Predictions")
tree_label.pack(anchor="w", pady=(6, 2))

columns = ("student", "predicted_score", "risk", "recommendation")
result_tree = ttk.Treeview(main, columns=columns, show="headings", height=12)

result_tree.heading("student", text="Student")
result_tree.heading("predicted_score", text="Predicted Score")
result_tree.heading("risk", text="Risk Level")
result_tree.heading("recommendation", text="Recommendation")

result_tree.column("student", width=140, anchor="w")
result_tree.column("predicted_score", width=130, anchor="center")
result_tree.column("risk", width=130, anchor="center")
result_tree.column("recommendation", width=620, anchor="w")

scrollbar = ttk.Scrollbar(main, orient="vertical", command=result_tree.yview)
result_tree.configure(yscrollcommand=scrollbar.set)

result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

root.mainloop()