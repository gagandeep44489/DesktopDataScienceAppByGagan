import tkinter as tk
from tkinter import messagebox


def parse_percent(entry_widget, field_name):
    text = entry_widget.get().strip()
    if text == "":
        raise ValueError(f"{field_name} cannot be empty")
    value = float(text)
    if value < 0 or value > 100:
        raise ValueError(f"{field_name} must be between 0 and 100")
    return value


def calculate_estimates():
    try:
        coursework_score = parse_percent(entry_coursework_score, "Coursework score")
        coursework_weight = parse_percent(entry_coursework_weight, "Coursework weight")
        final_exam_score = parse_percent(entry_final_exam_score, "Expected final exam score")
        final_exam_weight = parse_percent(entry_final_exam_weight, "Final exam weight")
        target_grade = parse_percent(entry_target_grade, "Target overall grade")

        total_weight = coursework_weight + final_exam_weight
        if abs(total_weight - 100.0) > 1e-9:
            raise ValueError("Coursework weight + final exam weight must equal 100")

        estimated_overall = (
            coursework_score * coursework_weight
            + final_exam_score * final_exam_weight
        ) / 100.0

        if final_exam_weight == 0:
            required_for_target = 0.0 if coursework_score >= target_grade else None
        else:
            required_for_target = (
                target_grade * 100.0 - coursework_score * coursework_weight
            ) / final_exam_weight

        result_lines = [f"Estimated Overall Grade: {estimated_overall:.2f}%"]

        if required_for_target is None:
            result_lines.append("Required Final Exam Score for target: Not possible (final exam weight is 0%).")
        elif required_for_target < 0:
            result_lines.append("Required Final Exam Score for target: 0.00% (target already secured).")
        elif required_for_target > 100:
            result_lines.append(
                f"Required Final Exam Score for target: {required_for_target:.2f}% (above 100%, target not achievable)."
            )
        else:
            result_lines.append(f"Required Final Exam Score for target: {required_for_target:.2f}%")

        result_label.config(text="\n".join(result_lines))

    except Exception as exc:
        messagebox.showerror("Input Error", str(exc))


# ---------------- GUI ---------------- #
root = tk.Tk()
root.title("Exam Score Estimator")
root.geometry("560x430")

header = tk.Label(
    root,
    text="Estimate your overall grade and required final exam score",
    font=("Arial", 12, "bold"),
)
header.pack(pady=10)

frame = tk.Frame(root)
frame.pack(pady=5)


def add_labeled_entry(parent, label_text, default_value):
    row = tk.Frame(parent)
    row.pack(fill="x", pady=4)

    label = tk.Label(row, text=label_text, width=35, anchor="w")
    label.pack(side="left", padx=5)

    entry = tk.Entry(row, width=18)
    entry.insert(0, default_value)
    entry.pack(side="right", padx=5)

    return entry


entry_coursework_score = add_labeled_entry(frame, "Current coursework score (%):", "78")
entry_coursework_weight = add_labeled_entry(frame, "Coursework weight (%):", "60")
entry_final_exam_score = add_labeled_entry(frame, "Expected final exam score (%):", "75")
entry_final_exam_weight = add_labeled_entry(frame, "Final exam weight (%):", "40")
entry_target_grade = add_labeled_entry(frame, "Target overall grade (%):", "80")

calc_button = tk.Button(
    root,
    text="Estimate",
    command=calculate_estimates,
    bg="#2E8B57",
    fg="white",
    width=18,
)
calc_button.pack(pady=14)

result_label = tk.Label(
    root,
    text="Estimated Overall Grade: --\nRequired Final Exam Score for target: --",
    font=("Arial", 11),
    justify="left",
)
result_label.pack(pady=10)

note = tk.Label(
    root,
    text="Tip: Weights must add up to 100.",
    fg="#555555",
    font=("Arial", 9),
)
note.pack(pady=2)

root.mainloop()