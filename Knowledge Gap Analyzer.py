# Knowledge Gap Analyzer - Desktop App in Python
# Purpose:
# Analyze learner topic scores and identify weak areas with recommendations.

import tkinter as tk
from tkinter import ttk, messagebox

RECOMMENDATIONS = {
    "CPU Architecture": "Review instruction cycle, datapath components, and control unit basics.",
    "Memory Hierarchy": "Revise cache mapping, cache levels (L1/L2/L3), and miss penalty concepts.",
    "Pipelining": "Practice hazards (data/control/structural) and forwarding/stalling techniques.",
    "Virtual Memory": "Study paging, TLB behavior, page faults, and replacement algorithms.",
    "Instruction Set": "Compare RISC vs CISC and focus on addressing modes and opcode formats.",
    "Parallelism": "Revisit ILP/TLP concepts, superscalar execution, and synchronization basics.",
}


def parse_topic_scores(raw_text):
    """Parse lines in the form: Topic, Score"""
    entries = []
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]

    for idx, line in enumerate(lines, start=1):
        if "," not in line:
            raise ValueError(f"Line {idx}: Use format 'Topic, Score'.")

        topic, score_text = [part.strip() for part in line.split(",", 1)]
        if not topic:
            raise ValueError(f"Line {idx}: Topic cannot be empty.")

        try:
            score = float(score_text)
        except ValueError as exc:
            raise ValueError(f"Line {idx}: Score must be a number.") from exc

        if not 0 <= score <= 100:
            raise ValueError(f"Line {idx}: Score must be between 0 and 100.")

        entries.append((topic, score))

    if not entries:
        raise ValueError("Please enter at least one topic and score.")

    return entries


def build_recommendation(topic):
    for key, advice in RECOMMENDATIONS.items():
        if key.lower() in topic.lower():
            return advice
    return "Review fundamentals, solve targeted practice questions, and re-test this topic."


def analyze_gaps():
    raw = input_box.get("1.0", tk.END)
    threshold = threshold_var.get()

    try:
        entries = parse_topic_scores(raw)
    except ValueError as err:
        messagebox.showerror("Input Error", str(err))
        return

    entries.sort(key=lambda item: item[1])
    weak_topics = [(topic, score) for topic, score in entries if score < threshold]
    average_score = sum(score for _, score in entries) / len(entries)

    output_box.delete("1.0", tk.END)
    output_box.insert(tk.END, "=== Knowledge Gap Analysis ===\n")
    output_box.insert(tk.END, f"Topics analyzed: {len(entries)}\n")
    output_box.insert(tk.END, f"Average score: {average_score:.2f}\n")
    output_box.insert(tk.END, f"Gap threshold: {threshold}\n\n")

    if not weak_topics:
        output_box.insert(tk.END, "Great job! No knowledge gaps found below the threshold.\n")
        return

    output_box.insert(tk.END, "Topics needing attention:\n")
    for rank, (topic, score) in enumerate(weak_topics, start=1):
        gap = threshold - score
        output_box.insert(
            tk.END,
            f"{rank}. {topic} -> Score: {score:.2f}, Gap: {gap:.2f}\n"
            f"   Recommendation: {build_recommendation(topic)}\n",
        )


# ---------------- Main Window ----------------
root = tk.Tk()
root.title("Knowledge Gap Analyzer")
root.geometry("820x600")
root.resizable(False, False)

style = ttk.Style(root)
style.theme_use("clam")

main_frame = ttk.Frame(root, padding=16)
main_frame.pack(fill=tk.BOTH, expand=True)

header = ttk.Label(main_frame, text="Knowledge Gap Analyzer", font=("Segoe UI", 16, "bold"))
header.pack(anchor="center", pady=(0, 10))

instruction = ttk.Label(
    main_frame,
    text="Enter one topic per line using format: Topic, Score (0-100). Example: Pipelining, 62",
)
instruction.pack(anchor="w")

input_box = tk.Text(main_frame, height=12, width=95)
input_box.pack(pady=8)

threshold_frame = ttk.Frame(main_frame)
threshold_frame.pack(fill=tk.X, pady=(0, 8))

threshold_label = ttk.Label(threshold_frame, text="Knowledge gap threshold:")
threshold_label.pack(side=tk.LEFT)

threshold_var = tk.IntVar(value=70)
threshold_spinbox = ttk.Spinbox(threshold_frame, from_=0, to=100, textvariable=threshold_var, width=6)
threshold_spinbox.pack(side=tk.LEFT, padx=8)

analyze_button = ttk.Button(threshold_frame, text="Analyze Gaps", command=analyze_gaps)
analyze_button.pack(side=tk.LEFT, padx=8)

output_title = ttk.Label(main_frame, text="Analysis Output")
output_title.pack(anchor="w")

output_box = tk.Text(main_frame, height=16, width=95)
output_box.pack(pady=8)

root.mainloop()