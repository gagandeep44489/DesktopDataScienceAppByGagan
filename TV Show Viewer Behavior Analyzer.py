# TV Show Viewer Behavior Analyzer - Desktop App in Python
# Purpose:
# Analyze viewer behavior for TV shows using manually entered session data.
# Useful for media analytics education, prototyping, and reporting.

import tkinter as tk
from tkinter import ttk, messagebox
from collections import Counter, defaultdict


def parse_sessions(raw_text: str):
    """
    Parse user-entered rows.
    Expected row format:
    show_name, viewer_id, watch_minutes, rating_1_to_10, completed_yes_no, platform
    """
    sessions = []
    errors = []

    for line_no, line in enumerate(raw_text.strip().splitlines(), start=1):
        if not line.strip():
            continue

        parts = [p.strip() for p in line.split(',')]
        if len(parts) != 6:
            errors.append(f"Line {line_no}: expected 6 comma-separated values")
            continue

        show_name, viewer_id, watch_minutes, rating, completed, platform = parts

        try:
            watch_minutes = float(watch_minutes)
            if watch_minutes < 0:
                raise ValueError
        except ValueError:
            errors.append(f"Line {line_no}: invalid watch_minutes '{parts[2]}'")
            continue

        try:
            rating = float(rating)
            if not (1 <= rating <= 10):
                raise ValueError
        except ValueError:
            errors.append(f"Line {line_no}: rating must be in range 1-10")
            continue

        completed_norm = completed.lower()
        if completed_norm not in {"yes", "no"}:
            errors.append(f"Line {line_no}: completed value must be yes/no")
            continue

        if not show_name or not viewer_id or not platform:
            errors.append(f"Line {line_no}: show, viewer_id, and platform cannot be empty")
            continue

        sessions.append(
            {
                "show": show_name,
                "viewer": viewer_id,
                "minutes": watch_minutes,
                "rating": rating,
                "completed": completed_norm == "yes",
                "platform": platform,
            }
        )

    return sessions, errors


def viewer_segment(total_minutes, sessions_count, completion_rate):
    if sessions_count >= 8 and completion_rate >= 0.7:
        return "Binge Viewer"
    if total_minutes >= 300 and completion_rate >= 0.5:
        return "Committed Viewer"
    if sessions_count <= 2 and total_minutes < 120:
        return "Casual Viewer"
    return "Regular Viewer"


def analyze_behavior():
    raw = input_text.get("1.0", tk.END)
    sessions, errors = parse_sessions(raw)

    output_text.delete("1.0", tk.END)
    detail_tree.delete(*detail_tree.get_children())

    if errors:
        messagebox.showwarning(
            "Input Warnings",
            "Some rows were skipped due to errors:\n\n" + "\n".join(errors[:10])
            + ("\n..." if len(errors) > 10 else ""),
        )

    if not sessions:
        messagebox.showerror("No Valid Data", "Please enter at least one valid session row.")
        return

    total_sessions = len(sessions)
    unique_viewers = len({s["viewer"] for s in sessions})
    total_watch_minutes = sum(s["minutes"] for s in sessions)
    avg_watch_per_session = total_watch_minutes / total_sessions
    completion_rate = sum(1 for s in sessions if s["completed"]) / total_sessions

    show_counter = Counter(s["show"] for s in sessions)
    show_rating_sum = defaultdict(float)
    show_rating_count = Counter()
    platform_counter = Counter(s["platform"] for s in sessions)

    viewer_agg = defaultdict(lambda: {"minutes": 0.0, "sessions": 0, "completed": 0})

    for s in sessions:
        show_rating_sum[s["show"]] += s["rating"]
        show_rating_count[s["show"]] += 1

        v = viewer_agg[s["viewer"]]
        v["minutes"] += s["minutes"]
        v["sessions"] += 1
        if s["completed"]:
            v["completed"] += 1

    top_shows = show_counter.most_common(5)
    top_platforms = platform_counter.most_common(3)

    summary_lines = [
        "TV Show Viewer Behavior Analysis",
        "=" * 36,
        f"Total Sessions: {total_sessions}",
        f"Unique Viewers: {unique_viewers}",
        f"Total Watch Time: {total_watch_minutes:.1f} minutes",
        f"Average Watch Time / Session: {avg_watch_per_session:.1f} minutes",
        f"Overall Completion Rate: {completion_rate * 100:.1f}%",
        "",
        "Top Shows by Sessions:",
    ]

    for show, count in top_shows:
        avg_rating = show_rating_sum[show] / show_rating_count[show]
        summary_lines.append(f"  - {show}: {count} sessions, avg rating {avg_rating:.2f}/10")

    summary_lines.append("\nTop Platforms:")
    for platform, count in top_platforms:
        summary_lines.append(f"  - {platform}: {count} sessions")

    output_text.insert(tk.END, "\n".join(summary_lines))

    for viewer, stats in sorted(viewer_agg.items(), key=lambda x: x[1]["minutes"], reverse=True):
        v_completion_rate = stats["completed"] / stats["sessions"]
        segment = viewer_segment(stats["minutes"], stats["sessions"], v_completion_rate)
        detail_tree.insert(
            "",
            tk.END,
            values=(
                viewer,
                stats["sessions"],
                f"{stats['minutes']:.1f}",
                f"{v_completion_rate * 100:.1f}%",
                segment,
            ),
        )


def load_sample_data():
    sample = """Planet Frontiers,V001,52,9,yes,StreamX
Planet Frontiers,V001,49,8.5,yes,StreamX
City Cops,V002,24,7,no,TVNow
Retro Quest,V003,61,9.5,yes,BingeBox
Retro Quest,V003,58,9,yes,BingeBox
Retro Quest,V004,45,8,no,BingeBox
City Cops,V005,32,6.5,no,TVNow
Planet Frontiers,V006,55,8,yes,StreamX
Comedy Nest,V002,28,7.5,no,TVNow
Comedy Nest,V007,26,8,no,MobilePlay
"""
    input_text.delete("1.0", tk.END)
    input_text.insert(tk.END, sample)


root = tk.Tk()
root.title("TV Show Viewer Behavior Analyzer")
root.geometry("980x720")

style = ttk.Style(root)
style.theme_use("clam")

main = ttk.Frame(root, padding=12)
main.pack(fill=tk.BOTH, expand=True)

header = ttk.Label(main, text="TV Show Viewer Behavior Analyzer", font=("Segoe UI", 16, "bold"))
header.pack(anchor="w", pady=(0, 8))

instructions = ttk.Label(
    main,
    text=(
        "Enter one session per line: "
        "show_name, viewer_id, watch_minutes, rating_1_to_10, completed_yes_no, platform"
    ),
)
instructions.pack(anchor="w")

input_text = tk.Text(main, height=11, width=120)
input_text.pack(fill=tk.X, pady=8)

button_row = ttk.Frame(main)
button_row.pack(fill=tk.X, pady=(0, 8))

ttk.Button(button_row, text="Load Sample Data", command=load_sample_data).pack(side=tk.LEFT)
ttk.Button(button_row, text="Analyze Behavior", command=analyze_behavior).pack(side=tk.LEFT, padx=8)

output_label = ttk.Label(main, text="Summary")
output_label.pack(anchor="w")

output_text = tk.Text(main, height=13, width=120)
output_text.pack(fill=tk.BOTH, expand=False, pady=6)

viewer_label = ttk.Label(main, text="Viewer Segments")
viewer_label.pack(anchor="w", pady=(6, 2))

columns = ("viewer", "sessions", "minutes", "completion", "segment")
detail_tree = ttk.Treeview(main, columns=columns, show="headings", height=10)
for col, width in [
    ("viewer", 140),
    ("sessions", 90),
    ("minutes", 130),
    ("completion", 130),
    ("segment", 220),
]:
    detail_tree.heading(col, text=col.capitalize())
    detail_tree.column(col, width=width, anchor=tk.CENTER)

detail_tree.pack(fill=tk.BOTH, expand=True)

load_sample_data()
root.mainloop()