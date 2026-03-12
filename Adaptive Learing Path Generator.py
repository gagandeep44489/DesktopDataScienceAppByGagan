import tkinter as tk
from tkinter import ttk, messagebox


TOPIC_CATALOG = {
    "Python Basics": ["variables", "loops", "functions", "syntax", "beginner"],
    "Data Structures": ["arrays", "lists", "dict", "stack", "queue", "trees", "graphs"],
    "Algorithms": ["sorting", "search", "recursion", "dynamic programming", "greedy"],
    "Databases": ["sql", "database", "postgres", "mysql", "nosql"],
    "Web Development": ["web", "api", "backend", "frontend", "http", "flask", "django"],
    "Machine Learning": ["ml", "machine learning", "model", "regression", "classification"],
    "System Design": ["scalability", "distributed", "load balancing", "design"],
    "Testing & Quality": ["test", "unit test", "integration", "qa", "debug"],
}

LEVEL_TO_DAYS = {
    "Beginner": 12,
    "Intermediate": 8,
    "Advanced": 5,
}

STYLE_MULTIPLIER = {
    "Balanced": 1.0,
    "Hands-on": 0.8,
    "Theory-first": 1.2,
}


def normalize_text(text: str) -> str:
    return " ".join(text.strip().lower().split())


def infer_priority_topics(goal: str, interests: str):
    combined = f"{normalize_text(goal)} {normalize_text(interests)}"
    scored = []
    for topic, keywords in TOPIC_CATALOG.items():
        score = sum(1 for key in keywords if key in combined)
        scored.append((topic, score))

    scored.sort(key=lambda x: x[1], reverse=True)
    non_zero = [topic for topic, score in scored if score > 0]

    if len(non_zero) >= 4:
        return non_zero[:4]
    if len(non_zero) > 0:
        remaining = [topic for topic, score in scored if topic not in non_zero]
        return (non_zero + remaining)[:4]

    return [
        "Python Basics",
        "Data Structures",
        "Algorithms",
        "Testing & Quality",
    ]


def estimate_weeks(weekly_hours: int, base_days_per_topic: int, style_factor: float, topic_count: int):
    total_days = base_days_per_topic * style_factor * topic_count
    total_hours = total_days * 2.5
    if weekly_hours <= 0:
        weekly_hours = 1
    weeks = max(2, round(total_hours / weekly_hours))
    return weeks


def build_weekly_plan(topics, total_weeks):
    plan = []
    if total_weeks < len(topics):
        for i in range(total_weeks):
            plan.append(f"Week {i+1}: Focus on {topics[i % len(topics)]} + mini review")
        return plan

    weeks_per_topic = max(1, total_weeks // len(topics))
    week_idx = 1

    for topic in topics:
        for _ in range(weeks_per_topic):
            plan.append(f"Week {week_idx}: Study {topic} (concepts + practice)")
            week_idx += 1

    while len(plan) < total_weeks:
        plan.append(f"Week {week_idx}: Project sprint + spaced revision")
        week_idx += 1

    if plan:
        plan[-1] = f"Week {len(plan)}: Capstone project + assessment"

    return plan


def generate_learning_path(goal, current_level, weekly_hours, learning_style, interests):
    if not goal.strip():
        raise ValueError("Please enter a learning goal.")

    topics = infer_priority_topics(goal, interests)
    base_days = LEVEL_TO_DAYS[current_level]
    style_factor = STYLE_MULTIPLIER[learning_style]
    weeks = estimate_weeks(weekly_hours, base_days, style_factor, len(topics))
    plan = build_weekly_plan(topics, weeks)

    milestone_interval = max(2, weeks // 4)
    milestones = []
    for i in range(milestone_interval, weeks + 1, milestone_interval):
        milestones.append(f"Week {i}: Milestone checkpoint (quiz + reflection)")

    if not milestones or milestones[-1] != f"Week {weeks}: Milestone checkpoint (quiz + reflection)":
        milestones.append(f"Week {weeks}: Milestone checkpoint (quiz + reflection)")

    recommendations = [
        "Use active recall 3x per week.",
        "Spend at least 1 session/week on debugging real problems.",
        "Keep a progress journal and update it after every study block.",
    ]

    return {
        "topics": topics,
        "weeks": weeks,
        "plan": plan,
        "milestones": milestones,
        "recommendations": recommendations,
    }


class AdaptiveLearningPathApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Adaptive Learning Path Generator")
        self.root.geometry("860x640")
        self.root.minsize(760, 560)

        self.create_widgets()

    def create_widgets(self):
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)

        input_frame = ttk.LabelFrame(container, text="Learner Profile", padding=12)
        input_frame.pack(fill="x")

        ttk.Label(input_frame, text="Learning Goal:").grid(row=0, column=0, sticky="w", pady=4)
        self.goal_entry = ttk.Entry(input_frame, width=70)
        self.goal_entry.grid(row=0, column=1, sticky="ew", pady=4)
        self.goal_entry.insert(0, "Become backend engineer with strong system design")

        ttk.Label(input_frame, text="Current Level:").grid(row=1, column=0, sticky="w", pady=4)
        self.level_var = tk.StringVar(value="Beginner")
        self.level_combo = ttk.Combobox(
            input_frame,
            textvariable=self.level_var,
            values=["Beginner", "Intermediate", "Advanced"],
            state="readonly",
            width=20,
        )
        self.level_combo.grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(input_frame, text="Weekly Hours:").grid(row=2, column=0, sticky="w", pady=4)
        self.hours_var = tk.IntVar(value=8)
        self.hours_spin = ttk.Spinbox(input_frame, from_=1, to=80, textvariable=self.hours_var, width=8)
        self.hours_spin.grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(input_frame, text="Learning Style:").grid(row=3, column=0, sticky="w", pady=4)
        self.style_var = tk.StringVar(value="Balanced")
        self.style_combo = ttk.Combobox(
            input_frame,
            textvariable=self.style_var,
            values=["Balanced", "Hands-on", "Theory-first"],
            state="readonly",
            width=20,
        )
        self.style_combo.grid(row=3, column=1, sticky="w", pady=4)

        ttk.Label(input_frame, text="Interests / Keywords:").grid(row=4, column=0, sticky="nw", pady=4)
        self.interests_text = tk.Text(input_frame, width=52, height=4)
        self.interests_text.grid(row=4, column=1, sticky="ew", pady=4)
        self.interests_text.insert("1.0", "APIs, SQL, distributed systems, debugging")

        input_frame.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(container)
        button_frame.pack(fill="x", pady=10)

        generate_btn = ttk.Button(button_frame, text="Generate Path", command=self.on_generate)
        generate_btn.pack(side="left")

        clear_btn = ttk.Button(button_frame, text="Clear Output", command=self.clear_output)
        clear_btn.pack(side="left", padx=8)

        output_frame = ttk.LabelFrame(container, text="Generated Adaptive Plan", padding=10)
        output_frame.pack(fill="both", expand=True)

        self.output = tk.Text(output_frame, wrap="word", font=("Consolas", 10))
        self.output.pack(fill="both", expand=True)

    def clear_output(self):
        self.output.delete("1.0", "end")

    def on_generate(self):
        goal = self.goal_entry.get()
        current_level = self.level_var.get()
        learning_style = self.style_var.get()
        interests = self.interests_text.get("1.0", "end")

        try:
            weekly_hours = int(self.hours_var.get())
        except (ValueError, tk.TclError):
            messagebox.showerror("Invalid input", "Weekly hours must be a number.")
            return

        try:
            result = generate_learning_path(
                goal=goal,
                current_level=current_level,
                weekly_hours=weekly_hours,
                learning_style=learning_style,
                interests=interests,
            )
        except ValueError as exc:
            messagebox.showerror("Missing input", str(exc))
            return

        lines = []
        lines.append("=== ADAPTIVE LEARNING PATH ===")
        lines.append(f"Goal: {goal}")
        lines.append(f"Current Level: {current_level}")
        lines.append(f"Weekly Hours: {weekly_hours}")
        lines.append(f"Learning Style: {learning_style}")
        lines.append("")

        lines.append("Priority Topics:")
        for idx, topic in enumerate(result["topics"], start=1):
            lines.append(f"  {idx}. {topic}")

        lines.append("")
        lines.append(f"Estimated Duration: {result['weeks']} weeks")
        lines.append("")

        lines.append("Weekly Plan:")
        lines.extend(f"  - {item}" for item in result["plan"])

        lines.append("")
        lines.append("Milestones:")
        lines.extend(f"  - {item}" for item in result["milestones"])

        lines.append("")
        lines.append("Study Recommendations:")
        lines.extend(f"  - {rec}" for rec in result["recommendations"])

        self.output.delete("1.0", "end")
        self.output.insert("1.0", "\n".join(lines))


if __name__ == "__main__":
    root = tk.Tk()
    app = AdaptiveLearningPathApp(root)
    root.mainloop()