# Topic Evolution Tracker - Desktop App in Python
# Purpose:
# Track how discussion topics evolve over time using manually entered records.
# Useful for content analysis, trend monitoring, and classroom demonstrations.

import tkinter as tk
from tkinter import ttk, messagebox
from collections import Counter, defaultdict
from datetime import datetime


def parse_records(raw_text: str):
    """
    Expected row format:
    date_yyyy-mm-dd, topic, source, mentions_count, sentiment_-1_to_1, category
    """
    records = []
    errors = []

    for line_no, line in enumerate(raw_text.strip().splitlines(), start=1):
        if not line.strip():
            continue

        parts = [p.strip() for p in line.split(',')]
        if len(parts) != 6:
            errors.append(f"Line {line_no}: expected 6 comma-separated values")
            continue

        date_str, topic, source, mentions, sentiment, category = parts

        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            errors.append(f"Line {line_no}: invalid date '{date_str}' (use YYYY-MM-DD)")
            continue

        if not topic or not source or not category:
            errors.append(f"Line {line_no}: topic, source, and category cannot be empty")
            continue

        try:
            mentions = int(mentions)
            if mentions < 0:
                raise ValueError
        except ValueError:
            errors.append(f"Line {line_no}: mentions_count must be a non-negative integer")
            continue

        try:
            sentiment = float(sentiment)
            if sentiment < -1 or sentiment > 1:
                raise ValueError
        except ValueError:
            errors.append(f"Line {line_no}: sentiment must be between -1 and 1")
            continue

        records.append(
            {
                "date": date_obj,
                "topic": topic,
                "source": source,
                "mentions": mentions,
                "sentiment": sentiment,
                "category": category,
            }
        )

    return records, errors


def classify_trend(earliest_mentions, latest_mentions):
    if earliest_mentions == 0 and latest_mentions > 0:
        return "Emerging"

    delta = latest_mentions - earliest_mentions
    if delta >= 10:
        return "Rapid Growth"
    if delta >= 1:
        return "Growing"
    if delta == 0:
        return "Stable"
    if delta <= -10:
        return "Sharp Decline"
    return "Declining"


def analyze_topics():
    raw = input_text.get("1.0", tk.END)
    records, errors = parse_records(raw)

    output_text.delete("1.0", tk.END)
    topic_tree.delete(*topic_tree.get_children())

    if errors:
        messagebox.showwarning(
            "Input Warnings",
            "Some rows were skipped due to errors:\n\n" + "\n".join(errors[:10])
            + ("\n..." if len(errors) > 10 else ""),
        )

    if not records:
        messagebox.showerror("No Valid Data", "Please enter at least one valid record row.")
        return

    topic_mentions = Counter()
    topic_sent_sum = defaultdict(float)
    topic_count = Counter()
    source_counter = Counter()
    category_counter = Counter()
    date_topic_mentions = defaultdict(lambda: defaultdict(int))

    for r in records:
        topic_mentions[r["topic"]] += r["mentions"]
        topic_sent_sum[r["topic"]] += r["sentiment"]
        topic_count[r["topic"]] += 1
        source_counter[r["source"]] += 1
        category_counter[r["category"]] += 1
        date_topic_mentions[r["topic"]][r["date"]] += r["mentions"]

    total_records = len(records)
    unique_topics = len(topic_mentions)
    total_mentions = sum(topic_mentions.values())
    avg_sentiment = sum(r["sentiment"] for r in records) / total_records

    top_topics = topic_mentions.most_common(5)
    top_categories = category_counter.most_common(3)
    top_sources = source_counter.most_common(3)

    summary_lines = [
        "Topic Evolution Analysis",
        "=" * 24,
        f"Total Records: {total_records}",
        f"Unique Topics: {unique_topics}",
        f"Total Mentions: {total_mentions}",
        f"Average Sentiment: {avg_sentiment:.3f}",
        "",
        "Top Topics by Mentions:",
    ]

    for topic, mentions in top_topics:
        avg_topic_sent = topic_sent_sum[topic] / topic_count[topic]
        summary_lines.append(f"  - {topic}: {mentions} mentions, avg sentiment {avg_topic_sent:.3f}")

    summary_lines.append("\nTop Categories:")
    for category, count in top_categories:
        summary_lines.append(f"  - {category}: {count} records")

    summary_lines.append("\nTop Sources:")
    for source, count in top_sources:
        summary_lines.append(f"  - {source}: {count} records")

    output_text.insert(tk.END, "\n".join(summary_lines))

    for topic in sorted(topic_mentions, key=topic_mentions.get, reverse=True):
        timeline = date_topic_mentions[topic]
        sorted_dates = sorted(timeline)
        earliest_date = sorted_dates[0]
        latest_date = sorted_dates[-1]
        earliest_mentions = timeline[earliest_date]
        latest_mentions = timeline[latest_date]
        trend = classify_trend(earliest_mentions, latest_mentions)
        avg_topic_sent = topic_sent_sum[topic] / topic_count[topic]

        topic_tree.insert(
            "",
            tk.END,
            values=(
                topic,
                topic_mentions[topic],
                earliest_date.isoformat(),
                latest_date.isoformat(),
                f"{avg_topic_sent:.3f}",
                trend,
            ),
        )


def load_sample_data():
    sample = """2026-01-02,AI Safety,Research Blog,12,0.45,Technology
2026-01-03,Quantum Chips,Tech News,7,0.30,Hardware
2026-01-05,AI Safety,Podcast,15,0.52,Technology
2026-01-08,Green Datacenters,Industry Report,9,0.61,Sustainability
2026-01-11,Quantum Chips,Research Blog,4,0.10,Hardware
2026-01-12,AI Safety,Social Media,21,0.40,Technology
2026-01-14,Green Datacenters,Tech News,11,0.66,Sustainability
2026-01-16,Edge Robotics,Tech News,6,0.25,Automation
2026-01-20,Edge Robotics,Podcast,14,0.38,Automation
2026-01-21,Quantum Chips,Industry Report,3,-0.05,Hardware
"""
    input_text.delete("1.0", tk.END)
    input_text.insert(tk.END, sample)


root = tk.Tk()
root.title("Topic Evolution Tracker")
root.geometry("1020x740")

style = ttk.Style(root)
style.theme_use("clam")

main = ttk.Frame(root, padding=12)
main.pack(fill=tk.BOTH, expand=True)

header = ttk.Label(main, text="Topic Evolution Tracker", font=("Segoe UI", 16, "bold"))
header.pack(anchor="w", pady=(0, 8))

instructions = ttk.Label(
    main,
    text=(
        "Enter one record per line: "
        "date_yyyy-mm-dd, topic, source, mentions_count, sentiment_-1_to_1, category"
    ),
)
instructions.pack(anchor="w")

input_text = tk.Text(main, height=11, width=125)
input_text.pack(fill=tk.X, pady=(6, 10))

button_row = ttk.Frame(main)
button_row.pack(fill=tk.X, pady=(0, 10))

analyze_btn = ttk.Button(button_row, text="Analyze Evolution", command=analyze_topics)
analyze_btn.pack(side=tk.LEFT)

sample_btn = ttk.Button(button_row, text="Load Sample Data", command=load_sample_data)
sample_btn.pack(side=tk.LEFT, padx=8)

output_label = ttk.Label(main, text="Summary")
output_label.pack(anchor="w")

output_text = tk.Text(main, height=13, width=125)
output_text.pack(fill=tk.BOTH, expand=False, pady=(4, 10))

columns = ("Topic", "Mentions", "First Seen", "Last Seen", "Avg Sentiment", "Trend")
topic_tree = ttk.Treeview(main, columns=columns, show="headings", height=10)

for col in columns:
    topic_tree.heading(col, text=col)

col_widths = {
    "Topic": 220,
    "Mentions": 100,
    "First Seen": 120,
    "Last Seen": 120,
    "Avg Sentiment": 120,
    "Trend": 180,
}

for col in columns:
    topic_tree.column(col, width=col_widths[col], anchor="center")

topic_tree.pack(fill=tk.BOTH, expand=True)

load_sample_data()
root.mainloop()