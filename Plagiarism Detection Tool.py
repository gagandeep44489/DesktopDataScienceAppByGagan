# 506. Plagiarism Detection Tool - Desktop App in Python
# Purpose:
# Compare two documents for potential plagiarism using text similarity metrics.

import math
import re
import tkinter as tk
from collections import Counter
from difflib import SequenceMatcher
from tkinter import messagebox, ttk


STOPWORDS = {
    "the", "a", "an", "and", "or", "but", "if", "then", "to", "of", "in", "on", "for",
    "with", "is", "are", "was", "were", "be", "been", "being", "as", "at", "by", "from",
    "that", "this", "these", "those", "it", "its", "into", "their", "there", "than", "so",
    "such", "can", "could", "would", "should", "will", "may", "might", "do", "does", "did",
}


def normalize_text(text: str):
    cleaned = re.sub(r"[^a-zA-Z0-9\s]", " ", text.lower())
    tokens = [t for t in cleaned.split() if t and t not in STOPWORDS]
    return tokens


def ngrams(tokens, n=3):
    if len(tokens) < n:
        return []
    return [tuple(tokens[i : i + n]) for i in range(len(tokens) - n + 1)]


def jaccard_similarity(set_a, set_b):
    if not set_a and not set_b:
        return 1.0
    union = set_a | set_b
    if not union:
        return 0.0
    return len(set_a & set_b) / len(union)


def cosine_similarity(counter_a, counter_b):
    if not counter_a or not counter_b:
        return 0.0

    intersection = set(counter_a.keys()) & set(counter_b.keys())
    dot = sum(counter_a[t] * counter_b[t] for t in intersection)

    norm_a = math.sqrt(sum(v * v for v in counter_a.values()))
    norm_b = math.sqrt(sum(v * v for v in counter_b.values()))

    if norm_a == 0 or norm_b == 0:
        return 0.0
    return dot / (norm_a * norm_b)


def find_matching_segments(text_a: str, text_b: str, min_chars=22, max_segments=12):
    matcher = SequenceMatcher(None, text_a.lower(), text_b.lower())
    blocks = matcher.get_matching_blocks()

    segments = []
    for block in blocks:
        if block.size >= min_chars:
            segment = text_a[block.a : block.a + block.size].strip()
            if segment and segment not in segments:
                segments.append(segment)
        if len(segments) >= max_segments:
            break

    return segments


def classify_risk(weighted_similarity):
    if weighted_similarity >= 0.75:
        return "High Risk"
    if weighted_similarity >= 0.50:
        return "Moderate Risk"
    if weighted_similarity >= 0.30:
        return "Low Risk"
    return "Minimal Risk"


def analyze_plagiarism():
    text_a = source_text.get("1.0", tk.END).strip()
    text_b = suspect_text.get("1.0", tk.END).strip()

    if not text_a or not text_b:
        messagebox.showwarning("Missing Text", "Please provide both Source and Suspect text.")
        return

    tokens_a = normalize_text(text_a)
    tokens_b = normalize_text(text_b)

    unigrams_a = Counter(tokens_a)
    unigrams_b = Counter(tokens_b)
    trigrams_a = set(ngrams(tokens_a, n=3))
    trigrams_b = set(ngrams(tokens_b, n=3))

    lexical_jaccard = jaccard_similarity(set(tokens_a), set(tokens_b))
    trigram_jaccard = jaccard_similarity(trigrams_a, trigrams_b)
    cosine = cosine_similarity(unigrams_a, unigrams_b)

    weighted = (0.30 * lexical_jaccard) + (0.40 * trigram_jaccard) + (0.30 * cosine)
    risk = classify_risk(weighted)

    matching_segments = find_matching_segments(text_a, text_b)

    output_text.delete("1.0", tk.END)
    output_text.insert(
        tk.END,
        "Plagiarism Analysis Result\n"
        + "=" * 30
        + f"\nLexical Jaccard Similarity: {lexical_jaccard * 100:.2f}%"
        + f"\nTrigram Jaccard Similarity: {trigram_jaccard * 100:.2f}%"
        + f"\nCosine Similarity (Term Frequency): {cosine * 100:.2f}%"
        + f"\nWeighted Similarity Score: {weighted * 100:.2f}%"
        + f"\nRisk Assessment: {risk}\n\n"
    )

    if matching_segments:
        output_text.insert(tk.END, "Potentially Copied Segments:\n")
        for i, seg in enumerate(matching_segments, start=1):
            cleaned_seg = " ".join(seg.split())
            output_text.insert(tk.END, f"{i}. {cleaned_seg}\n")
    else:
        output_text.insert(tk.END, "No substantial matching segments found.")

    results_tree.delete(*results_tree.get_children())
    results_tree.insert("", tk.END, values=("Lexical Jaccard", f"{lexical_jaccard * 100:.2f}%"))
    results_tree.insert("", tk.END, values=("Trigram Jaccard", f"{trigram_jaccard * 100:.2f}%"))
    results_tree.insert("", tk.END, values=("Cosine Similarity", f"{cosine * 100:.2f}%"))
    results_tree.insert("", tk.END, values=("Weighted Score", f"{weighted * 100:.2f}%"))
    results_tree.insert("", tk.END, values=("Risk", risk))


def load_sample_texts():
    sample_source = (
        "Computer architecture is a set of rules and methods that describe the functionality, "
        "organization, and implementation of computer systems. It defines how hardware "
        "components are designed and how they interact to execute instructions efficiently."
    )

    sample_suspect = (
        "Computer architecture includes the rules and methods used to describe the "
        "organization and implementation of computer systems. It explains how hardware "
        "parts are designed and how they work together to execute instructions efficiently."
    )

    source_text.delete("1.0", tk.END)
    suspect_text.delete("1.0", tk.END)
    output_text.delete("1.0", tk.END)
    results_tree.delete(*results_tree.get_children())

    source_text.insert(tk.END, sample_source)
    suspect_text.insert(tk.END, sample_suspect)


root = tk.Tk()
root.title("506. Plagiarism Detection Tool")
root.geometry("1080x760")

style = ttk.Style(root)
style.theme_use("clam")

main = ttk.Frame(root, padding=12)
main.pack(fill=tk.BOTH, expand=True)

header = ttk.Label(main, text="Plagiarism Detection Tool", font=("Segoe UI", 16, "bold"))
header.pack(anchor="w", pady=(0, 8))

instruction = ttk.Label(
    main,
    text=(
        "Paste original content in Source Text and the document to check in Suspect Text, "
        "then click Analyze Plagiarism."
    ),
)
instruction.pack(anchor="w")

panes = ttk.PanedWindow(main, orient=tk.HORIZONTAL)
panes.pack(fill=tk.BOTH, expand=False, pady=8)

left_frame = ttk.Labelframe(panes, text="Source Text", padding=8)
right_frame = ttk.Labelframe(panes, text="Suspect Text", padding=8)

source_text = tk.Text(left_frame, width=62, height=14, wrap=tk.WORD)
source_text.pack(fill=tk.BOTH, expand=True)

suspect_text = tk.Text(right_frame, width=62, height=14, wrap=tk.WORD)
suspect_text.pack(fill=tk.BOTH, expand=True)

panes.add(left_frame, weight=1)
panes.add(right_frame, weight=1)

button_row = ttk.Frame(main)
button_row.pack(fill=tk.X, pady=(2, 10))

analyze_button = ttk.Button(button_row, text="Analyze Plagiarism", command=analyze_plagiarism)
analyze_button.pack(side=tk.LEFT)

sample_button = ttk.Button(button_row, text="Load Sample Text", command=load_sample_texts)
sample_button.pack(side=tk.LEFT, padx=8)

metric_frame = ttk.Labelframe(main, text="Similarity Metrics", padding=8)
metric_frame.pack(fill=tk.X, pady=(0, 10))

columns = ("metric", "value")
results_tree = ttk.Treeview(metric_frame, columns=columns, show="headings", height=5)
results_tree.heading("metric", text="Metric")
results_tree.heading("value", text="Value")
results_tree.column("metric", width=260, anchor=tk.W)
results_tree.column("value", width=140, anchor=tk.CENTER)
results_tree.pack(fill=tk.X)

output_label = ttk.Label(main, text="Detailed Report")
output_label.pack(anchor="w")

output_text = tk.Text(main, width=130, height=14, wrap=tk.WORD)
output_text.pack(fill=tk.BOTH, expand=True, pady=6)

root.mainloop()