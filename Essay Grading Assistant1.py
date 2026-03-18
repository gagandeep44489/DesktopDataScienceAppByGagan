# Essay Grading Assistant - Desktop App in Python
# Purpose:
# Help teachers and learners grade essays with a rubric-based workflow.
# Includes quick automated signals, manual score override, and feedback generation.

import re
import tkinter as tk
from tkinter import ttk, messagebox


TRANSITIONS = {
    "however",
    "therefore",
    "moreover",
    "furthermore",
    "consequently",
    "meanwhile",
    "additionally",
    "instead",
    "thus",
    "overall",
    "first",
    "second",
    "finally",
    "in conclusion",
}

POSITIVE_WORDS = {
    "clear",
    "strong",
    "effective",
    "insightful",
    "coherent",
    "persuasive",
    "logical",
    "focused",
    "well",
}

NEGATIVE_WORDS = {
    "unclear",
    "weak",
    "confusing",
    "incomplete",
    "vague",
    "disorganized",
    "limited",
    "poor",
}


def split_essays(raw_text: str):
    """
    Split batch input into essays.
    Format:
      Title: <essay title>
      <essay body>
      ---
    The delimiter line '---' separates essays.
    """
    blocks = [b.strip() for b in raw_text.strip().split("\n---\n") if b.strip()]
    essays = []
    errors = []

    for idx, block in enumerate(blocks, start=1):
        lines = [ln.rstrip() for ln in block.splitlines() if ln.strip()]
        if not lines:
            continue

        title = f"Essay {idx}"
        body_start = 0

        if lines[0].lower().startswith("title:"):
            title = lines[0].split(":", 1)[1].strip() or title
            body_start = 1

        body = "\n".join(lines[body_start:]).strip()
        if not body:
            errors.append(f"Block {idx}: essay body is empty")
            continue

        essays.append({"title": title, "body": body})

    if not essays and raw_text.strip():
        errors.append("No valid essays found. Ensure entries are separated with '---'.")

    return essays, errors


def tokenize_words(text: str):
    return re.findall(r"[A-Za-z']+", text)


def split_sentences(text: str):
    sentences = re.split(r"(?<=[.!?])\s+", text.strip())
    return [s for s in sentences if s]


def score_content(essay_text: str):
    words = tokenize_words(essay_text)
    word_count = len(words)

    if word_count >= 450:
        return 10, word_count
    if word_count >= 350:
        return 8.5, word_count
    if word_count >= 250:
        return 7.0, word_count
    if word_count >= 180:
        return 5.5, word_count
    return 4.0, word_count


def score_organization(essay_text: str):
    lower = essay_text.lower()
    paragraph_count = len([p for p in essay_text.split("\n") if p.strip()])
    transition_hits = sum(1 for t in TRANSITIONS if t in lower)

    base = 4.5
    base += min(paragraph_count, 6) * 0.6
    base += min(transition_hits, 5) * 0.5

    return min(base, 10.0), paragraph_count, transition_hits


def score_grammar_style(essay_text: str):
    sentences = split_sentences(essay_text)
    words = tokenize_words(essay_text)

    if not sentences:
        return 4.0, 0.0, 0.0

    avg_sentence_len = len(words) / max(len(sentences), 1)

    long_sentences = sum(1 for s in sentences if len(tokenize_words(s)) > 30)
    short_sentences = sum(1 for s in sentences if len(tokenize_words(s)) < 5)
    punctuation_issues = len(re.findall(r"\s+[,.!?;:]", essay_text))

    balance_penalty = min(long_sentences + short_sentences, 6) * 0.4
    punctuation_penalty = min(punctuation_issues, 8) * 0.25

    score = 9.0 - balance_penalty - punctuation_penalty

    return max(min(score, 10.0), 3.5), avg_sentence_len, punctuation_issues


def score_vocabulary(essay_text: str):
    words = [w.lower() for w in tokenize_words(essay_text)]
    if not words:
        return 3.5, 0.0

    unique_ratio = len(set(words)) / len(words)
    positive_hits = sum(1 for w in words if w in POSITIVE_WORDS)
    negative_hits = sum(1 for w in words if w in NEGATIVE_WORDS)

    score = 4.5 + unique_ratio * 6 + min(positive_hits, 5) * 0.2 - min(negative_hits, 5) * 0.3
    return max(min(score, 10.0), 3.5), unique_ratio


def clamp_0_10(value: float):
    return max(0.0, min(10.0, value))


def parse_weight(name: str, var: tk.StringVar):
    raw = var.get().strip()
    try:
        val = float(raw)
    except ValueError:
        raise ValueError(f"Weight for '{name}' must be numeric")

    if val < 0:
        raise ValueError(f"Weight for '{name}' cannot be negative")

    return val


def calculate_overall(scores, weights):
    total_weight = sum(weights.values())
    if total_weight <= 0:
        raise ValueError("At least one rubric weight must be greater than zero")

    weighted = sum(scores[k] * weights[k] for k in scores)
    return weighted / total_weight


def grade_band(score_10):
    if score_10 >= 9.0:
        return "A"
    if score_10 >= 8.0:
        return "B"
    if score_10 >= 7.0:
        return "C"
    if score_10 >= 6.0:
        return "D"
    return "F"


def feedback_for_score(label, score):
    if score >= 8.5:
        return f"{label}: Excellent performance with minimal revision needed."
    if score >= 7.0:
        return f"{label}: Good work overall; strengthen details for a higher band."
    if score >= 5.5:
        return f"{label}: Developing; revise structure and evidence for clarity."
    return f"{label}: Needs significant improvement and targeted rewriting."


def evaluate_essay(essay_body: str, weights):
    content, word_count = score_content(essay_body)
    org, para_count, transition_hits = score_organization(essay_body)
    grammar, avg_sentence_len, punctuation_issues = score_grammar_style(essay_body)
    vocab, unique_ratio = score_vocabulary(essay_body)

    scores = {
        "Content": clamp_0_10(content),
        "Organization": clamp_0_10(org),
        "Grammar & Style": clamp_0_10(grammar),
        "Vocabulary": clamp_0_10(vocab),
    }

    overall = calculate_overall(scores, weights)

    diagnostics = {
        "Word Count": word_count,
        "Paragraphs": para_count,
        "Transition Hits": transition_hits,
        "Avg Sentence Length": avg_sentence_len,
        "Punctuation Issues": punctuation_issues,
        "Lexical Diversity": unique_ratio,
    }

    return scores, overall, diagnostics


def run_grading():
    raw = input_text.get("1.0", tk.END)
    essays, errors = split_essays(raw)

    output_text.delete("1.0", tk.END)
    detail_tree.delete(*detail_tree.get_children())

    if errors:
        messagebox.showwarning(
            "Input Warnings",
            "Some entries were skipped:\n\n" + "\n".join(errors[:10]) + ("\n..." if len(errors) > 10 else ""),
        )

    if not essays:
        messagebox.showerror("No Valid Essays", "Please provide at least one valid essay block.")
        return

    try:
        weights = {
            "Content": parse_weight("Content", content_weight_var),
            "Organization": parse_weight("Organization", org_weight_var),
            "Grammar & Style": parse_weight("Grammar & Style", grammar_weight_var),
            "Vocabulary": parse_weight("Vocabulary", vocab_weight_var),
        }
    except ValueError as ex:
        messagebox.showerror("Rubric Error", str(ex))
        return

    results = []

    for essay in essays:
        scores, overall, diagnostics = evaluate_essay(essay["body"], weights)
        grade = grade_band(overall)
        results.append(
            {
                "title": essay["title"],
                "scores": scores,
                "overall": overall,
                "grade": grade,
                "diagnostics": diagnostics,
            }
        )

    avg_overall = sum(r["overall"] for r in results) / len(results)

    summary = [
        "Essay Grading Assistant Report",
        "=" * 32,
        f"Essays Processed: {len(results)}",
        f"Average Score: {avg_overall:.2f}/10",
        f"Average Percentage: {avg_overall * 10:.1f}%",
        "",
    ]

    for r in sorted(results, key=lambda x: x["overall"], reverse=True):
        summary.append(
            f"- {r['title']}: {r['overall']:.2f}/10 ({r['overall'] * 10:.1f}%), Grade {r['grade']}"
        )

    output_text.insert(tk.END, "\n".join(summary))

    for r in sorted(results, key=lambda x: x["overall"], reverse=True):
        scores = r["scores"]
        diag = r["diagnostics"]
        feedback = " | ".join(
            [
                feedback_for_score("Content", scores["Content"]),
                feedback_for_score("Organization", scores["Organization"]),
            ]
        )

        detail_tree.insert(
            "",
            tk.END,
            values=(
                r["title"],
                f"{scores['Content']:.1f}",
                f"{scores['Organization']:.1f}",
                f"{scores['Grammar & Style']:.1f}",
                f"{scores['Vocabulary']:.1f}",
                f"{r['overall']:.2f}",
                r["grade"],
                diag["Word Count"],
                f"{diag['Lexical Diversity']:.2f}",
                feedback,
            ),
        )


def load_sample_data():
    sample = """Title: Impact of Technology in Education
Technology has reshaped education by improving access to resources and enabling flexible learning. Students can review lectures online, collaborate through shared documents, and access high-quality references from anywhere.
However, digital access alone does not guarantee deep learning. Teachers must design structured tasks, provide timely feedback, and encourage critical thinking rather than passive browsing.
In conclusion, technology is most effective when paired with sound pedagogy, equity planning, and consistent student support.
---
Title: Why Community Service Matters
Community service helps people develop empathy, teamwork, and civic responsibility. When students volunteer, they meet people from different backgrounds and gain a practical understanding of local challenges.
Moreover, service projects can build communication and leadership skills that are valuable in academic and professional settings. Students learn to plan, listen, and adapt.
Overall, service activities strengthen both communities and the students who participate in them.
---
Title: Should School Uniforms Be Mandatory?
School uniforms can reduce visible economic differences and simplify daily routines for families. They may also improve school identity and reduce peer pressure linked to fashion.
On the other hand, critics argue that uniforms limit self-expression. A balanced policy can allow modest personalization while preserving a respectful dress code.
Therefore, uniform policies should be flexible, inclusive, and developed with student input.
"""
    input_text.delete("1.0", tk.END)
    input_text.insert(tk.END, sample)


root = tk.Tk()
root.title("Essay Grading Assistant")
root.geometry("1260x800")

style = ttk.Style(root)
style.theme_use("clam")

main = ttk.Frame(root, padding=12)
main.pack(fill=tk.BOTH, expand=True)

header = ttk.Label(main, text="Essay Grading Assistant", font=("Segoe UI", 17, "bold"))
header.pack(anchor="w", pady=(0, 8))

instructions = ttk.Label(
    main,
    text=(
        "Enter essay blocks separated by '---'. Optional first line: 'Title: ...'. "
        "Then click Grade Essays."
    ),
)
instructions.pack(anchor="w")

rubric_frame = ttk.LabelFrame(main, text="Rubric Weights", padding=8)
rubric_frame.pack(fill=tk.X, pady=(8, 8))

content_weight_var = tk.StringVar(value="4")
org_weight_var = tk.StringVar(value="3")
grammar_weight_var = tk.StringVar(value="2")
vocab_weight_var = tk.StringVar(value="1")

for i, (label, var) in enumerate(
    [
        ("Content", content_weight_var),
        ("Organization", org_weight_var),
        ("Grammar & Style", grammar_weight_var),
        ("Vocabulary", vocab_weight_var),
    ]
):
    ttk.Label(rubric_frame, text=label).grid(row=0, column=i * 2, padx=(0, 6), pady=2, sticky="w")
    ttk.Entry(rubric_frame, textvariable=var, width=7).grid(
        row=0, column=i * 2 + 1, padx=(0, 14), pady=2, sticky="w"
    )

input_text = tk.Text(main, height=15, width=155, wrap=tk.WORD)
input_text.pack(fill=tk.X, pady=6)

button_row = ttk.Frame(main)
button_row.pack(fill=tk.X, pady=(0, 8))

ttk.Button(button_row, text="Load Sample Essays", command=load_sample_data).pack(side=tk.LEFT)
ttk.Button(button_row, text="Grade Essays", command=run_grading).pack(side=tk.LEFT, padx=8)

output_label = ttk.Label(main, text="Summary")
output_label.pack(anchor="w")

output_text = tk.Text(main, height=10, width=155)
output_text.pack(fill=tk.X, pady=6)

columns = (
    "title",
    "content",
    "org",
    "grammar",
    "vocab",
    "overall",
    "grade",
    "words",
    "lexdiv",
    "feedback",
)

detail_tree = ttk.Treeview(main, columns=columns, show="headings", height=12)

headings = [
    ("title", "Essay Title", 170),
    ("content", "Content", 70),
    ("org", "Organization", 90),
    ("grammar", "Grammar", 80),
    ("vocab", "Vocabulary", 80),
    ("overall", "Overall/10", 80),
    ("grade", "Grade", 65),
    ("words", "Words", 70),
    ("lexdiv", "Lex Div", 70),
    ("feedback", "Auto Feedback", 430),
]

for col, text, width in headings:
    detail_tree.heading(col, text=text)
    detail_tree.column(col, width=width, anchor="center")

detail_tree.column("feedback", anchor="w")
detail_tree.pack(fill=tk.BOTH, expand=True)

footer = ttk.Label(
    main,
    text="Note: Automated scores are heuristics. Use teacher judgment for final grading decisions.",
    foreground="#444",
)
footer.pack(anchor="w", pady=(6, 0))

root.mainloop()