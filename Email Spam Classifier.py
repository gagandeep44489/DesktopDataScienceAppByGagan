# Email Spam Classifier - Desktop App in Python
# Purpose:
# Classify email text as Spam or Ham (not spam) using a lightweight Naive Bayes model.
# Useful for learning text classification and prototyping filtering logic.

import math
import re
import tkinter as tk
from collections import Counter, defaultdict
from tkinter import messagebox, ttk


def tokenize(text: str):
    """Lowercase and split text into alphanumeric tokens."""
    return re.findall(r"[a-zA-Z0-9']+", text.lower())


class NaiveBayesSpamClassifier:
    def __init__(self):
        self.word_counts = {
            "spam": Counter(),
            "ham": Counter(),
        }
        self.class_doc_counts = Counter()
        self.total_docs = 0
        self.vocab = set()

    def fit(self, documents):
        """
        documents: list of tuples -> (label, text)
        label must be 'spam' or 'ham'
        """
        self.word_counts = {"spam": Counter(), "ham": Counter()}
        self.class_doc_counts = Counter()
        self.total_docs = 0
        self.vocab = set()

        for label, text in documents:
            label_norm = label.strip().lower()
            if label_norm not in {"spam", "ham"}:
                continue
            tokens = tokenize(text)
            self.word_counts[label_norm].update(tokens)
            self.class_doc_counts[label_norm] += 1
            self.total_docs += 1
            self.vocab.update(tokens)

    def _class_log_prob(self, label, tokens):
        if self.total_docs == 0:
            return float("-inf")

        # Prior probability with tiny smoothing for robustness.
        prior = (self.class_doc_counts[label] + 1) / (self.total_docs + 2)
        log_prob = math.log(prior)

        label_token_total = sum(self.word_counts[label].values())
        vocab_size = max(len(self.vocab), 1)

        for token in tokens:
            token_count = self.word_counts[label][token]
            # Laplace smoothing
            likelihood = (token_count + 1) / (label_token_total + vocab_size)
            log_prob += math.log(likelihood)

        return log_prob

    def predict_proba(self, text: str):
        tokens = tokenize(text)
        spam_log = self._class_log_prob("spam", tokens)
        ham_log = self._class_log_prob("ham", tokens)

        # Stable softmax for 2 classes
        max_log = max(spam_log, ham_log)
        spam_exp = math.exp(spam_log - max_log)
        ham_exp = math.exp(ham_log - max_log)
        denom = spam_exp + ham_exp

        return {
            "spam": spam_exp / denom,
            "ham": ham_exp / denom,
        }

    def predict(self, text: str):
        probs = self.predict_proba(text)
        return "spam" if probs["spam"] >= probs["ham"] else "ham"


classifier = NaiveBayesSpamClassifier()


def parse_training_data(raw_text: str):
    """
    Expected format per line:
    label,text
    where label is spam or ham.
    """
    parsed = []
    errors = []

    for line_no, line in enumerate(raw_text.strip().splitlines(), start=1):
        if not line.strip():
            continue

        if "," not in line:
            errors.append(f"Line {line_no}: missing comma separator")
            continue

        label, text = line.split(",", 1)
        label_norm = label.strip().lower()
        text = text.strip()

        if label_norm not in {"spam", "ham"}:
            errors.append(f"Line {line_no}: label must be spam or ham")
            continue

        if not text:
            errors.append(f"Line {line_no}: text cannot be empty")
            continue

        parsed.append((label_norm, text))

    return parsed, errors


def train_model():
    training_raw = training_text.get("1.0", tk.END)
    documents, errors = parse_training_data(training_raw)

    if errors:
        messagebox.showwarning(
            "Training Data Warnings",
            "Some rows were skipped due to errors:\n\n"
            + "\n".join(errors[:10])
            + ("\n..." if len(errors) > 10 else ""),
        )

    if len(documents) < 2:
        messagebox.showerror("Insufficient Data", "Please provide at least 2 valid training rows.")
        return

    class_counts = Counter(label for label, _ in documents)
    if not class_counts["spam"] or not class_counts["ham"]:
        messagebox.showerror("Class Balance Required", "Training data must include both spam and ham.")
        return

    classifier.fit(documents)

    model_summary.delete("1.0", tk.END)
    model_summary.insert(
        tk.END,
        "Model trained successfully\n"
        + "=" * 28
        + "\n"
        + f"Total emails: {classifier.total_docs}\n"
        + f"Spam emails: {classifier.class_doc_counts['spam']}\n"
        + f"Ham emails: {classifier.class_doc_counts['ham']}\n"
        + f"Vocabulary size: {len(classifier.vocab)}\n",
    )


def classify_email():
    email_text = email_input.get("1.0", tk.END).strip()
    if not email_text:
        messagebox.showerror("No Email Text", "Please enter email content to classify.")
        return

    if classifier.total_docs == 0:
        messagebox.showerror("Model Not Trained", "Please train the model first.")
        return

    probs = classifier.predict_proba(email_text)
    label = "SPAM" if probs["spam"] >= probs["ham"] else "HAM"

    result_var.set(
        f"Prediction: {label} | Spam probability: {probs['spam'] * 100:.2f}% | Ham probability: {probs['ham'] * 100:.2f}%"
    )

    token_counts = defaultdict(int)
    for token in tokenize(email_text):
        token_counts[token] += 1

    detail_tree.delete(*detail_tree.get_children())
    for token, count in sorted(token_counts.items(), key=lambda x: x[1], reverse=True):
        spam_strength = classifier.word_counts["spam"][token]
        ham_strength = classifier.word_counts["ham"][token]
        detail_tree.insert("", tk.END, values=(token, count, spam_strength, ham_strength))


def load_sample_training_data():
    sample = """spam,Congratulations! You have won a free prize claim now
spam,Limited offer buy now and get 70 percent discount
ham,Hi team please find the project update attached
ham,Can we schedule the architecture review meeting tomorrow
spam,URGENT your account is suspended click this link immediately
ham,Thank you for your payment the invoice is confirmed
ham,Let's have lunch and discuss the sprint goals
spam,You are selected for a cash reward respond today"""

    training_text.delete("1.0", tk.END)
    training_text.insert(tk.END, sample)


def load_sample_email():
    sample_email = """Congratulations, claim your reward now!
You have been selected for a limited time cash prize offer."""
    email_input.delete("1.0", tk.END)
    email_input.insert(tk.END, sample_email)


root = tk.Tk()
root.title("Email Spam Classifier")
root.geometry("1000x760")

style = ttk.Style(root)
style.theme_use("clam")

main = ttk.Frame(root, padding=12)
main.pack(fill=tk.BOTH, expand=True)

header = ttk.Label(main, text="Email Spam Classifier", font=("Segoe UI", 16, "bold"))
header.pack(anchor="w", pady=(0, 8))

instructions = ttk.Label(
    main,
    text="Training format: one email per line as label,text where label is spam or ham.",
)
instructions.pack(anchor="w")

training_text = tk.Text(main, height=10, width=120)
training_text.pack(fill=tk.X, pady=8)

buttons = ttk.Frame(main)
buttons.pack(fill=tk.X, pady=(0, 8))

ttk.Button(buttons, text="Load Sample Training Data", command=load_sample_training_data).pack(side=tk.LEFT)
ttk.Button(buttons, text="Train Model", command=train_model).pack(side=tk.LEFT, padx=8)

summary_label = ttk.Label(main, text="Model Summary")
summary_label.pack(anchor="w")

model_summary = tk.Text(main, height=6, width=120)
model_summary.pack(fill=tk.X, pady=6)

email_label = ttk.Label(main, text="Email Content to Classify")
email_label.pack(anchor="w")

email_input = tk.Text(main, height=8, width=120)
email_input.pack(fill=tk.X, pady=6)

email_buttons = ttk.Frame(main)
email_buttons.pack(fill=tk.X, pady=(0, 8))

ttk.Button(email_buttons, text="Load Sample Email", command=load_sample_email).pack(side=tk.LEFT)
ttk.Button(email_buttons, text="Classify Email", command=classify_email).pack(side=tk.LEFT, padx=8)

result_var = tk.StringVar(value="Prediction: (not classified)")
result_label = ttk.Label(main, textvariable=result_var, font=("Segoe UI", 11, "bold"))
result_label.pack(anchor="w", pady=(4, 8))

token_label = ttk.Label(main, text="Token Evidence (from training data)")
token_label.pack(anchor="w", pady=(4, 2))

columns = ("token", "count_in_email", "spam_count_in_training", "ham_count_in_training")
detail_tree = ttk.Treeview(main, columns=columns, show="headings", height=12)

for col, width in [
    ("token", 250),
    ("count_in_email", 120),
    ("spam_count_in_training", 170),
    ("ham_count_in_training", 170),
]:
    detail_tree.heading(col, text=col.replace("_", " ").title())
    detail_tree.column(col, width=width, anchor=tk.CENTER)

detail_tree.pack(fill=tk.BOTH, expand=True)

root.mainloop()