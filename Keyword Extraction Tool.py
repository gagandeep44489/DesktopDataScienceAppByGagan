# Keyword Extraction Tool - Desktop Desktop App in Python
# Features:
# - Input text area
# - Extract keywords using TF-IDF
# - Adjustable number of keywords
# - Export keywords to TXT file
# - Simple, clean Tkinter UI

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from sklearn.feature_extraction.text import TfidfVectorizer
import re

# ---------------- Utility Functions ----------------

def clean_text(text):
    text = text.lower()
    text = re.sub(r"[^a-zA-Z\s]", "", text)
    return text


def extract_keywords(text, top_n):
    if not text.strip():
        return []

    text = clean_text(text)
    vectorizer = TfidfVectorizer(stop_words='english')
    tfidf_matrix = vectorizer.fit_transform([text])
    scores = zip(vectorizer.get_feature_names_out(), tfidf_matrix.toarray()[0])
    sorted_words = sorted(scores, key=lambda x: x[1], reverse=True)
    return sorted_words[:top_n]

# ---------------- GUI Functions ----------------

def run_extraction():
    try:
        top_n = int(keyword_count.get())
        text = text_input.get("1.0", tk.END)
        keywords = extract_keywords(text, top_n)

        output_box.delete("1.0", tk.END)
        for word, score in keywords:
            output_box.insert(tk.END, f"{word}  (score: {score:.4f})\n")

    except ValueError:
        messagebox.showerror("Error", "Please enter a valid number of keywords")


def export_keywords():
    content = output_box.get("1.0", tk.END).strip()
    if not content:
        messagebox.showwarning("Warning", "No keywords to export")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")]
    )
    if file_path:
        with open(file_path, "w") as f:
            f.write(content)
        messagebox.showinfo("Success", "Keywords exported successfully")

# ---------------- Main Window ----------------

root = tk.Tk()
root.title("Keyword Extraction Tool")
root.geometry("850x600")
root.resizable(False, False)

style = ttk.Style(root)
style.theme_use('clam')

# ---------------- Layout ----------------

frame = ttk.Frame(root, padding=20)
frame.pack(fill=tk.BOTH, expand=True)

# Input Label
input_label = ttk.Label(frame, text="Enter Text", font=("Segoe UI", 11, "bold"))
input_label.grid(row=0, column=0, sticky="w")

# Text Input
text_input = tk.Text(frame, height=12, width=90)
text_input.grid(row=1, column=0, columnspan=3, pady=8)

# Keyword Count
count_label = ttk.Label(frame, text="Number of Keywords")
count_label.grid(row=2, column=0, sticky="w", pady=5)

keyword_count = ttk.Entry(frame, width=10)
keyword_count.insert(0, "10")
keyword_count.grid(row=2, column=1, sticky="w")

# Buttons
extract_btn = ttk.Button(frame, text="Extract Keywords", command=run_extraction)
extract_btn.grid(row=2, column=2, padx=10)

export_btn = ttk.Button(frame, text="Export to TXT", command=export_keywords)
export_btn.grid(row=4, column=2, pady=10)

# Output Label
output_label = ttk.Label(frame, text="Extracted Keywords", font=("Segoe UI", 11, "bold"))
output_label.grid(row=3, column=0, sticky="w", pady=5)

# Output Box
output_box = tk.Text(frame, height=12, width=90)
output_box.grid(row=4, column=0, columnspan=2)

root.mainloop()
