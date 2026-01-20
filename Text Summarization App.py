import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

from nltk.corpus import stopwords
from nltk.tokenize import sent_tokenize, word_tokenize
from collections import defaultdict
import heapq
import re

# Ensure required NLTK resources are available
try:
    stop_words = set(stopwords.words("english"))
except LookupError:
    nltk.download("punkt")
    nltk.download("stopwords")
    stop_words = set(stopwords.words("english"))


class TextSummarizationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Text Summarization App")
        self.root.geometry("900x600")

        self.create_widgets()

    def create_widgets(self):
        title = ttk.Label(
            self.root,
            text="Text Summarization Desktop App",
            font=("Arial", 18, "bold")
        )
        title.pack(pady=10)

        input_label = ttk.Label(self.root, text="Input Text:")
        input_label.pack(anchor="w", padx=10)

        self.input_text = scrolledtext.ScrolledText(
            self.root, height=12, wrap=tk.WORD
        )
        self.input_text.pack(fill="both", padx=10, pady=5, expand=True)

        controls_frame = ttk.Frame(self.root)
        controls_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(controls_frame, text="Summary Sentences:").pack(side="left")

        self.sentences_var = tk.IntVar(value=3)
        self.sentences_spinbox = ttk.Spinbox(
            controls_frame,
            from_=1,
            to=10,
            width=5,
            textvariable=self.sentences_var
        )
        self.sentences_spinbox.pack(side="left", padx=5)

        summarize_btn = ttk.Button(
            controls_frame,
            text="Generate Summary",
            command=self.generate_summary
        )
        summarize_btn.pack(side="right")

        output_label = ttk.Label(self.root, text="Summary:")
        output_label.pack(anchor="w", padx=10, pady=(10, 0))

        self.output_text = scrolledtext.ScrolledText(
            self.root, height=10, wrap=tk.WORD
        )
        self.output_text.pack(fill="both", padx=10, pady=5, expand=True)

    def generate_summary(self):
        text = self.input_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("Input Error", "Please enter text to summarize.")
            return

        summary = self.summarize_text(text, self.sentences_var.get())
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, summary)

    def summarize_text(self, text, num_sentences):
        text = re.sub(r"\s+", " ", text)
        sentences = sent_tokenize(text)

        if len(sentences) <= num_sentences:
            return text

        word_frequencies = defaultdict(int)
        words = word_tokenize(text.lower())

        for word in words:
            if word.isalpha() and word not in stop_words:
                word_frequencies[word] += 1

        max_freq = max(word_frequencies.values())
        for word in word_frequencies:
            word_frequencies[word] /= max_freq

        sentence_scores = defaultdict(float)
        for sentence in sentences:
            for word in word_tokenize(sentence.lower()):
                if word in word_frequencies:
                    sentence_scores[sentence] += word_frequencies[word]

        summary_sentences = heapq.nlargest(
            num_sentences, sentence_scores, key=sentence_scores.get
        )

        summary = " ".join(summary_sentences)
        return summary


if __name__ == "__main__":
    root = tk.Tk()
    app = TextSummarizationApp(root)
    root.mainloop()
