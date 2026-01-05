import tkinter as tk
from tkinter import ttk, scrolledtext
import re
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.stem import WordNetLemmatizer

# Uncomment first time only
# nltk.download('punkt')
# nltk.download('stopwords')
# nltk.download('wordnet')

class TextPreprocessingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Text Preprocessing Pipeline Tool")
        self.root.geometry("900x600")

        self.stop_words = set(stopwords.words("english"))
        self.lemmatizer = WordNetLemmatizer()

        self.create_ui()

    def create_ui(self):
        options_frame = ttk.LabelFrame(self.root, text="Preprocessing Steps", padding=10)
        options_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        self.lowercase = tk.BooleanVar(value=True)
        self.remove_punct = tk.BooleanVar(value=True)
        self.remove_stopwords = tk.BooleanVar(value=True)
        self.lemmatize = tk.BooleanVar(value=True)

        ttk.Checkbutton(options_frame, text="Lowercase", variable=self.lowercase).pack(anchor="w")
        ttk.Checkbutton(options_frame, text="Remove Punctuation", variable=self.remove_punct).pack(anchor="w")
        ttk.Checkbutton(options_frame, text="Remove Stopwords", variable=self.remove_stopwords).pack(anchor="w")
        ttk.Checkbutton(options_frame, text="Lemmatization", variable=self.lemmatize).pack(anchor="w")

        ttk.Button(options_frame, text="Process Text", command=self.process_text).pack(fill=tk.X, pady=10)

        text_frame = ttk.Frame(self.root, padding=10)
        text_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        ttk.Label(text_frame, text="Input Text").pack(anchor="w")
        self.input_text = scrolledtext.ScrolledText(text_frame, height=10)
        self.input_text.pack(fill=tk.BOTH, expand=True, pady=5)

        ttk.Label(text_frame, text="Processed Output").pack(anchor="w")
        self.output_text = scrolledtext.ScrolledText(text_frame, height=10)
        self.output_text.pack(fill=tk.BOTH, expand=True, pady=5)

    def process_text(self):
        text = self.input_text.get("1.0", tk.END)

        if self.lowercase.get():
            text = text.lower()

        if self.remove_punct.get():
            text = re.sub(r"[^\w\s]", "", text)

        tokens = word_tokenize(text)

        if self.remove_stopwords.get():
            tokens = [t for t in tokens if t not in self.stop_words]

        if self.lemmatize.get():
            tokens = [self.lemmatizer.lemmatize(t) for t in tokens]

        result = " ".join(tokens)

        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, result)


if __name__ == "__main__":
    root = tk.Tk()
    app = TextPreprocessingApp(root)
    root.mainloop()
