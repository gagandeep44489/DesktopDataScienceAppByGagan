import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from googletrans import Translator
from textblob import TextBlob
import socket


class TranslationSentimentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatic Translation & Sentiment Dashboard")
        self.root.geometry("900x600")

        # Prevent app freeze on slow internet
        socket.setdefaulttimeout(8)
        self.translator = Translator(timeout=8)

        self.create_widgets()

    def create_widgets(self):
        title = ttk.Label(
            self.root,
            text="Automatic Translation & Sentiment Dashboard",
            font=("Arial", 18, "bold")
        )
        title.pack(pady=10)

        ttk.Label(self.root, text="Input Text:").pack(anchor="w", padx=10)
        self.input_text = scrolledtext.ScrolledText(
            self.root, height=10, wrap=tk.WORD
        )
        self.input_text.pack(fill="both", padx=10, pady=5, expand=True)

        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(control_frame, text="Translate To:").pack(side="left")

        self.language_var = tk.StringVar(value="fr")
        self.language_combo = ttk.Combobox(
            control_frame,
            textvariable=self.language_var,
            state="readonly",
            width=18
        )
        self.language_combo["values"] = [
            "fr - French",
            "hi - Hindi",
            "es - Spanish",
            "de - German",
            "zh-cn - Chinese",
            "ja - Japanese"
        ]
        self.language_combo.current(0)
        self.language_combo.pack(side="left", padx=5)

        ttk.Button(
            control_frame,
            text="Analyze",
            command=self.process_text
        ).pack(side="right")

        ttk.Label(self.root, text="Translated Text:").pack(anchor="w", padx=10)
        self.translated_text = scrolledtext.ScrolledText(
            self.root, height=8, wrap=tk.WORD
        )
        self.translated_text.pack(fill="both", padx=10, pady=5, expand=True)

        self.sentiment_label = ttk.Label(
            self.root,
            text="Sentiment: ",
            font=("Arial", 12, "bold")
        )
        self.sentiment_label.pack(pady=10)

        self.status_label = ttk.Label(
            self.root,
            text="Status: Ready",
            foreground="green"
        )
        self.status_label.pack(pady=5)

    def process_text(self):
        text = self.input_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("Input Error", "Please enter text.")
            return

        self.status_label.config(text="Status: Processing...", foreground="blue")
        self.root.update_idletasks()

        # -------- TRANSLATION (SAFE FALLBACK) --------
        lang_code = self.language_var.get().split(" ")[0]
        try:
            translated = self.translator.translate(text, dest=lang_code).text
            self.status_label.config(text="Status: Completed", foreground="green")
        except Exception:
            translated = text  # fallback to original text
            self.status_label.config(
                text="Status: Translation skipped (offline)",
                foreground="orange"
            )

        self.translated_text.delete("1.0", tk.END)
        self.translated_text.insert(tk.END, translated)

        # -------- SENTIMENT ANALYSIS (OFFLINE) --------
        blob = TextBlob(text)
        polarity = blob.sentiment.polarity

        if polarity > 0.05:
            sentiment = "Positive"
        elif polarity < -0.05:
            sentiment = "Negative"
        else:
            sentiment = "Neutral"

        self.sentiment_label.config(
            text=f"Sentiment: {sentiment} (Polarity: {polarity:.2f})"
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = TranslationSentimentApp(root)
    root.mainloop()
