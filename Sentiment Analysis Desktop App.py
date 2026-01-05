import tkinter as tk
from tkinter import ttk, scrolledtext
from textblob import TextBlob


class SentimentAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sentiment Analysis Desktop App")
        self.root.geometry("700x500")

        self.create_ui()

    def create_ui(self):
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            main_frame,
            text="Sentiment Analysis Tool",
            font=("Arial", 14, "bold")
        ).pack(pady=10)

        ttk.Label(main_frame, text="Enter Text Below:").pack(anchor="w")

        self.text_input = scrolledtext.ScrolledText(main_frame, height=10)
        self.text_input.pack(fill=tk.BOTH, expand=True, pady=5)

        ttk.Button(
            main_frame,
            text="Analyze Sentiment",
            command=self.analyze_sentiment
        ).pack(pady=15)

        result_frame = ttk.LabelFrame(main_frame, text="Analysis Result", padding=10)
        result_frame.pack(fill=tk.X, pady=10)

        self.sentiment_label = ttk.Label(result_frame, text="Sentiment: ")
        self.sentiment_label.pack(anchor="w", pady=5)

        self.polarity_label = ttk.Label(result_frame, text="Polarity Score: ")
        self.polarity_label.pack(anchor="w", pady=5)

    def analyze_sentiment(self):
        text = self.text_input.get("1.0", tk.END).strip()
        if not text:
            return

        analysis = TextBlob(text)
        polarity = analysis.sentiment.polarity

        if polarity > 0:
            sentiment = "Positive"
        elif polarity < 0:
            sentiment = "Negative"
        else:
            sentiment = "Neutral"

        self.sentiment_label.config(text=f"Sentiment: {sentiment}")
        self.polarity_label.config(text=f"Polarity Score: {polarity:.2f}")


if __name__ == "__main__":
    root = tk.Tk()
    app = SentimentAnalysisApp(root)
    root.mainloop()
