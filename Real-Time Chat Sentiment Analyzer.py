import re
import tkinter as tk
from collections import Counter, deque
from tkinter import ttk


POSITIVE_WORDS = {
    "amazing", "awesome", "brilliant", "calm", "cool", "delightful", "excellent",
    "fantastic", "good", "great", "happy", "helpful", "impressive", "joy", "kind",
    "love", "nice", "optimistic", "perfect", "positive", "productive", "relaxed",
    "smile", "strong", "success", "support", "thankful", "thrilled", "wonderful",
}

NEGATIVE_WORDS = {
    "angry", "annoyed", "awful", "bad", "boring", "confused", "disappointed", "fail",
    "frustrated", "hate", "horrible", "issue", "lag", "negative", "nervous", "pain",
    "poor", "problem", "sad", "scared", "stressed", "terrible", "tired", "ugly",
    "upset", "weak", "worried", "wrong", "error", "toxic",
}


class SentimentAnalyzerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Real-Time Chat Sentiment Analyzer")
        self.root.geometry("980x620")

        self.chat_messages = []
        self.sentiment_history = deque(maxlen=40)
        self.sentiment_totals = Counter({"Positive": 0, "Neutral": 0, "Negative": 0})

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill="both", expand=True)

        container.columnconfigure(0, weight=3)
        container.columnconfigure(1, weight=2)
        container.rowconfigure(1, weight=1)

        title = ttk.Label(
            container,
            text="Analyze incoming chat messages and monitor sentiment in real time",
            font=("Segoe UI", 12, "bold"),
        )
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        # Left: message entry + chat history
        left = ttk.LabelFrame(container, text="Chat Feed", padding=10)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)

        self.entry_var = tk.StringVar()
        self.message_entry = ttk.Entry(left, textvariable=self.entry_var)
        self.message_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8), pady=(0, 8))
        self.message_entry.bind("<Return>", lambda _event: self.add_message())

        send_button = ttk.Button(left, text="Add Message", command=self.add_message)
        send_button.grid(row=0, column=1, sticky="ew", pady=(0, 8))

        self.chat_list = tk.Listbox(left, font=("Consolas", 10), activestyle="none")
        self.chat_list.grid(row=1, column=0, columnspan=2, sticky="nsew")

        # Right: summary + trend graph
        right = ttk.LabelFrame(container, text="Sentiment Dashboard", padding=10)
        right.grid(row=1, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)

        self.summary_label = ttk.Label(
            right,
            text="No messages yet.",
            justify="left",
            font=("Segoe UI", 10),
        )
        self.summary_label.grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.topics_label = ttk.Label(
            right,
            text="Top words: -",
            justify="left",
            font=("Segoe UI", 10),
        )
        self.topics_label.grid(row=1, column=0, sticky="w", pady=(0, 8))

        self.trend_canvas = tk.Canvas(right, width=360, height=220, bg="white", highlightthickness=1)
        self.trend_canvas.grid(row=2, column=0, sticky="nsew", pady=(4, 8))

        controls = ttk.Frame(right)
        controls.grid(row=3, column=0, sticky="ew")
        controls.columnconfigure(0, weight=1)

        clear_button = ttk.Button(controls, text="Clear Feed", command=self.clear_all)
        clear_button.grid(row=0, column=0, sticky="w")

        sample_button = ttk.Button(controls, text="Load Sample Stream", command=self.load_sample_stream)
        sample_button.grid(row=0, column=1, sticky="e")

    @staticmethod
    def score_message(message: str) -> tuple[str, int]:
        words = re.findall(r"[a-zA-Z']+", message.lower())
        pos_hits = sum(word in POSITIVE_WORDS for word in words)
        neg_hits = sum(word in NEGATIVE_WORDS for word in words)
        score = pos_hits - neg_hits

        if score > 0:
            return "Positive", score
        if score < 0:
            return "Negative", score
        return "Neutral", score

    def add_message(self, text: str | None = None) -> None:
        message = (text if text is not None else self.entry_var.get()).strip()
        if not message:
            return

        sentiment, score = self.score_message(message)
        self.chat_messages.append((message, sentiment, score))
        self.sentiment_history.append(score)
        self.sentiment_totals[sentiment] += 1

        color_prefix = {
            "Positive": "+",
            "Negative": "-",
            "Neutral": "=",
        }[sentiment]
        self.chat_list.insert("end", f"{color_prefix} [{sentiment:<8}] {message}")
        self.chat_list.yview_moveto(1.0)

        self.entry_var.set("")
        self._refresh_dashboard()

    def _refresh_dashboard(self) -> None:
        total = len(self.chat_messages)
        pos = self.sentiment_totals["Positive"]
        neu = self.sentiment_totals["Neutral"]
        neg = self.sentiment_totals["Negative"]

        mood = "Neutral"
        if pos > max(neg, neu):
            mood = "Overall mood is Positive"
        elif neg > max(pos, neu):
            mood = "Overall mood is Negative"
        else:
            mood = "Overall mood is Mixed/Neutral"

        self.summary_label.config(
            text=(
                f"Messages analyzed: {total}\n"
                f"Positive: {pos}    Neutral: {neu}    Negative: {neg}\n"
                f"{mood}"
            )
        )

        words = []
        for message, _sentiment, _score in self.chat_messages:
            words.extend(re.findall(r"[a-zA-Z']+", message.lower()))

        stop_words = {"the", "a", "an", "is", "are", "and", "to", "of", "in", "it", "that", "for", "on", "i", "we", "you"}
        filtered = [w for w in words if len(w) > 2 and w not in stop_words]
        top_words = Counter(filtered).most_common(5)
        topic_text = ", ".join(f"{word} ({count})" for word, count in top_words) if top_words else "-"
        self.topics_label.config(text=f"Top words: {topic_text}")

        self._draw_trend()

    def _draw_trend(self) -> None:
        canvas = self.trend_canvas
        canvas.delete("all")

        width = int(canvas["width"])
        height = int(canvas["height"])
        padding = 20

        # Axes
        canvas.create_line(padding, height // 2, width - padding, height // 2, fill="#777", dash=(3, 2))
        canvas.create_line(padding, padding, padding, height - padding, fill="#222")
        canvas.create_text(padding - 8, padding, text="+", fill="green", font=("Segoe UI", 10, "bold"))
        canvas.create_text(padding - 8, height - padding, text="-", fill="red", font=("Segoe UI", 10, "bold"))
        canvas.create_text(width - 55, 12, text="Sentiment Trend", fill="#444", font=("Segoe UI", 9, "bold"))

        if not self.sentiment_history:
            canvas.create_text(width // 2, height // 2, text="Add messages to plot sentiment", fill="#666")
            return

        points = list(self.sentiment_history)
        max_abs = max(1, max(abs(v) for v in points))
        x_step = (width - 2 * padding) / max(1, len(points) - 1)

        plot_points = []
        for i, value in enumerate(points):
            x = padding + i * x_step
            y_ratio = value / max_abs
            y = (height // 2) - y_ratio * ((height - 2 * padding) / 2)
            plot_points.extend([x, y])
            color = "green" if value > 0 else "red" if value < 0 else "#555"
            canvas.create_oval(x - 3, y - 3, x + 3, y + 3, fill=color, outline=color)

        if len(plot_points) >= 4:
            canvas.create_line(*plot_points, fill="#1f77b4", width=2, smooth=True)

    def clear_all(self) -> None:
        self.chat_messages.clear()
        self.sentiment_history.clear()
        self.sentiment_totals = Counter({"Positive": 0, "Neutral": 0, "Negative": 0})
        self.chat_list.delete(0, "end")
        self.summary_label.config(text="No messages yet.")
        self.topics_label.config(text="Top words: -")
        self._draw_trend()

    def load_sample_stream(self) -> None:
        samples = [
            "This new update is amazing and the UI looks great!",
            "I am frustrated by the lag and this error keeps appearing.",
            "Support was helpful and solved my issue quickly.",
            "The feature is okay, but I am confused about the settings.",
            "Fantastic response time today, love this improvement!",
        ]
        for msg in samples:
            self.add_message(msg)


def main() -> None:
    root = tk.Tk()
    app = SentimentAnalyzerApp(root)
    app._draw_trend()
    root.mainloop()


if __name__ == "__main__":
    main()