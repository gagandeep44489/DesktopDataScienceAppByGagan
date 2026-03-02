import re
import tkinter as tk
from tkinter import messagebox


class PodcastTopicClassifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Podcast Topic Classifier")
        self.root.geometry("900x700")
        self.root.resizable(False, False)

        self.topic_keywords = {
            "Technology": {
                "ai", "machine learning", "software", "programming", "code", "startup",
                "cloud", "data", "cybersecurity", "robotics", "app", "blockchain"
            },
            "Business": {
                "market", "finance", "entrepreneur", "investment", "leadership", "sales",
                "strategy", "economy", "startup", "management", "productivity", "brand"
            },
            "Health & Wellness": {
                "fitness", "nutrition", "mental health", "wellness", "sleep", "diet",
                "exercise", "yoga", "therapy", "stress", "meditation", "recovery"
            },
            "Education": {
                "learning", "student", "teacher", "school", "course", "study",
                "exam", "curriculum", "classroom", "university", "tutorial", "skills"
            },
            "Entertainment": {
                "movie", "music", "celebrity", "comedy", "drama", "series",
                "tv", "gaming", "festival", "pop culture", "review", "artist"
            },
            "Science": {
                "space", "physics", "biology", "chemistry", "research", "experiment",
                "genetics", "climate", "astronomy", "neuroscience", "evolution", "scientific"
            },
            "Sports": {
                "football", "basketball", "cricket", "tennis", "athlete", "coach",
                "tournament", "league", "match", "training", "olympics", "sports"
            }
        }

        self.create_widgets()

    def create_widgets(self):
        tk.Label(
            self.root,
            text="Podcast Topic Classifier",
            font=("Arial", 18, "bold")
        ).pack(pady=12)

        tk.Label(
            self.root,
            text="Paste podcast title, summary, or transcript excerpt:",
            font=("Arial", 11)
        ).pack(anchor="w", padx=25)

        self.input_text = tk.Text(self.root, width=105, height=12, wrap="word", font=("Arial", 10))
        self.input_text.pack(padx=25, pady=6)

        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=8)

        tk.Button(button_frame, text="Classify Topic", width=18, command=self.classify_topic).grid(row=0, column=0, padx=8)
        tk.Button(button_frame, text="Load Demo Text", width=18, command=self.load_demo_text).grid(row=0, column=1, padx=8)
        tk.Button(button_frame, text="Reset", width=18, command=self.reset).grid(row=0, column=2, padx=8)

        tk.Label(self.root, text="Classification Result", font=("Arial", 13, "bold")).pack(pady=(12, 4))

        self.result_var = tk.StringVar(value="No classification yet.")
        tk.Label(
            self.root,
            textvariable=self.result_var,
            justify="left",
            font=("Arial", 11),
            fg="#123b7a"
        ).pack(padx=25, anchor="w")

        tk.Label(self.root, text="Keyword Match Breakdown", font=("Arial", 13, "bold")).pack(pady=(12, 4))

        self.breakdown_text = tk.Text(self.root, width=105, height=13, wrap="word", font=("Consolas", 10))
        self.breakdown_text.pack(padx=25, pady=6)
        self.breakdown_text.config(state="disabled")

    @staticmethod
    def tokenize(text):
        return re.findall(r"[a-zA-Z]+(?:\s[a-zA-Z]+)?", text.lower())

    def score_topics(self, tokens):
        scores = {}
        for topic, keywords in self.topic_keywords.items():
            match_count = sum(1 for token in tokens if token in keywords)
            scores[topic] = match_count
        return scores

    def classify_topic(self):
        text = self.input_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showerror("Input Error", "Please enter podcast text to classify.")
            return

        tokens = self.tokenize(text)
        if not tokens:
            messagebox.showerror("Input Error", "No valid words found in the provided text.")
            return

        topic_scores = self.score_topics(tokens)
        total_matches = sum(topic_scores.values())

        top_topic = max(topic_scores, key=topic_scores.get)
        top_score = topic_scores[top_topic]

        if total_matches == 0:
            self.result_var.set(
                "Predicted Topic: General / Unknown\n"
                "Confidence: 0.00%\n"
                "Reason: No topic keywords were detected."
            )
        else:
            confidence = (top_score / total_matches) * 100
            self.result_var.set(
                f"Predicted Topic: {top_topic}\n"
                f"Confidence: {confidence:.2f}%\n"
                f"Total Keyword Matches: {total_matches}"
            )

        sorted_scores = sorted(topic_scores.items(), key=lambda item: item[1], reverse=True)
        breakdown_lines = ["Topic                     Matches", "-" * 36]
        for topic, score in sorted_scores:
            breakdown_lines.append(f"{topic:<25} {score}")

        self.breakdown_text.config(state="normal")
        self.breakdown_text.delete("1.0", tk.END)
        self.breakdown_text.insert(tk.END, "\n".join(breakdown_lines))
        self.breakdown_text.config(state="disabled")

    def load_demo_text(self):
        demo = (
            "In this episode, we discuss the latest advances in AI, cloud software, and "
            "cybersecurity strategy for startup founders. We also cover programming skills "
            "and data product management trends."
        )
        self.input_text.delete("1.0", tk.END)
        self.input_text.insert(tk.END, demo)

    def reset(self):
        self.input_text.delete("1.0", tk.END)
        self.result_var.set("No classification yet.")
        self.breakdown_text.config(state="normal")
        self.breakdown_text.delete("1.0", tk.END)
        self.breakdown_text.config(state="disabled")


if __name__ == "__main__":
    root = tk.Tk()
    app = PodcastTopicClassifierApp(root)
    root.mainloop()