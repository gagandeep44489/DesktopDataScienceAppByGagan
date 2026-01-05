import tkinter as tk
from tkinter import ttk, scrolledtext
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.decomposition import LatentDirichletAllocation


class TopicModelingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Topic Modeling Visualizer")
        self.root.geometry("950x600")

        self.create_ui()

    def create_ui(self):
        control_frame = ttk.LabelFrame(self.root, text="Controls", padding=10)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        ttk.Label(control_frame, text="Number of Topics:").pack(anchor="w")
        self.topic_count = tk.IntVar(value=3)
        ttk.Spinbox(control_frame, from_=2, to=10, textvariable=self.topic_count, width=10).pack(pady=5)

        ttk.Button(
            control_frame,
            text="Run Topic Modeling",
            command=self.run_topic_modeling
        ).pack(fill=tk.X, pady=10)

        text_frame = ttk.Frame(self.root, padding=10)
        text_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        ttk.Label(text_frame, text="Enter Documents (one per line):").pack(anchor="w")
        self.text_input = scrolledtext.ScrolledText(text_frame, height=12)
        self.text_input.pack(fill=tk.BOTH, expand=True, pady=5)

        self.figure, self.ax = plt.subplots(figsize=(6, 4))
        self.canvas = FigureCanvasTkAgg(self.figure, master=text_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def run_topic_modeling(self):
        raw_text = self.text_input.get("1.0", tk.END).strip()
        if not raw_text:
            return

        documents = [doc for doc in raw_text.split("\n") if doc.strip()]
        n_topics = self.topic_count.get()

        vectorizer = CountVectorizer(stop_words="english", max_df=0.95, min_df=2)
        doc_term_matrix = vectorizer.fit_transform(documents)

        lda = LatentDirichletAllocation(
            n_components=n_topics,
            random_state=42
        )
        lda.fit(doc_term_matrix)

        feature_names = vectorizer.get_feature_names_out()

        self.ax.clear()

        for topic_idx, topic in enumerate(lda.components_):
            top_indices = topic.argsort()[-5:][::-1]
            top_words = [feature_names[i] for i in top_indices]
            weights = topic[top_indices]

            self.ax.bar(
                [f"T{topic_idx+1}-{w}" for w in top_words],
                weights
            )

        self.ax.set_title("Top Words per Topic")
        self.ax.set_ylabel("Word Importance")
        self.ax.tick_params(axis='x', rotation=45)

        self.canvas.draw()


if __name__ == "__main__":
    root = tk.Tk()
    app = TopicModelingApp(root)
    root.mainloop()
