import tkinter as tk
from tkinter import filedialog, messagebox
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

class DocumentSimilarityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Similarity Finder")
        self.root.geometry("900x600")

        self.create_widgets()

    def create_widgets(self):
        # Document 1
        frame1 = tk.LabelFrame(self.root, text="Document 1")
        frame1.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.text1 = tk.Text(frame1, height=10, wrap=tk.WORD)
        self.text1.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        tk.Button(frame1, text="Load File", command=lambda: self.load_file(self.text1)).pack(pady=3)

        # Document 2
        frame2 = tk.LabelFrame(self.root, text="Document 2")
        frame2.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.text2 = tk.Text(frame2, height=10, wrap=tk.WORD)
        self.text2.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        tk.Button(frame2, text="Load File", command=lambda: self.load_file(self.text2)).pack(pady=3)

        # Action buttons
        action_frame = tk.Frame(self.root)
        action_frame.pack(pady=10)

        tk.Button(action_frame, text="Calculate Similarity", width=25, command=self.calculate_similarity).grid(row=0, column=0, padx=5)
        tk.Button(action_frame, text="Clear", width=15, command=self.clear_text).grid(row=0, column=1, padx=5)

        # Result
        self.result_label = tk.Label(self.root, text="Similarity: -- %", font=("Arial", 14, "bold"))
        self.result_label.pack(pady=10)

    def load_file(self, text_widget):
        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if file_path:
            with open(file_path, "r", encoding="utf-8") as file:
                content = file.read()
                text_widget.delete(1.0, tk.END)
                text_widget.insert(tk.END, content)

    def calculate_similarity(self):
        doc1 = self.text1.get(1.0, tk.END).strip()
        doc2 = self.text2.get(1.0, tk.END).strip()

        if not doc1 or not doc2:
            messagebox.showwarning("Input Required", "Please provide both documents.")
            return

        vectorizer = TfidfVectorizer(stop_words="english")
        tfidf_matrix = vectorizer.fit_transform([doc1, doc2])

        similarity_score = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0]
        similarity_percent = round(similarity_score * 100, 2)

        self.result_label.config(text=f"Similarity: {similarity_percent} %")

    def clear_text(self):
        self.text1.delete(1.0, tk.END)
        self.text2.delete(1.0, tk.END)
        self.result_label.config(text="Similarity: -- %")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentSimilarityApp(root)
    root.mainloop()
