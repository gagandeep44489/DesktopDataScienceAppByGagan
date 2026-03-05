import math
import random
import re
import tkinter as tk
from collections import Counter
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText


TOKEN_PATTERN = re.compile(r"[A-Za-z0-9']+")


def tokenize(text: str):
    return [token.lower() for token in TOKEN_PATTERN.findall(text)]


def cosine_similarity(vec_a, vec_b):
    dot_product = sum(a * b for a, b in zip(vec_a, vec_b))
    norm_a = math.sqrt(sum(a * a for a in vec_a))
    norm_b = math.sqrt(sum(b * b for b in vec_b))
    if norm_a == 0 or norm_b == 0:
        return 0.0
    return dot_product / (norm_a * norm_b)


def build_tfidf_vectors(documents):
    tokenized_documents = [tokenize(document) for document in documents]
    vocabulary = sorted({token for doc in tokenized_documents for token in doc})
    vocab_index = {token: idx for idx, token in enumerate(vocabulary)}

    doc_freq = Counter()
    for doc in tokenized_documents:
        doc_freq.update(set(doc))

    total_docs = max(1, len(documents))
    idf = {
        token: math.log((1 + total_docs) / (1 + doc_freq[token])) + 1
        for token in vocabulary
    }

    vectors = []
    for doc in tokenized_documents:
        tf = Counter(doc)
        doc_len = max(1, len(doc))
        vector = [0.0] * len(vocabulary)
        for token, count in tf.items():
            idx = vocab_index[token]
            vector[idx] = (count / doc_len) * idf[token]
        vectors.append(vector)

    return vectors, vocabulary


def mean_vector(vectors):
    if not vectors:
        return []
    dimensions = len(vectors[0])
    center = [0.0] * dimensions
    for vector in vectors:
        for idx, value in enumerate(vector):
            center[idx] += value
    return [value / len(vectors) for value in center]


def kmeans(vectors, k, max_iterations=30, seed=42):
    if not vectors:
        return [], []

    k = max(1, min(k, len(vectors)))
    rng = random.Random(seed)
    centroids = [vectors[i][:] for i in rng.sample(range(len(vectors)), k)]
    assignments = [0] * len(vectors)

    for _ in range(max_iterations):
        changed = False

        for idx, vector in enumerate(vectors):
            similarities = [cosine_similarity(vector, centroid) for centroid in centroids]
            best_cluster = max(range(k), key=lambda cluster_idx: similarities[cluster_idx])
            if assignments[idx] != best_cluster:
                assignments[idx] = best_cluster
                changed = True

        if not changed:
            break

        grouped = [[] for _ in range(k)]
        for vector, cluster_idx in zip(vectors, assignments):
            grouped[cluster_idx].append(vector)

        for cluster_idx, cluster_vectors in enumerate(grouped):
            if cluster_vectors:
                centroids[cluster_idx] = mean_vector(cluster_vectors)
            else:
                centroids[cluster_idx] = vectors[rng.randrange(len(vectors))][:]

    return assignments, centroids


def cluster_keywords(doc_vectors, assignments, vocabulary, top_n=6):
    keyword_map = {}
    cluster_count = max(assignments) + 1 if assignments else 0

    for cluster_idx in range(cluster_count):
        cluster_docs = [doc_vectors[idx] for idx, assigned in enumerate(assignments) if assigned == cluster_idx]
        if not cluster_docs:
            keyword_map[cluster_idx] = []
            continue
        center = mean_vector(cluster_docs)
        ranked = sorted(range(len(vocabulary)), key=lambda token_idx: center[token_idx], reverse=True)
        keyword_map[cluster_idx] = [vocabulary[token_idx] for token_idx in ranked[:top_n] if center[token_idx] > 0]

    return keyword_map


class DocumentClusteringExplorer:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Clustering Explorer")
        self.root.geometry("980x640")

        self.documents = []

        self._build_ui()

    def _build_ui(self):
        root_frame = ttk.Frame(self.root, padding=12)
        root_frame.pack(fill=tk.BOTH, expand=True)

        controls = ttk.LabelFrame(root_frame, text="Inputs", padding=10)
        controls.pack(fill=tk.X)

        ttk.Label(controls, text="Cluster Count (k):").grid(row=0, column=0, sticky="w")
        self.k_var = tk.IntVar(value=3)
        ttk.Spinbox(controls, from_=2, to=20, textvariable=self.k_var, width=8).grid(row=0, column=1, padx=(8, 12), sticky="w")

        ttk.Button(controls, text="Load .txt Files", command=self.load_files).grid(row=0, column=2, padx=4)
        ttk.Button(controls, text="Use Text Below", command=self.use_editor_text).grid(row=0, column=3, padx=4)
        ttk.Button(controls, text="Run Clustering", command=self.run_clustering).grid(row=0, column=4, padx=4)
        ttk.Button(controls, text="Clear", command=self.clear_all).grid(row=0, column=5, padx=4)

        help_text = (
            "Paste one document per line in the editor or load multiple .txt files. "
            "Then click 'Run Clustering'."
        )
        ttk.Label(root_frame, text=help_text).pack(anchor="w", pady=(10, 4))

        splitter = ttk.PanedWindow(root_frame, orient=tk.HORIZONTAL)
        splitter.pack(fill=tk.BOTH, expand=True)

        editor_frame = ttk.LabelFrame(splitter, text="Document Editor")
        self.editor = ScrolledText(editor_frame, wrap=tk.WORD, font=("Consolas", 10))
        self.editor.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        splitter.add(editor_frame, weight=1)

        results_frame = ttk.LabelFrame(splitter, text="Clustering Results")
        self.results = ScrolledText(results_frame, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas", 10))
        self.results.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        splitter.add(results_frame, weight=1)

    def set_results(self, text):
        self.results.configure(state=tk.NORMAL)
        self.results.delete("1.0", tk.END)
        self.results.insert(tk.END, text)
        self.results.configure(state=tk.DISABLED)

    def load_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Select Text Files",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
        )
        if not file_paths:
            return

        loaded_documents = []
        for path in file_paths:
            try:
                text = Path(path).read_text(encoding="utf-8").strip()
            except UnicodeDecodeError:
                text = Path(path).read_text(encoding="latin-1").strip()
            if text:
                loaded_documents.append(text)

        if not loaded_documents:
            messagebox.showwarning("No Content", "No readable text content was found in selected files.")
            return

        self.documents = loaded_documents
        self.editor.delete("1.0", tk.END)
        self.editor.insert(tk.END, "\n".join(loaded_documents))
        messagebox.showinfo("Loaded", f"Loaded {len(loaded_documents)} documents from files.")

    def use_editor_text(self):
        raw_text = self.editor.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showwarning("Missing Input", "Please provide one or more documents in the editor.")
            return

        documents = [line.strip() for line in raw_text.splitlines() if line.strip()]
        if len(documents) < 2:
            messagebox.showwarning("Need More Data", "Provide at least two documents (one per line).")
            return

        self.documents = documents
        messagebox.showinfo("Documents Ready", f"Prepared {len(documents)} documents from editor text.")

    def run_clustering(self):
        if not self.documents:
            self.use_editor_text()
            if not self.documents:
                return

        k = self.k_var.get()
        if k < 2:
            messagebox.showwarning("Invalid k", "Cluster count must be at least 2.")
            return

        vectors, vocabulary = build_tfidf_vectors(self.documents)
        if not vocabulary:
            messagebox.showwarning("Empty Vocabulary", "Unable to extract words from documents.")
            return

        assignments, _ = kmeans(vectors, k)
        keywords = cluster_keywords(vectors, assignments, vocabulary)

        cluster_to_docs = {}
        for idx, cluster_idx in enumerate(assignments):
            cluster_to_docs.setdefault(cluster_idx, []).append((idx, self.documents[idx]))

        output_lines = []
        for cluster_idx in sorted(cluster_to_docs):
            cluster_docs = cluster_to_docs[cluster_idx]
            cluster_keywords_text = ", ".join(keywords.get(cluster_idx, [])) or "(no strong keywords)"
            output_lines.append(f"Cluster {cluster_idx + 1} ({len(cluster_docs)} docs)")
            output_lines.append(f"Top keywords: {cluster_keywords_text}")
            for doc_idx, doc_text in cluster_docs:
                preview = doc_text[:120].replace("\n", " ")
                output_lines.append(f"  - Doc {doc_idx + 1}: {preview}")
            output_lines.append("")

        self.set_results("\n".join(output_lines).strip())

    def clear_all(self):
        self.documents = []
        self.editor.delete("1.0", tk.END)
        self.set_results("")


if __name__ == "__main__":
    app_root = tk.Tk()
    DocumentClusteringExplorer(app_root)
    app_root.mainloop()