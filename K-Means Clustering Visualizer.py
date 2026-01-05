import tkinter as tk
from tkinter import ttk
import numpy as np
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class KMeansVisualizer:
    def __init__(self, root):
        self.root = root
        self.root.title("K-Means Clustering Visualizer")
        self.root.geometry("900x600")

        self.data = None

        self.create_ui()
        self.create_plot()

    def create_ui(self):
        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.pack(side=tk.LEFT, fill=tk.Y)

        ttk.Label(control_frame, text="K-Means Controls", font=("Arial", 12, "bold")).pack(pady=10)

        ttk.Label(control_frame, text="Number of Clusters (K):").pack(anchor="w")
        self.k_value = tk.IntVar(value=3)
        ttk.Spinbox(control_frame, from_=2, to=10, textvariable=self.k_value, width=10).pack(pady=5)

        ttk.Button(control_frame, text="Generate Data", command=self.generate_data).pack(fill=tk.X, pady=10)
        ttk.Button(control_frame, text="Run K-Means", command=self.run_kmeans).pack(fill=tk.X)

    def create_plot(self):
        self.fig, self.ax = plt.subplots(figsize=(6, 5))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.root)
        self.canvas.get_tk_widget().pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    def generate_data(self):
        self.data = np.random.rand(300, 2)
        self.ax.clear()
        self.ax.scatter(self.data[:, 0], self.data[:, 1], c="gray")
        self.ax.set_title("Generated Data Points")
        self.canvas.draw()

    def run_kmeans(self):
        if self.data is None:
            return

        k = self.k_value.get()
        kmeans = KMeans(n_clusters=k, random_state=42)
        labels = kmeans.fit_predict(self.data)
        centroids = kmeans.cluster_centers_

        self.ax.clear()
        self.ax.scatter(self.data[:, 0], self.data[:, 1], c=labels, cmap="viridis")
        self.ax.scatter(centroids[:, 0], centroids[:, 1],
                        c="red", s=200, marker="X")
        self.ax.set_title(f"K-Means Clustering (K={k})")
        self.canvas.draw()


if __name__ == "__main__":
    root = tk.Tk()
    app = KMeansVisualizer(root)
    root.mainloop()
