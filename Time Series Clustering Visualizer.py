import tkinter as tk
from tkinter import messagebox
import numpy as np
from sklearn.cluster import KMeans
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# -------------------------
# Generate Sample Time Series
# -------------------------
def generate_time_series(n_series=30, length=50):
    data = []
    for i in range(n_series):
        if i < 10:
            series = np.sin(np.linspace(0, 3, length)) + np.random.normal(0, 0.2, length)
        elif i < 20:
            series = np.cos(np.linspace(0, 3, length)) + np.random.normal(0, 0.2, length)
        else:
            series = np.linspace(0, 1, length) + np.random.normal(0, 0.2, length)
        data.append(series)
    return np.array(data)

data = generate_time_series()

# -------------------------
# Clustering & Visualization
# -------------------------
def cluster_and_plot():
    try:
        k = int(cluster_entry.get())
        if k <= 0:
            raise ValueError

        kmeans = KMeans(n_clusters=k, random_state=42)
        labels = kmeans.fit_predict(data)

        fig.clear()
        for cluster_id in range(k):
            ax = fig.add_subplot(k, 1, cluster_id + 1)
            for i in range(len(data)):
                if labels[i] == cluster_id:
                    ax.plot(data[i], alpha=0.6)
            ax.set_title(f"Cluster {cluster_id + 1}")
            ax.set_ylabel("Value")

        ax.set_xlabel("Time")
        canvas.draw()

    except ValueError:
        messagebox.showerror("Error", "Enter a valid number of clusters")

# -------------------------
# GUI Setup
# -------------------------
root = tk.Tk()
root.title("Time Series Clustering Visualizer")
root.geometry("850x650")
root.resizable(False, False)

title = tk.Label(
    root,
    text="Time Series Clustering Visualizer",
    font=("Arial", 16, "bold")
)
title.pack(pady=10)

control_frame = tk.Frame(root)
control_frame.pack()

tk.Label(
    control_frame,
    text="Number of Clusters:",
    font=("Arial", 11)
).grid(row=0, column=0, padx=10)

cluster_entry = tk.Entry(control_frame, width=10)
cluster_entry.insert(0, "3")
cluster_entry.grid(row=0, column=1)

cluster_btn = tk.Button(
    control_frame,
    text="Run Clustering",
    font=("Arial", 11),
    bg="#3498db",
    fg="white",
    command=cluster_and_plot
)
cluster_btn.grid(row=0, column=2, padx=10)

# -------------------------
# Plot Area
# -------------------------
fig = plt.Figure(figsize=(8, 5), dpi=100)
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack(pady=15)

footer = tk.Label(
    root,
    text="Unsupervised Learning | K-Means | Python Desktop App",
    font=("Arial", 9),
    fg="gray"
)
footer.pack(side="bottom", pady=5)

root.mainloop()
