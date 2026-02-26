import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# Global dataframe
df = None

# -----------------------------
# Load Song CSV
# -----------------------------
def load_file():
    global df
    file_path = filedialog.askopenfilename(
        title="Select Song CSV File",
        filetypes=[("CSV Files", "*.csv")]
    )
    if file_path:
        try:
            df = pd.read_csv(file_path)
            required_cols = {"Title", "Artist", "Genre", "Keywords", "Rating"}
            if not required_cols.issubset(df.columns):
                messagebox.showerror(
                    "Invalid CSV",
                    "CSV must contain columns: Title, Artist, Genre, Keywords, Rating"
                )
                df = None
                return
            messagebox.showinfo("Success", "Music dataset loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")

# -----------------------------
# Predict Playlist
# -----------------------------
def predict_playlist():
    global df
    if df is None:
        messagebox.showerror("Error", "Please load a music dataset first.")
        return

    keywords = keyword_entry.get().strip().lower()
    genre = genre_var.get()

    if not keywords:
        messagebox.showwarning("Input Required", "Enter some keywords.")
        return

    filtered = df if genre == "Any" else df[df["Genre"].str.lower() == genre.lower()]

    recommendations = []
    for _, row in filtered.iterrows():
        row_keywords = str(row["Keywords"]).lower()
        score = sum(word in row_keywords for word in keywords.split())
        if score > 0:
            recommendations.append((row["Title"], row["Artist"], row["Rating"], score))

    recommendations.sort(key=lambda x: (-x[3], -x[2]))  # Sort by score, then rating

    result_text.delete("1.0", tk.END)
    if recommendations:
        result_text.insert(tk.END, "Predicted Playlist:\n\n")
        for title, artist, rating, score in recommendations[:15]:
            result_text.insert(tk.END, f"{title} by {artist} | Rating: {rating} | Match Score: {score}\n")
    else:
        result_text.insert(tk.END, "No matching songs found.")

# -----------------------------
# GUI Setup
# -----------------------------
root = tk.Tk()
root.title("Music Playlist Predictor")
root.geometry("700x550")

tk.Label(root, text="Music Playlist Predictor",
         font=("Arial", 16, "bold")).pack(pady=10)

tk.Button(root, text="Load Music CSV", command=load_file,
          bg="blue", fg="white", width=25).pack(pady=5)

tk.Label(root, text="Enter keywords (mood, genre, artist, etc.):").pack(pady=5)
keyword_entry = tk.Entry(root, width=60)
keyword_entry.pack(pady=5)

tk.Label(root, text="Select Genre:").pack(pady=5)
genre_var = tk.StringVar()
genre_var.set("Any")
genre_menu = tk.OptionMenu(root, genre_var, "Any", "Pop", "Rock", "Hip-Hop", "Jazz", "Classical", "Electronic")
genre_menu.pack(pady=5)

tk.Button(root, text="Predict Playlist", command=predict_playlist,
          bg="green", fg="white", width=25).pack(pady=10)

tk.Label(root, text="Playlist Results:", font=("Arial", 12, "bold")).pack(pady=5)
result_text = tk.Text(root, height=20, width=80, bg="#f4f4f4")
result_text.pack(padx=10, pady=5)

root.mainloop()