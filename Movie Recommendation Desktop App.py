import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# Global dataframe
df = None

# -----------------------------
# Load Movie CSV
# -----------------------------
def load_file():
    global df
    file_path = filedialog.askopenfilename(
        title="Select Movie CSV File",
        filetypes=[("CSV Files", "*.csv")]
    )
    if file_path:
        try:
            df = pd.read_csv(file_path)
            required_cols = {"Title", "Genre", "Keywords", "Rating"}
            if not required_cols.issubset(df.columns):
                messagebox.showerror(
                    "Invalid CSV",
                    "CSV must contain columns: Title, Genre, Keywords, Rating"
                )
                df = None
                return
            messagebox.showinfo("Success", "Movie dataset loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")

# -----------------------------
# Recommend Movies
# -----------------------------
def recommend_movies():
    global df
    if df is None:
        messagebox.showerror("Error", "Please load a movie dataset first.")
        return

    keywords = keyword_entry.get().strip().lower()
    genre = genre_var.get()

    if not keywords:
        messagebox.showwarning("Input Required", "Enter some keywords.")
        return

    # Filter by genre if not "Any"
    filtered = df if genre == "Any" else df[df["Genre"].str.lower() == genre.lower()]

    recommendations = []
    for _, row in filtered.iterrows():
        row_keywords = str(row["Keywords"]).lower()
        score = sum(word in row_keywords for word in keywords.split())
        if score > 0:
            recommendations.append((row["Title"], row["Rating"], score))

    recommendations.sort(key=lambda x: (-x[2], -x[1]))  # Sort by score, then rating

    result_text.delete("1.0", tk.END)
    if recommendations:
        result_text.insert(tk.END, "Recommended Movies:\n\n")
        for title, rating, score in recommendations[:10]:
            result_text.insert(tk.END, f"{title} | Rating: {rating} | Match Score: {score}\n")
    else:
        result_text.insert(tk.END, "No matching movies found.")

# -----------------------------
# GUI Setup
# -----------------------------
root = tk.Tk()
root.title("Movie Recommendation App")
root.geometry("650x500")

tk.Label(root, text="Movie Recommendation Desktop App",
         font=("Arial", 16, "bold")).pack(pady=10)

tk.Button(root, text="Load Movie CSV", command=load_file,
          bg="blue", fg="white", width=25).pack(pady=5)

tk.Label(root, text="Enter keywords (e.g., action superhero Marvel):").pack(pady=5)
keyword_entry = tk.Entry(root, width=50)
keyword_entry.pack(pady=5)

tk.Label(root, text="Select Genre:").pack(pady=5)
genre_var = tk.StringVar()
genre_var.set("Any")
genre_menu = tk.OptionMenu(root, genre_var, "Any", "Action", "Romance", "Sci-Fi", "Comedy", "Drama", "Thriller")
genre_menu.pack(pady=5)

tk.Button(root, text="Recommend Movies", command=recommend_movies,
          bg="green", fg="white", width=25).pack(pady=10)

tk.Label(root, text="Results:", font=("Arial", 12, "bold")).pack(pady=5)
result_text = tk.Text(root, height=15, width=70, bg="#f4f4f4")
result_text.pack(padx=10, pady=5)

root.mainloop()