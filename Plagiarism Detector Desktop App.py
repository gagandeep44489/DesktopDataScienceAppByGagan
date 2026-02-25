import tkinter as tk
from tkinter import messagebox
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

def check_plagiarism():
    text1 = text_input1.get("1.0", tk.END).strip()
    text2 = text_input2.get("1.0", tk.END).strip()

    if not text1 or not text2:
        messagebox.showwarning("Input Error", "Please enter both texts!")
        return

    try:
        vectorizer = TfidfVectorizer()
        tfidf_matrix = vectorizer.fit_transform([text1, text2])

        similarity = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])
        percentage = similarity[0][0] * 100

        result_label.config(
            text=f"Similarity Score: {percentage:.2f}%"
        )

        if percentage > 80:
            status_label.config(text="High Plagiarism Risk ⚠", fg="red")
        elif percentage > 50:
            status_label.config(text="Moderate Similarity", fg="orange")
        else:
            status_label.config(text="Low Similarity ✅", fg="green")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# GUI Setup
root = tk.Tk()
root.title("Plagiarism Detector")
root.geometry("700x600")
root.resizable(False, False)

tk.Label(root, text="Plagiarism Detector",
         font=("Arial", 16, "bold")).pack(pady=10)

tk.Label(root, text="Enter Text 1").pack()
text_input1 = tk.Text(root, height=8, width=80)
text_input1.pack(pady=5)

tk.Label(root, text="Enter Text 2").pack()
text_input2 = tk.Text(root, height=8, width=80)
text_input2.pack(pady=5)

tk.Button(root, text="Check Similarity",
          command=check_plagiarism,
          bg="blue", fg="white").pack(pady=15)

result_label = tk.Label(root, text="", font=("Arial", 14, "bold"))
result_label.pack(pady=5)

status_label = tk.Label(root, text="", font=("Arial", 12, "bold"))
status_label.pack(pady=5)

root.mainloop()