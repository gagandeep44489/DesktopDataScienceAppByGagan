import tkinter as tk
from tkinter import filedialog, messagebox
from wordcloud import WordCloud
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
import os

class WordCloudApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word Cloud Generator")
        self.root.geometry("900x600")

        self.text_data = ""
        self.wordcloud_image = None

        self.create_widgets()

    def create_widgets(self):
        # Text input area
        text_frame = tk.Frame(self.root)
        text_frame.pack(pady=10, fill=tk.X)

        tk.Label(text_frame, text="Enter Text:", font=("Arial", 11, "bold")).pack(anchor="w")

        self.text_box = tk.Text(text_frame, height=8, wrap=tk.WORD)
        self.text_box.pack(fill=tk.X, padx=5)

        # Buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="Load Text File", width=18, command=self.load_text).grid(row=0, column=0, padx=5)
        tk.Button(button_frame, text="Generate Word Cloud", width=22, command=self.generate_wordcloud).grid(row=0, column=1, padx=5)
        tk.Button(button_frame, text="Save Image", width=18, command=self.save_image).grid(row=0, column=2, padx=5)

        # Image display area
        image_frame = tk.Frame(self.root, bd=2, relief=tk.SUNKEN)
        image_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.image_label = tk.Label(image_frame)
        self.image_label.pack(expand=True)

    def load_text(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Text Files", "*.txt")]
        )
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as file:
                    content = file.read()
                    self.text_box.delete(1.0, tk.END)
                    self.text_box.insert(tk.END, content)
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def generate_wordcloud(self):
        self.text_data = self.text_box.get(1.0, tk.END).strip()

        if not self.text_data:
            messagebox.showwarning("Warning", "Please enter or load text first.")
            return

        wc = WordCloud(
            width=800,
            height=400,
            background_color="white",
            stopwords=None
        ).generate(self.text_data)

        self.wordcloud_image = wc.to_image()

        tk_image = ImageTk.PhotoImage(self.wordcloud_image)
        self.image_label.config(image=tk_image)
        self.image_label.image = tk_image

    def save_image(self):
        if self.wordcloud_image is None:
            messagebox.showwarning("Warning", "No word cloud to save.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")]
        )

        if file_path:
            self.wordcloud_image.save(file_path)
            messagebox.showinfo("Saved", "Word cloud image saved successfully.")

if __name__ == "__main__":
    root = tk.Tk()
    app = WordCloudApp(root)
    root.mainloop()
