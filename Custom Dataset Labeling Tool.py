import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

class DatasetLabelingTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Custom Dataset Labeling Tool")
        self.root.geometry("800x600")

        self.image_files = []
        self.current_index = 0
        self.labels = {}

        self.create_ui()

    def create_ui(self):
        # Buttons
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10)

        tk.Button(top_frame, text="Load Image Folder", command=self.load_folder).pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="Save Labels", command=self.save_labels).pack(side=tk.LEFT, padx=5)

        # Image Display
        self.image_label = tk.Label(self.root)
        self.image_label.pack(pady=10)

        # Label Entry
        label_frame = tk.Frame(self.root)
        label_frame.pack()

        tk.Label(label_frame, text="Label:").pack(side=tk.LEFT)
        self.label_entry = tk.Entry(label_frame, width=30)
        self.label_entry.pack(side=tk.LEFT, padx=5)

        # Navigation
        nav_frame = tk.Frame(self.root)
        nav_frame.pack(pady=10)

        tk.Button(nav_frame, text="Previous", command=self.prev_image).pack(side=tk.LEFT, padx=10)
        tk.Button(nav_frame, text="Next", command=self.next_image).pack(side=tk.LEFT, padx=10)

    def load_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return

        self.image_files = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith((".png", ".jpg", ".jpeg"))
        ]

        if not self.image_files:
            messagebox.showerror("Error", "No images found!")
            return

        self.current_index = 0
        self.show_image()

    def show_image(self):
        image_path = self.image_files[self.current_index]
        img = Image.open(image_path)
        img = img.resize((400, 400))
        photo = ImageTk.PhotoImage(img)

        self.image_label.config(image=photo)
        self.image_label.image = photo

        self.label_entry.delete(0, tk.END)
        if image_path in self.labels:
            self.label_entry.insert(0, self.labels[image_path])

        self.root.title(f"Labeling: {os.path.basename(image_path)}")

    def save_current_label(self):
        label = self.label_entry.get().strip()
        if label:
            self.labels[self.image_files[self.current_index]] = label

    def next_image(self):
        self.save_current_label()
        if self.current_index < len(self.image_files) - 1:
            self.current_index += 1
            self.show_image()

    def prev_image(self):
        self.save_current_label()
        if self.current_index > 0:
            self.current_index -= 1
            self.show_image()

    def save_labels(self):
        with open("labels.csv", "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["image_path", "label"])
            for img, label in self.labels.items():
                writer.writerow([img, label])

        messagebox.showinfo("Saved", "Labels saved to labels.csv")

if __name__ == "__main__":
    root = tk.Tk()
    app = DatasetLabelingTool(root)
    root.mainloop()
