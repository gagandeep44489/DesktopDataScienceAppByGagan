# OCR Desktop Tool
# Author: Gagandeep Singh
# Purpose: Educational / Productivity Tool
# Tech Stack: Python, Tkinter, Tesseract OCR, Pillow, OpenCV

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk
import pytesseract
import cv2
import os

# ================================
# CONFIGURATION (IMPORTANT)
# ================================
# Update this path if Tesseract is not in your system PATH
# Example (Windows): r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# ================================
# IMAGE PREPROCESSING
# ================================
def preprocess_image(image_path):
    img = cv2.imread(image_path)
    if img is None:
        raise ValueError("Invalid image file")

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    return gray

# ================================
# OCR FUNCTION
# ================================
def perform_ocr():
    if not image_path.get():
        messagebox.showerror("Error", "Please upload an image first")
        return

    try:
        processed_img = preprocess_image(image_path.get())
        text = pytesseract.image_to_string(processed_img)

        text_area.delete("1.0", tk.END)
        text_area.insert(tk.END, text)

    except Exception as e:
        messagebox.showerror("OCR Error", str(e))

# ================================
# UPLOAD IMAGE
# ================================
def upload_image():
    path = filedialog.askopenfilename(
        filetypes=[("Images", "*.png *.jpg *.jpeg *.tiff")]
    )

    if path:
        image_path.set(path)
        img = Image.open(path).resize((300, 300))
        img_tk = ImageTk.PhotoImage(img)
        image_label.config(image=img_tk)
        image_label.image = img_tk
        text_area.delete("1.0", tk.END)

# ================================
# SAVE TEXT
# ================================
def save_text():
    text = text_area.get("1.0", tk.END).strip()
    if not text:
        messagebox.showwarning("Warning", "No text to save")
        return

    file = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text File", "*.txt")]
    )

    if file:
        with open(file, "w", encoding="utf-8") as f:
            f.write(text)
        messagebox.showinfo("Saved", "Text saved successfully")

# ================================
# UI SETUP
# ================================
root = tk.Tk()
root.title("OCR Desktop Tool")
root.geometry("800x600")
root.resizable(False, False)

image_path = tk.StringVar()

header = tk.Label(root, text="OCR Desktop Tool", font=("Arial", 18, "bold"))
header.pack(pady=10)

# Image Panel
panel = tk.Frame(root)
panel.pack()

image_label = tk.Label(panel)
image_label.grid(row=0, column=0, padx=20)

# Text Panel
text_area = scrolledtext.ScrolledText(panel, width=50, height=20)
text_area.grid(row=0, column=1, padx=20)

# Buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=15)

tk.Button(button_frame, text="Upload Image", width=18, command=upload_image).grid(row=0, column=0, padx=10)
tk.Button(button_frame, text="Extract Text (OCR)", width=18, command=perform_ocr).grid(row=0, column=1, padx=10)
tk.Button(button_frame, text="Save Text", width=18, command=save_text).grid(row=0, column=2, padx=10)

footer = tk.Label(root, text="Powered by Tesseract OCR | Educational Tool", fg="gray")
footer.pack(side="bottom", pady=10)

root.mainloop()