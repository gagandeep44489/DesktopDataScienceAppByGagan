# Color Histogram Analyzer Desktop App
# Author: Gagandeep Singh
# Purpose: Educational / Computer Vision Tool
# Tech Stack: Python, Tkinter, OpenCV, Pillow, NumPy, Matplotlib

import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import cv2
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

# =====================================
# LOAD & PROCESS IMAGE
# =====================================
def load_image(path):
    img = cv2.imread(path)
    if img is None:
        raise ValueError("Invalid image file")
    img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    return img

# =====================================
# COMPUTE COLOR HISTOGRAM
# =====================================
def compute_histogram(img):
    colors = ('r', 'g', 'b')
    hist_data = {}

    for i, color in enumerate(colors):
        hist = cv2.calcHist([img], [i], None, [256], [0, 256])
        hist = hist.flatten()
        hist_data[color] = hist

    return hist_data

# =====================================
# ANALYZE IMAGE
# =====================================
def analyze_image():
    if not image_path.get():
        messagebox.showerror("Error", "Please upload an image")
        return

    try:
        img = load_image(image_path.get())
        hist_data = compute_histogram(img)
        display_histogram(hist_data)
    except Exception as e:
        messagebox.showerror("Error", str(e))

# =====================================
# DISPLAY HISTOGRAM
# =====================================
def display_histogram(hist_data):
    ax.clear()

    ax.plot(hist_data['r'], label='Red')
    ax.plot(hist_data['g'], label='Green')
    ax.plot(hist_data['b'], label='Blue')

    ax.set_title("Color Histogram (RGB)")
    ax.set_xlabel("Pixel Intensity")
    ax.set_ylabel("Frequency")
    ax.legend()

    canvas.draw()

# =====================================
# UPLOAD IMAGE
# =====================================
def upload_image():
    path = filedialog.askopenfilename(
        filetypes=[("Images", "*.png *.jpg *.jpeg")]
    )

    if path:
        image_path.set(path)
        img = Image.open(path).resize((300, 300))
        img_tk = ImageTk.PhotoImage(img)
        image_label.config(image=img_tk)
        image_label.image = img_tk
        ax.clear()
        canvas.draw()

# =====================================
# UI SETUP
# =====================================
root = tk.Tk()
root.title("Color Histogram Analyzer")
root.geometry("900x600")
root.resizable(False, False)

image_path = tk.StringVar()

header = tk.Label(root, text="Color Histogram Analyzer", font=("Arial", 18, "bold"))
header.pack(pady=10)

main_frame = tk.Frame(root)
main_frame.pack()

# Image Display
image_label = tk.Label(main_frame)
image_label.grid(row=0, column=0, padx=20)

# Histogram Plot
fig, ax = plt.subplots(figsize=(5, 4))
canvas = FigureCanvasTkAgg(fig, master=main_frame)
canvas.get_tk_widget().grid(row=0, column=1, padx=20)

# Buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=15)

tk.Button(button_frame, text="Upload Image", width=18, command=upload_image).grid(row=0, column=0, padx=10)
tk.Button(button_frame, text="Analyze Histogram", width=18, command=analyze_image).grid(row=0, column=1, padx=10)

footer = tk.Label(root, text="Educational Tool – RGB Color Distribution Analysis", fg="gray")
footer.pack(side="bottom", pady=10)

root.mainloop()