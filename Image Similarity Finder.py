# Image Similarity Finder Desktop App
# Author: Gagandeep Singh
# Purpose: Educational / Portfolio Project
# Tech Stack: Python, Tkinter, OpenCV, NumPy, Pillow, Scikit-learn

import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import cv2
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
import os

# ======================================
# FEATURE EXTRACTION (Histogram-based)
# ======================================
def extract_features(image_path):
    img = cv2.imread(image_path)
    if img is None:
        raise ValueError("Invalid image")

    img = cv2.resize(img, (256, 256))
    img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

    # Color histogram (8x8x8)
    hist = cv2.calcHist([img], [0, 1, 2], None, [8, 8, 8], [0, 256, 0, 256, 0, 256])
    cv2.normalize(hist, hist)
    return hist.flatten().reshape(1, -1)

# ======================================
# LOAD IMAGE DATABASE
# ======================================
def load_image_database(folder):
    features = []
    paths = []

    for file in os.listdir(folder):
        if file.lower().endswith((".png", ".jpg", ".jpeg")):
            path = os.path.join(folder, file)
            try:
                feat = extract_features(path)
                features.append(feat)
                paths.append(path)
            except:
                pass

    if not features:
        raise ValueError("No valid images found in folder")

    return np.vstack(features), paths

# ======================================
# FIND MOST SIMILAR IMAGE
# ======================================
def find_similarity():
    if not query_image.get() or not dataset_folder.get():
        messagebox.showerror("Error", "Upload query image and select dataset folder")
        return

    try:
        query_feat = extract_features(query_image.get())
        db_features, db_paths = load_image_database(dataset_folder.get())

        similarities = cosine_similarity(query_feat, db_features)[0]
        best_idx = np.argmax(similarities)

        score = similarities[best_idx]
        best_image_path = db_paths[best_idx]

        display_result(best_image_path, score)

    except Exception as e:
        messagebox.showerror("Error", str(e))

# ======================================
# DISPLAY RESULT
# ======================================
def display_result(image_path, score):
    img = Image.open(image_path).resize((250, 250))
    img_tk = ImageTk.PhotoImage(img)
    result_image_label.config(image=img_tk)
    result_image_label.image = img_tk

    result_label.config(text=f"Most Similar Image\nSimilarity Score: {score:.3f}")

# ======================================
# UPLOAD QUERY IMAGE
# ======================================
def upload_query_image():
    path = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg")])
    if path:
        query_image.set(path)
        img = Image.open(path).resize((250, 250))
        img_tk = ImageTk.PhotoImage(img)
        query_image_label.config(image=img_tk)
        query_image_label.image = img_tk

# ======================================
# SELECT DATASET FOLDER
# ======================================
def select_dataset_folder():
    folder = filedialog.askdirectory()
    if folder:
        dataset_folder.set(folder)

# ======================================
# UI SETUP
# ======================================
root = tk.Tk()
root.title("Image Similarity Finder")
root.geometry("800x500")
root.resizable(False, False)

query_image = tk.StringVar()
dataset_folder = tk.StringVar()

header = tk.Label(root, text="Image Similarity Finder", font=("Arial", 18, "bold"))
header.pack(pady=10)

frame = tk.Frame(root)
frame.pack()

# Query Image
query_frame = tk.Frame(frame)
query_frame.grid(row=0, column=0, padx=20)

query_image_label = tk.Label(query_frame)
query_image_label.pack()

tk.Button(query_frame, text="Upload Query Image", command=upload_query_image).pack(pady=5)

# Result Image
result_frame = tk.Frame(frame)
result_frame.grid(row=0, column=1, padx=20)

result_image_label = tk.Label(result_frame)
result_image_label.pack()

result_label = tk.Label(result_frame, text="", font=("Arial", 12, "bold"))
result_label.pack(pady=5)

# Controls
control_frame = tk.Frame(root)
control_frame.pack(pady=15)

tk.Button(control_frame, text="Select Image Dataset Folder", command=select_dataset_folder).grid(row=0, column=0, padx=10)
tk.Button(control_frame, text="Find Similar Image", command=find_similarity).grid(row=0, column=1, padx=10)

footer = tk.Label(root, text="Educational project – Content-based Image Retrieval", fg="gray")
footer.pack(side="bottom", pady=10)

root.mainloop()