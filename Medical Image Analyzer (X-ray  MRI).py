# Medical Image Analyzer Desktop App (X-ray / MRI)
# Author: Gagandeep Singh
# Purpose: Educational demo – NOT for medical diagnosis
# Tech Stack: Python, Tkinter, OpenCV, Pillow, NumPy, Scikit-learn

import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import cv2
import numpy as np
from sklearn.linear_model import LogisticRegression

# ==================================================
# ML MODEL (Corrected Feature Size)
# ==================================================
# Image resized to 100x100 → 10,000 features
model = LogisticRegression(max_iter=1000)

# Dummy training data (placeholder)
X_dummy = np.random.rand(20, 10000)  # 20 samples, 10,000 features
y_dummy = np.random.randint(0, 2, 20)

model.fit(X_dummy, y_dummy)

# ==================================================
# FEATURE EXTRACTION
# ==================================================
def extract_features(image_path):
    img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)

    if img is None:
        raise ValueError("Invalid image file")

    img = cv2.resize(img, (100, 100))
    img = img.astype(np.float32) / 255.0

    # Flatten to 1D vector (1, 10000)
    features = img.flatten().reshape(1, -1)
    return features

# ==================================================
# IMAGE ANALYSIS
# ==================================================
def analyze_image():
    if not file_path.get():
        messagebox.showerror("Error", "Please upload an image first")
        return

    try:
        features = extract_features(file_path.get())
        prediction = model.predict(features)[0]
        probability = model.predict_proba(features)[0][prediction]

        if prediction == 1:
            result_label.config(
                text=f"Result: Abnormality Detected\nConfidence: {probability:.2f}",
                fg="red"
            )
        else:
            result_label.config(
                text=f"Result: Normal Scan\nConfidence: {probability:.2f}",
                fg="green"
            )

    except Exception as e:
        messagebox.showerror("Processing Error", str(e))

# ==================================================
# IMAGE UPLOAD
# ==================================================
def upload_image():
    path = filedialog.askopenfilename(
        filetypes=[("Medical Images", "*.png *.jpg *.jpeg")]
    )

    if path:
        file_path.set(path)
        img = Image.open(path).resize((300, 300))
        img_tk = ImageTk.PhotoImage(img)
        image_label.config(image=img_tk)
        image_label.image = img_tk
        result_label.config(text="")

# ==================================================
# UI SETUP
# ==================================================
root = tk.Tk()
root.title("Medical Image Analyzer (X-ray / MRI)")
root.geometry("520x620")
root.resizable(False, False)

file_path = tk.StringVar()

header = tk.Label(
    root,
    text="Medical Image Analyzer",
    font=("Arial", 18, "bold")
)
header.pack(pady=12)

image_label = tk.Label(root)
image_label.pack(pady=10)

upload_btn = tk.Button(
    root,
    text="Upload X-ray / MRI Image",
    width=25,
    command=upload_image
)
upload_btn.pack(pady=8)

analyze_btn = tk.Button(
    root,
    text="Analyze Image",
    width=25,
    command=analyze_image
)
analyze_btn.pack(pady=8)

result_label = tk.Label(
    root,
    text="",
    font=("Arial", 14, "bold")
)
result_label.pack(pady=20)

footer = tk.Label(
    root,
    text="For educational use only – Not for medical diagnosis",
    fg="gray"
)
footer.pack(side="bottom", pady=10)

root.mainloop()