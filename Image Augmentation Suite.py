import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import cv2
import os
import numpy as np

class ImageAugmentationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image Augmentation Suite")
        self.root.geometry("1000x600")

        self.input_folder = ""
        self.output_folder = ""
        self.images = []
        self.current_image = None
        self.current_index = 0

        # Top Panel - Folder Selection
        top_frame = tk.Frame(root)
        top_frame.pack(pady=10)

        tk.Button(top_frame, text="Select Input Folder", command=self.select_input_folder).grid(row=0, column=0, padx=10)
        tk.Button(top_frame, text="Select Output Folder", command=self.select_output_folder).grid(row=0, column=1, padx=10)

        # Middle Panel - Image Preview
        self.preview_frame = tk.Frame(root)
        self.preview_frame.pack(pady=10)

        self.label_original = tk.Label(self.preview_frame, text="Original Image")
        self.label_original.grid(row=0, column=0, padx=20)

        self.label_augmented = tk.Label(self.preview_frame, text="Augmented Image")
        self.label_augmented.grid(row=0, column=1, padx=20)

        self.panel_original = tk.Label(self.preview_frame)
        self.panel_original.grid(row=1, column=0, padx=20)

        self.panel_augmented = tk.Label(self.preview_frame)
        self.panel_augmented.grid(row=1, column=1, padx=20)

        # Bottom Panel - Augmentation Controls
        bottom_frame = tk.Frame(root)
        bottom_frame.pack(pady=10)

        self.rotate_var = tk.IntVar(value=0)
        tk.Label(bottom_frame, text="Rotate (degrees)").grid(row=0, column=0)
        tk.Scale(bottom_frame, from_=0, to=360, orient=tk.HORIZONTAL, variable=self.rotate_var).grid(row=0, column=1)

        self.flip_var = tk.StringVar(value="None")
        tk.Label(bottom_frame, text="Flip").grid(row=1, column=0)
        tk.OptionMenu(bottom_frame, self.flip_var, "None", "Horizontal", "Vertical", "Both").grid(row=1, column=1)

        tk.Button(bottom_frame, text="Apply Augmentation", command=self.apply_augmentation).grid(row=2, column=0, pady=10)
        tk.Button(bottom_frame, text="Save Augmented Image", command=self.save_image).grid(row=2, column=1, pady=10)
        tk.Button(bottom_frame, text="Next Image", command=self.next_image).grid(row=2, column=2, pady=10)

    def select_input_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_folder = folder
            self.images = [os.path.join(folder, f) for f in os.listdir(folder)
                           if f.lower().endswith((".jpg", ".png", ".jpeg"))]
            if not self.images:
                messagebox.showerror("Error", "No images found in folder")
                return
            self.current_index = 0
            self.load_image(self.images[self.current_index])

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder = folder

    def load_image(self, path):
        self.current_image = cv2.imread(path)
        self.show_image(self.current_image, self.panel_original)
        self.show_image(self.current_image, self.panel_augmented)

    def show_image(self, img, panel):
        img = cv2.resize(img, (400, 300))
        img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        img_pil = Image.fromarray(img_rgb)
        img_tk = ImageTk.PhotoImage(img_pil)
        panel.configure(image=img_tk)
        panel.image = img_tk

    def apply_augmentation(self):
        if self.current_image is None:
            messagebox.showerror("Error", "No image loaded")
            return

        img = self.current_image.copy()

        # Rotation
        angle = self.rotate_var.get()
        if angle != 0:
            h, w = img.shape[:2]
            M = cv2.getRotationMatrix2D((w/2, h/2), angle, 1)
            img = cv2.warpAffine(img, M, (w, h))

        # Flipping
        flip_mode = self.flip_var.get()
        if flip_mode == "Horizontal":
            img = cv2.flip(img, 1)
        elif flip_mode == "Vertical":
            img = cv2.flip(img, 0)
        elif flip_mode == "Both":
            img = cv2.flip(img, -1)

        self.augmented_image = img
        self.show_image(self.augmented_image, self.panel_augmented)

    def save_image(self):
        if not hasattr(self, 'augmented_image'):
            messagebox.showerror("Error", "No augmented image to save")
            return
        if not self.output_folder:
            messagebox.showerror("Error", "Select output folder first")
            return

        filename = os.path.basename(self.images[self.current_index])
        save_path = os.path.join(self.output_folder, f"aug_{filename}")
        cv2.imwrite(save_path, self.augmented_image)
        messagebox.showinfo("Saved", f"Image saved: {save_path}")

    def next_image(self):
        if not self.images:
            return
        self.current_index = (self.current_index + 1) % len(self.images)
        self.load_image(self.images[self.current_index])

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageAugmentationApp(root)
    root.mainloop()
