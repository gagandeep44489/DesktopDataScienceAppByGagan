import tkinter as tk
from tkinter import filedialog, messagebox
import cv2
import numpy as np
from PIL import Image, ImageTk

class ImageSegmentationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image Segmentation Desktop Tool")
        self.root.geometry("1000x600")

        self.original_image = None
        self.cv_image = None

        # UI Frames
        top_frame = tk.Frame(root)
        top_frame.pack(pady=10)

        tk.Button(top_frame, text="Load Image", command=self.load_image).grid(row=0, column=0, padx=10)
        tk.Button(top_frame, text="Apply Segmentation", command=self.segment_image).grid(row=0, column=1, padx=10)
        tk.Button(top_frame, text="Exit", command=root.quit).grid(row=0, column=2, padx=10)

        self.image_frame = tk.Frame(root)
        self.image_frame.pack()

        self.label_original = tk.Label(self.image_frame, text="Original Image")
        self.label_original.grid(row=0, column=0, padx=20)

        self.label_segmented = tk.Label(self.image_frame, text="Segmented Image")
        self.label_segmented.grid(row=0, column=1, padx=20)

        self.panel_original = tk.Label(self.image_frame)
        self.panel_original.grid(row=1, column=0, padx=20)

        self.panel_segmented = tk.Label(self.image_frame)
        self.panel_segmented.grid(row=1, column=1, padx=20)

    def load_image(self):
        path = filedialog.askopenfilename(
            filetypes=[("Image Files", "*.jpg *.png *.jpeg")]
        )
        if not path:
            return

        self.cv_image = cv2.imread(path)
        self.original_image = self.cv_image.copy()
        self.show_image(self.cv_image, self.panel_original)

    def segment_image(self):
        if self.cv_image is None:
            messagebox.showerror("Error", "Please load an image first")
            return

        image = cv2.cvtColor(self.cv_image, cv2.COLOR_BGR2RGB)
        pixel_vals = image.reshape((-1, 3))
        pixel_vals = np.float32(pixel_vals)

        k = 4  # Number of segments
        criteria = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 100, 0.2)
        _, labels, centers = cv2.kmeans(
            pixel_vals, k, None, criteria, 10, cv2.KMEANS_RANDOM_CENTERS
        )

        centers = np.uint8(centers)
        segmented_data = centers[labels.flatten()]
        segmented_image = segmented_data.reshape(image.shape)

        segmented_image = cv2.cvtColor(segmented_image, cv2.COLOR_RGB2BGR)
        self.show_image(segmented_image, self.panel_segmented)

    def show_image(self, img, panel):
        img = cv2.resize(img, (400, 300))
        img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        img_pil = Image.fromarray(img_rgb)
        img_tk = ImageTk.PhotoImage(img_pil)
        panel.configure(image=img_tk)
        panel.image = img_tk

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageSegmentationApp(root)
    root.mainloop()
