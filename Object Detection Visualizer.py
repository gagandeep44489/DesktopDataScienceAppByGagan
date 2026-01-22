import tkinter as tk
from tkinter import filedialog, messagebox
import cv2
import numpy as np
from PIL import Image, ImageTk
from ultralytics import YOLO

class ObjectDetectionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Object Detection Visualizer")
        self.root.geometry("900x600")

        # Load YOLOv8 model
        self.model = YOLO("yolov8n.pt")

        self.image_label = tk.Label(self.root)
        self.image_label.pack(pady=10)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack()

        tk.Button(btn_frame, text="Load Image", command=self.load_image).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="Detect Objects", command=self.detect_objects).grid(row=0, column=1, padx=10)
        tk.Button(btn_frame, text="Exit", command=self.root.quit).grid(row=0, column=2, padx=10)

        self.cv_image = None

    def load_image(self):
        path = filedialog.askopenfilename(
            filetypes=[("Image Files", "*.jpg *.png *.jpeg")]
        )
        if not path:
            return

        self.cv_image = cv2.imread(path)
        self.show_image(self.cv_image)

    def detect_objects(self):
        if self.cv_image is None:
            messagebox.showerror("Error", "Please load an image first")
            return

        results = self.model(self.cv_image)[0]

        for box in results.boxes:
            x1, y1, x2, y2 = map(int, box.xyxy[0])
            conf = float(box.conf[0])
            cls = int(box.cls[0])
            label = self.model.names[cls]

            cv2.rectangle(self.cv_image, (x1, y1), (x2, y2), (0, 255, 0), 2)
            cv2.putText(
                self.cv_image,
                f"{label} {conf:.2f}",
                (x1, y1 - 8),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.6,
                (0, 255, 0),
                2
            )

        self.show_image(self.cv_image)

    def show_image(self, img):
        rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(rgb)
        pil_img = pil_img.resize((800, 450))
        tk_img = ImageTk.PhotoImage(pil_img)
        self.image_label.configure(image=tk_img)
        self.image_label.image = tk_img

if __name__ == "__main__":
    root = tk.Tk()
    app = ObjectDetectionApp(root)
    root.mainloop()
