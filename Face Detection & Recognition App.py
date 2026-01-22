import cv2
import os
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

class FaceRecognitionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Face Detection & Recognition App")
        self.root.geometry("900x600")

        self.face_cascade = cv2.CascadeClassifier(
            cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
        )

        self.recognizer = cv2.face.LBPHFaceRecognizer_create()
        self.label_map = {}
        self.trained = False

        self.image_label = tk.Label(root)
        self.image_label.pack(pady=10)

        btn_frame = tk.Frame(root)
        btn_frame.pack()

        tk.Button(btn_frame, text="Train Faces", command=self.train_faces).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="Load Image", command=self.load_image).grid(row=0, column=1, padx=10)
        tk.Button(btn_frame, text="Exit", command=root.quit).grid(row=0, column=2, padx=10)

        self.cv_image = None

    def train_faces(self):
        faces_dir = "faces"
        if not os.path.exists(faces_dir):
            messagebox.showerror("Error", "Faces folder not found")
            return

        images = []
        labels = []
        label_id = 0

        for person in os.listdir(faces_dir):
            person_path = os.path.join(faces_dir, person)
            if not os.path.isdir(person_path):
                continue

            self.label_map[label_id] = person

            for img_name in os.listdir(person_path):
                img_path = os.path.join(person_path, img_name)
                img = cv2.imread(img_path, cv2.IMREAD_GRAYSCALE)
                if img is None:
                    continue

                faces = self.face_cascade.detectMultiScale(img, 1.3, 5)
                for (x, y, w, h) in faces:
                    images.append(img[y:y+h, x:x+w])
                    labels.append(label_id)

            label_id += 1

        if len(images) == 0:
            messagebox.showerror("Error", "No faces found for training")
            return

        self.recognizer.train(images, np.array(labels))
        self.trained = True
        messagebox.showinfo("Success", "Face training completed")

    def load_image(self):
        if not self.trained:
            messagebox.showwarning("Warning", "Please train faces first")
            return

        path = filedialog.askopenfilename(
            filetypes=[("Image Files", "*.jpg *.png *.jpeg")]
        )
        if not path:
            return

        self.cv_image = cv2.imread(path)
        self.detect_and_recognize()
        self.show_image(self.cv_image)

    def detect_and_recognize(self):
        gray = cv2.cvtColor(self.cv_image, cv2.COLOR_BGR2GRAY)
        faces = self.face_cascade.detectMultiScale(gray, 1.3, 5)

        for (x, y, w, h) in faces:
            face = gray[y:y+h, x:x+w]
            label, confidence = self.recognizer.predict(face)

            name = self.label_map.get(label, "Unknown")
            cv2.rectangle(self.cv_image, (x, y), (x+w, y+h), (0, 255, 0), 2)
            cv2.putText(
                self.cv_image,
                f"{name}",
                (x, y - 8),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.8,
                (0, 255, 0),
                2
            )

    def show_image(self, img):
        img = cv2.resize(img, (800, 450))
        rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        pil = Image.fromarray(rgb)
        tk_img = ImageTk.PhotoImage(pil)
        self.image_label.configure(image=tk_img)
        self.image_label.image = tk_img

if __name__ == "__main__":
    root = tk.Tk()
    app = FaceRecognitionApp(root)
    root.mainloop()
