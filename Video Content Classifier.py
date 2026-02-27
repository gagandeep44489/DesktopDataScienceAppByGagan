import tkinter as tk
from tkinter import filedialog, messagebox
import cv2
import numpy as np
import threading
from PIL import Image, ImageTk
from tensorflow.keras.applications.mobilenet_v2 import (
    MobileNetV2,
    preprocess_input,
    decode_predictions
)
from tensorflow.keras.preprocessing import image


class VideoClassifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Video Content Classifier")
        self.root.geometry("600x450")
        self.root.resizable(False, False)

        # Load Model Once
        self.model = MobileNetV2(weights="imagenet")

        self.video_path = None

        # Title
        tk.Label(
            root,
            text="Video Content Classifier",
            font=("Arial", 16, "bold")
        ).pack(pady=15)

        # Buttons
        tk.Button(
            root,
            text="Upload Video",
            width=20,
            command=self.load_video
        ).pack(pady=5)

        tk.Button(
            root,
            text="Classify Video",
            width=20,
            command=self.start_classification
        ).pack(pady=5)

        # Result Label
        self.result_label = tk.Label(
            root,
            text="",
            font=("Arial", 12),
            wraplength=500
        )
        self.result_label.pack(pady=20)

        # Status Label
        self.status_label = tk.Label(
            root,
            text="",
            fg="blue"
        )
        self.status_label.pack()

    # ----------------------------

    def load_video(self):
        self.video_path = filedialog.askopenfilename(
            filetypes=[("Video Files", "*.mp4 *.avi *.mov")]
        )

        if self.video_path:
            messagebox.showinfo("Success", "Video loaded successfully!")

    # ----------------------------

    def start_classification(self):
        if not self.video_path:
            messagebox.showerror("Error", "Please upload a video first.")
            return

        self.status_label.config(text="Processing video... Please wait.")
        threading.Thread(target=self.classify_video).start()

    # ----------------------------

    def classify_video(self):
        cap = cv2.VideoCapture(self.video_path)

        if not cap.isOpened():
            self.status_label.config(text="")
            messagebox.showerror("Error", "Unable to open video file.")
            return

        frame_rate = int(cap.get(cv2.CAP_PROP_FPS))
        predictions = []
        count = 0

        while True:
            ret, frame = cap.read()
            if not ret:
                break

            # Sample every 2 seconds
            if frame_rate > 0 and count % (frame_rate * 2) == 0:
                try:
                    frame_resized = cv2.resize(frame, (224, 224))
                    img_array = image.img_to_array(frame_resized)
                    img_array = np.expand_dims(img_array, axis=0)
                    img_array = preprocess_input(img_array)

                    preds = self.model.predict(img_array, verbose=0)
                    decoded = decode_predictions(preds, top=1)[0][0]

                    label = decoded[1]
                    confidence = float(decoded[2])

                    predictions.append((label, confidence))

                except Exception as e:
                    continue

            count += 1

        cap.release()

        self.status_label.config(text="")

        if predictions:
            # Get most frequent label
            labels = [p[0] for p in predictions]
            final_label = max(set(labels), key=labels.count)

            # Average confidence
            confidences = [p[1] for p in predictions if p[0] == final_label]
            avg_conf = sum(confidences) / len(confidences)

            result_text = (
                f"Predicted Video Category: {final_label}\n"
                f"Average Confidence: {avg_conf:.2%}"
            )

            self.result_label.config(text=result_text)

        else:
            self.result_label.config(
                text="Could not classify video."
            )


# ----------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = VideoClassifierApp(root)
    root.mainloop()