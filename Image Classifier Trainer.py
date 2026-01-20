import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk
import os
import numpy as np
import matplotlib.pyplot as plt
from tensorflow.keras.preprocessing.image import ImageDataGenerator
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Conv2D, MaxPooling2D, Flatten, Dense
from tensorflow.keras.callbacks import Callback

# ---------- Custom Callback to update GUI ----------
class TrainingCallback(Callback):
    def __init__(self, status_label):
        self.status_label = status_label

    def on_epoch_end(self, epoch, logs=None):
        self.status_label.config(
            text=f"Epoch {epoch+1} - loss: {logs['loss']:.4f}, accuracy: {logs['accuracy']:.4f}"
        )
        self.status_label.update_idletasks()


class ImageClassifierTrainerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image Classifier Trainer")
        self.root.geometry("800x600")

        # GUI Elements
        ttk.Label(root, text="Image Classifier Trainer", font=("Arial", 18, "bold")).pack(pady=10)

        self.status_label = ttk.Label(root, text="Status: Waiting", foreground="green")
        self.status_label.pack(pady=5)

        ttk.Button(root, text="Select Dataset Folder", command=self.select_dataset).pack(pady=5)
        ttk.Button(root, text="Train Model", command=self.train_model).pack(pady=5)
        ttk.Button(root, text="Test Image", command=self.test_image).pack(pady=5)

        self.dataset_path = ""
        self.model = None
        self.img_size = (64, 64)

    def select_dataset(self):
        path = filedialog.askdirectory(title="Select Dataset Folder")
        if path:
            self.dataset_path = path
            self.status_label.config(text=f"Dataset selected: {path}", foreground="blue")

    def train_model(self):
        if not self.dataset_path:
            messagebox.showwarning("Warning", "Please select a dataset folder first.")
            return

        self.status_label.config(text="Training started...", foreground="orange")
        self.root.update_idletasks()

        # Image data generator
        datagen = ImageDataGenerator(rescale=1./255, validation_split=0.2)

        train_gen = datagen.flow_from_directory(
            self.dataset_path,
            target_size=self.img_size,
            batch_size=16,
            class_mode='categorical',
            subset='training'
        )

        val_gen = datagen.flow_from_directory(
            self.dataset_path,
            target_size=self.img_size,
            batch_size=16,
            class_mode='categorical',
            subset='validation'
        )

        num_classes = len(train_gen.class_indices)

        # Simple CNN Model
        self.model = Sequential([
            Conv2D(32, (3,3), activation='relu', input_shape=(*self.img_size, 3)),
            MaxPooling2D(2,2),
            Conv2D(64, (3,3), activation='relu'),
            MaxPooling2D(2,2),
            Flatten(),
            Dense(128, activation='relu'),
            Dense(num_classes, activation='softmax')
        ])

        self.model.compile(optimizer='adam', loss='categorical_crossentropy', metrics=['accuracy'])

        # Train with callback
        self.model.fit(train_gen, validation_data=val_gen, epochs=5, callbacks=[TrainingCallback(self.status_label)])

        self.status_label.config(text="Training completed!", foreground="green")
        self.model.save("image_classifier_model.h5")
        messagebox.showinfo("Success", "Model trained and saved as 'image_classifier_model.h5'.")

    def test_image(self):
        if self.model is None:
            messagebox.showwarning("Warning", "Please train the model first.")
            return

        file_path = filedialog.askopenfilename(title="Select Image", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        if not file_path:
            return

        img = Image.open(file_path).resize(self.img_size)
        img_array = np.expand_dims(np.array(img)/255.0, axis=0)

        pred = self.model.predict(img_array)
        class_index = np.argmax(pred)
        class_name = list(self.model.class_indices.keys())[class_index]

        messagebox.showinfo("Prediction", f"The image is classified as: {class_name}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ImageClassifierTrainerApp(root)
    root.mainloop()
