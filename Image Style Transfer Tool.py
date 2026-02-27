import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import tensorflow as tf
import tensorflow_hub as hub
import numpy as np

class StyleTransferApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image Style Transfer Tool")
        self.root.geometry("900x600")

        self.content_path = None
        self.style_path = None
        self.output_image = None

        # Load style transfer model once
        self.model = hub.load(
            "https://tfhub.dev/google/magenta/arbitrary-image-stylization-v1-256/2"
        )

        # UI Elements
        tk.Label(root, text="Image Style Transfer Tool",
                 font=("Arial", 16, "bold")).pack(pady=10)

        tk.Button(root, text="Upload Content Image",
                  command=self.load_content).pack(pady=5)

        tk.Button(root, text="Upload Style Image",
                  command=self.load_style).pack(pady=5)

        tk.Button(root, text="Apply Style Transfer",
                  command=self.apply_style).pack(pady=10)

        tk.Button(root, text="Save Output Image",
                  command=self.save_output).pack(pady=5)

        self.result_label = tk.Label(root)
        self.result_label.pack(pady=20)

    # --------------------------

    def load_content(self):
        self.content_path = filedialog.askopenfilename(
            filetypes=[("Image Files", "*.jpg *.png *.jpeg")]
        )
        if self.content_path:
            messagebox.showinfo("Success", "Content image loaded!")

    def load_style(self):
        self.style_path = filedialog.askopenfilename(
            filetypes=[("Image Files", "*.jpg *.png *.jpeg")]
        )
        if self.style_path:
            messagebox.showinfo("Success", "Style image loaded!")

    # --------------------------

    def load_image(self, path):
        img = Image.open(path).resize((512, 512))
        img = np.array(img) / 255.0
        img = img.astype(np.float32)
        return tf.convert_to_tensor(img)[tf.newaxis, ...]

    # --------------------------

    def apply_style(self):
        if not self.content_path or not self.style_path:
            messagebox.showerror("Error", "Upload both images first.")
            return

        try:
            content_image = self.load_image(self.content_path)
            style_image = self.load_image(self.style_path)

            stylized_image = self.model(content_image, style_image)[0]
            self.output_image = stylized_image.numpy()

            img = Image.fromarray((self.output_image[0] * 255).astype(np.uint8))
            img = img.resize((400, 400))

            tk_img = ImageTk.PhotoImage(img)
            self.result_label.configure(image=tk_img)
            self.result_label.image = tk_img

        except Exception as e:
            messagebox.showerror("Error", str(e))

    # --------------------------

    def save_output(self):
        if self.output_image is None:
            messagebox.showerror("Error", "No output image to save.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")]
        )

        if save_path:
            img = Image.fromarray((self.output_image[0] * 255).astype(np.uint8))
            img.save(save_path)
            messagebox.showinfo("Success", "Image saved successfully!")

# --------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = StyleTransferApp(root)
    root.mainloop()