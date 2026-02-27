import tkinter as tk
from tkinter import messagebox
import numpy as np
from textblob import TextBlob
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler

class MemePopularityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Meme Popularity Predictor")
        self.root.geometry("500x400")

        # Train simple demo model
        self.train_model()

        # UI
        tk.Label(root, text="Meme Popularity Predictor",
                 font=("Arial", 16, "bold")).pack(pady=10)

        tk.Label(root, text="Enter Meme Caption:").pack()
        self.caption_entry = tk.Text(root, height=5, width=50)
        self.caption_entry.pack(pady=5)

        tk.Button(root, text="Predict Popularity",
                  command=self.predict).pack(pady=10)

        self.result_label = tk.Label(root, text="", font=("Arial", 12))
        self.result_label.pack(pady=20)

    # ------------------------

    def train_model(self):
        # Synthetic dataset (demo purpose)
        np.random.seed(42)
        caption_length = np.random.randint(5, 150, 200)
        hashtags = np.random.randint(0, 10, 200)
        sentiment = np.random.uniform(-1, 1, 200)

        # Rule-based popularity simulation
        popularity = (
            (caption_length > 20) &
            (hashtags >= 2) &
            (sentiment > 0)
        ).astype(int)

        X = np.column_stack((caption_length, hashtags, sentiment))
        y = popularity

        self.scaler = StandardScaler()
        X_scaled = self.scaler.fit_transform(X)

        self.model = LogisticRegression()
        self.model.fit(X_scaled, y)

    # ------------------------

    def extract_features(self, text):
        caption_length = len(text)
        hashtags = text.count("#")
        sentiment = TextBlob(text).sentiment.polarity
        return np.array([[caption_length, hashtags, sentiment]])

    # ------------------------

    def predict(self):
        caption = self.caption_entry.get("1.0", tk.END).strip()

        if not caption:
            messagebox.showerror("Error", "Please enter a meme caption.")
            return

        features = self.extract_features(caption)
        features_scaled = self.scaler.transform(features)

        prediction = self.model.predict(features_scaled)[0]
        probability = self.model.predict_proba(features_scaled)[0][1]

        if prediction == 1:
            result = f"🔥 High Popularity Expected\nConfidence: {probability:.2%}"
        else:
            result = f"📉 Low Popularity Expected\nConfidence: {1-probability:.2%}"

        self.result_label.config(text=result)

# ------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = MemePopularityApp(root)
    root.mainloop()