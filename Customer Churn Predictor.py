import sys
from pathlib import Path
import numpy as np
import pandas as pd
import joblib

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel,
    QLineEdit, QPushButton, QMessageBox
)

from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LogisticRegression


MODEL_PATH = Path("churn_model.pkl")


# -----------------------------
# Train model if not exists
# -----------------------------
def train_and_save_model():
    # Sample churn dataset (can be replaced with real data)
    data = pd.DataFrame({
        "tenure": [1, 3, 6, 12, 24, 36, 48, 60],
        "monthly_charges": [95, 90, 85, 75, 65, 55, 50, 45],
        "total_charges": [95, 270, 510, 900, 1560, 1980, 2400, 2700],
        "churn": [1, 1, 1, 1, 0, 0, 0, 0]
    })

    X = data[["tenure", "monthly_charges", "total_charges"]]
    y = data["churn"]

    pipeline = Pipeline([
        ("scaler", StandardScaler()),
        ("model", LogisticRegression())
    ])

    pipeline.fit(X, y)
    joblib.dump(pipeline, MODEL_PATH)


# -----------------------------
# Desktop Application
# -----------------------------
class CustomerChurnApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Customer Churn Predictor")
        self.setGeometry(300, 150, 420, 320)

        self.model = joblib.load(MODEL_PATH)

        QLabel("Tenure (months):", self).move(30, 40)
        QLabel("Monthly Charges:", self).move(30, 90)
        QLabel("Total Charges:", self).move(30, 140)

        self.tenure_input = QLineEdit(self)
        self.tenure_input.move(200, 40)

        self.monthly_input = QLineEdit(self)
        self.monthly_input.move(200, 90)

        self.total_input = QLineEdit(self)
        self.total_input.move(200, 140)

        self.predict_btn = QPushButton("Predict Churn", self)
        self.predict_btn.move(140, 190)
        self.predict_btn.clicked.connect(self.predict_churn)

        self.result_label = QLabel("", self)
        self.result_label.move(30, 240)
        self.result_label.resize(360, 40)

    def predict_churn(self):
        try:
            tenure = float(self.tenure_input.text())
            monthly = float(self.monthly_input.text())
            total = float(self.total_input.text())

            X = np.array([[tenure, monthly, total]])
            prob = self.model.predict_proba(X)[0][1]

            if prob >= 0.5:
                result = "Likely to Churn"
            else:
                result = "Likely to Stay"

            self.result_label.setText(
                f"Prediction: {result} | Churn Probability: {prob:.2%}"
            )

        except Exception:
            QMessageBox.warning(
                self,
                "Input Error",
                "Please enter valid numeric values."
            )


# -----------------------------
# Main
# -----------------------------
def main():
    if not MODEL_PATH.exists():
        train_and_save_model()

    app = QApplication(sys.argv)
    window = CustomerChurnApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
