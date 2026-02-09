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


MODEL_PATH = Path("loan_default_model.pkl")


# -----------------------------
# Train model if not exists
# -----------------------------
def train_and_save_model():
    # Simulated loan dataset (replace with real data later)
    data = pd.DataFrame({
        "annual_income": [200000, 300000, 400000, 600000, 800000, 1000000],
        "loan_amount": [300000, 250000, 200000, 200000, 150000, 100000],
        "credit_score": [550, 600, 650, 700, 750, 800],
        "loan_tenure": [60, 48, 36, 36, 24, 12],
        "default": [1, 1, 1, 0, 0, 0]
    })

    X = data[["annual_income", "loan_amount", "credit_score", "loan_tenure"]]
    y = data["default"]

    pipeline = Pipeline([
        ("scaler", StandardScaler()),
        ("model", LogisticRegression())
    ])

    pipeline.fit(X, y)
    joblib.dump(pipeline, MODEL_PATH)


# -----------------------------
# Desktop Application
# -----------------------------
class LoanDefaultRiskApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Loan Default Risk Estimator")
        self.setGeometry(300, 150, 460, 380)

        self.model = joblib.load(MODEL_PATH)

        QLabel("Annual Income (₹):", self).move(30, 40)
        QLabel("Loan Amount (₹):", self).move(30, 90)
        QLabel("Credit Score:", self).move(30, 140)
        QLabel("Loan Tenure (months):", self).move(30, 190)

        self.income_input = QLineEdit(self)
        self.income_input.move(220, 40)

        self.loan_input = QLineEdit(self)
        self.loan_input.move(220, 90)

        self.credit_input = QLineEdit(self)
        self.credit_input.move(220, 140)

        self.tenure_input = QLineEdit(self)
        self.tenure_input.move(220, 190)

        self.predict_btn = QPushButton("Estimate Risk", self)
        self.predict_btn.move(160, 240)
        self.predict_btn.clicked.connect(self.predict_risk)

        self.result_label = QLabel("", self)
        self.result_label.move(30, 290)
        self.result_label.resize(400, 50)

    def predict_risk(self):
        try:
            income = float(self.income_input.text())
            loan = float(self.loan_input.text())
            credit = float(self.credit_input.text())
            tenure = float(self.tenure_input.text())

            X = np.array([[income, loan, credit, tenure]])
            prob = self.model.predict_proba(X)[0][1]

            if prob >= 0.5:
                risk = "High Default Risk"
            else:
                risk = "Low Default Risk"

            self.result_label.setText(
                f"Risk Assessment: {risk} | Default Probability: {prob:.2%}"
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
    window = LoanDefaultRiskApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
