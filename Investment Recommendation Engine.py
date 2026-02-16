import tkinter as tk
from tkinter import messagebox
import numpy as np
import matplotlib.pyplot as plt

class InvestmentApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Investment Recommendation Engine")
        self.root.geometry("500x500")

        tk.Label(root, text="Age").pack()
        self.age_entry = tk.Entry(root)
        self.age_entry.pack()

        tk.Label(root, text="Monthly Investment Amount").pack()
        self.amount_entry = tk.Entry(root)
        self.amount_entry.pack()

        tk.Label(root, text="Risk Tolerance (Low/Medium/High)").pack()
        self.risk_entry = tk.Entry(root)
        self.risk_entry.pack()

        tk.Button(root, text="Generate Recommendation",
                  command=self.generate_recommendation).pack(pady=10)

        tk.Button(root, text="Show Portfolio Chart",
                  command=self.show_chart).pack(pady=10)

        self.result_label = tk.Label(root, text="", font=("Arial", 11))
        self.result_label.pack(pady=20)

        self.allocation = None

    def generate_recommendation(self):

        try:
            age = int(self.age_entry.get())
            amount = float(self.amount_entry.get())
            risk = self.risk_entry.get().lower()

            if risk == "low":
                self.allocation = {
                    "Stocks": 30,
                    "Bonds": 40,
                    "Gold": 20,
                    "Cash": 10
                }

            elif risk == "medium":
                self.allocation = {
                    "Stocks": 50,
                    "Bonds": 25,
                    "Gold": 15,
                    "Cash": 10
                }

            elif risk == "high":
                self.allocation = {
                    "Stocks": 70,
                    "Bonds": 15,
                    "Gold": 10,
                    "Cash": 5
                }

            else:
                messagebox.showerror("Error", "Enter valid risk level")
                return

            expected_return = self.calculate_expected_return()
            future_value = self.project_growth(amount, expected_return)

            self.result_label.config(
                text=f"Expected Annual Return: {round(expected_return*100,2)}%\n"
                     f"Projected 10-Year Value: {round(future_value,2)}"
            )

        except:
            messagebox.showerror("Error", "Invalid Input")

    def calculate_expected_return(self):

        returns = {
            "Stocks": 0.12,
            "Bonds": 0.06,
            "Gold": 0.08,
            "Cash": 0.04
        }

        expected_return = sum(
            (self.allocation[asset]/100) * returns[asset]
            for asset in self.allocation
        )

        return expected_return

    def project_growth(self, monthly_investment, annual_return):

        years = 10
        months = years * 12
        monthly_return = annual_return / 12

        future_value = 0
        for _ in range(months):
            future_value = (future_value + monthly_investment) * (1 + monthly_return)

        return future_value

    def show_chart(self):

        if not self.allocation:
            messagebox.showerror("Error", "Generate recommendation first")
            return

        labels = list(self.allocation.keys())
        sizes = list(self.allocation.values())

        plt.figure()
        plt.pie(sizes, labels=labels, autopct='%1.1f%%')
        plt.title("Portfolio Allocation")
        plt.show()

if __name__ == "__main__":
    root = tk.Tk()
    app = InvestmentApp(root)
    root.mainloop()
