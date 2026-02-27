import tkinter as tk
from tkinter import messagebox
import matplotlib.pyplot as plt

class GameScoreAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Game Score Analyzer")
        self.root.geometry("500x500")
        
        self.players = []
        self.scores = []
        
        # Title Label
        tk.Label(root, text="Game Score Analyzer", font=("Arial", 16, "bold")).pack(pady=10)
        
        # Player Name Input
        tk.Label(root, text="Player Name:").pack()
        self.name_entry = tk.Entry(root)
        self.name_entry.pack()
        
        # Score Input
        tk.Label(root, text="Score:").pack()
        self.score_entry = tk.Entry(root)
        self.score_entry.pack()
        
        # Buttons
        tk.Button(root, text="Add Score", command=self.add_score).pack(pady=5)
        tk.Button(root, text="Show Statistics", command=self.show_stats).pack(pady=5)
        tk.Button(root, text="Show Chart", command=self.show_chart).pack(pady=5)
        tk.Button(root, text="Clear Data", command=self.clear_data).pack(pady=5)
        
        # Output Area
        self.output_text = tk.Text(root, height=10, width=50)
        self.output_text.pack(pady=10)
        
    def add_score(self):
        name = self.name_entry.get()
        score = self.score_entry.get()
        
        if name == "" or score == "":
            messagebox.showerror("Error", "Please enter both name and score.")
            return
        
        try:
            score = float(score)
        except ValueError:
            messagebox.showerror("Error", "Score must be a number.")
            return
        
        self.players.append(name)
        self.scores.append(score)
        
        self.output_text.insert(tk.END, f"{name} - {score}\n")
        
        self.name_entry.delete(0, tk.END)
        self.score_entry.delete(0, tk.END)
    
    def show_stats(self):
        if not self.scores:
            messagebox.showwarning("Warning", "No data available.")
            return
        
        total = sum(self.scores)
        average = total / len(self.scores)
        highest = max(self.scores)
        lowest = min(self.scores)
        
        stats = (
            f"\nTotal Score: {total}\n"
            f"Average Score: {average:.2f}\n"
            f"Highest Score: {highest}\n"
            f"Lowest Score: {lowest}\n"
        )
        
        self.output_text.insert(tk.END, stats)
    
    def show_chart(self):
        if not self.scores:
            messagebox.showwarning("Warning", "No data to display.")
            return
        
        plt.figure()
        plt.bar(self.players, self.scores)
        plt.xlabel("Players")
        plt.ylabel("Scores")
        plt.title("Game Score Analysis")
        plt.show()
    
    def clear_data(self):
        self.players.clear()
        self.scores.clear()
        self.output_text.delete(1.0, tk.END)

# Run Application
if __name__ == "__main__":
    root = tk.Tk()
    app = GameScoreAnalyzer(root)
    root.mainloop()