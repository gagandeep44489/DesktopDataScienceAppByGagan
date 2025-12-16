import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder

class FeatureImportanceVisualizer:
    def __init__(self, root):
        self.root = root
        self.root.title("ML Feature Importance Visualizer")
        self.root.geometry("900x600")

        self.df = None

        # Top controls
        control_frame = tk.Frame(root)
        control_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

        load_btn = tk.Button(control_frame, text="Load CSV", command=self.load_csv)
        load_btn.pack(side=tk.LEFT, padx=5)

        tk.Label(control_frame, text="Target Column:").pack(side=tk.LEFT, padx=5)
        self.target_var = tk.StringVar()
        self.target_menu = tk.OptionMenu(control_frame, self.target_var, "")
        self.target_menu.pack(side=tk.LEFT, padx=5)

        visualize_btn = tk.Button(control_frame, text="Visualize Importance", command=self.visualize_importance)
        visualize_btn.pack(side=tk.LEFT, padx=10)

        # Plot frame
        self.plot_frame = tk.Frame(root)
        self.plot_frame.pack(fill=tk.BOTH, expand=True)

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return
        try:
            self.df = pd.read_csv(file_path)
            self.update_target_menu(self.df.columns)
            messagebox.showinfo("Success", "CSV loaded successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV:\n{e}")

    def update_target_menu(self, columns):
        self.target_menu['menu'].delete(0, 'end')
        for col in columns:
            self.target_menu['menu'].add_command(label=col, command=tk._setit(self.target_var, col))
        self.target_var.set(columns[0])

    def visualize_importance(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Load a CSV file first")
            return

        target_col = self.target_var.get()
        if target_col == "":
            messagebox.showwarning("Warning", "Select a target column")
            return

        try:
            X = self.df.drop(columns=[target_col])
            y = self.df[target_col]

            # Encode non-numeric features
            for col in X.select_dtypes(include=['object']).columns:
                X[col] = LabelEncoder().fit_transform(X[col])
            if y.dtype == 'object':
                y = LabelEncoder().fit_transform(y)

            model = RandomForestClassifier(n_estimators=100, random_state=42)
            model.fit(X, y)
            importance = model.feature_importances_

            feature_importance = pd.Series(importance, index=X.columns).sort_values(ascending=False)

            for widget in self.plot_frame.winfo_children():
                widget.destroy()

            fig, ax = plt.subplots(figsize=(8, 5))
            feature_importance.plot(kind='bar', ax=ax)
            ax.set_title("Feature Importance")
            ax.set_ylabel("Importance")
            ax.set_xlabel("Features")

            canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to visualize feature importance:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FeatureImportanceVisualizer(root)
    root.mainloop()