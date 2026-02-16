import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from mlxtend.frequent_patterns import apriori, association_rules
from mlxtend.preprocessing import TransactionEncoder

class MarketBasketApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Market Basket Analysis Tool")
        self.root.geometry("900x600")

        self.data = None
        self.rules = None

        # UI Elements
        tk.Label(root, text="Minimum Support").pack()
        self.support_entry = tk.Entry(root)
        self.support_entry.insert(0, "0.02")
        self.support_entry.pack()

        tk.Label(root, text="Minimum Confidence").pack()
        self.conf_entry = tk.Entry(root)
        self.conf_entry.insert(0, "0.3")
        self.conf_entry.pack()

        tk.Button(root, text="Load CSV", command=self.load_file).pack(pady=5)
        tk.Button(root, text="Run Analysis", command=self.run_analysis).pack(pady=5)
        tk.Button(root, text="Save Rules", command=self.save_rules).pack(pady=5)

        # Table
        self.tree = ttk.Treeview(root)
        self.tree.pack(fill="both", expand=True)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.data = pd.read_csv(file_path)
            messagebox.showinfo("Success", "File Loaded Successfully")

    def run_analysis(self):
        try:
            min_support = float(self.support_entry.get())
            min_conf = float(self.conf_entry.get())

            if self.data is None:
                messagebox.showerror("Error", "Load a CSV file first")
                return

            # Convert to basket format
            transactions = self.data.groupby('TransactionID')['Item'].apply(list).tolist()

            te = TransactionEncoder()
            te_array = te.fit(transactions).transform(transactions)
            df = pd.DataFrame(te_array, columns=te.columns_)

            # Apriori
            frequent_itemsets = apriori(df, min_support=min_support, use_colnames=True)

            # Association Rules
            self.rules = association_rules(frequent_itemsets, metric="confidence", min_threshold=min_conf)

            if self.rules.empty:
                messagebox.showinfo("Result", "No rules found with given thresholds")
                return

            self.display_rules()

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def display_rules(self):
        self.tree.delete(*self.tree.get_children())

        columns = list(self.rules.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for _, row in self.rules.iterrows():
            self.tree.insert("", "end", values=list(row))

    def save_rules(self):
        if self.rules is not None:
            file_path = filedialog.asksaveasfilename(defaultextension=".csv")
            if file_path:
                self.rules.to_csv(file_path, index=False)
                messagebox.showinfo("Saved", "Rules saved successfully")

if __name__ == "__main__":
    root = tk.Tk()
    app = MarketBasketApp(root)
    root.mainloop()
