import tkinter as tk
from tkinter import messagebox
import networkx as nx
import matplotlib.pyplot as plt

def generate_map():
    text = input_text.get("1.0", tk.END).strip()
    
    if not text:
        messagebox.showwarning("Warning", "Please enter concept relationships.")
        return

    G = nx.DiGraph()

    lines = text.split("\n")
    
    for line in lines:
        if "->" in line:
            parts = line.split("->")
            if len(parts) == 2:
                source = parts[0].strip()
                target = parts[1].strip()
                G.add_edge(source, target)

    if len(G.nodes) == 0:
        messagebox.showerror("Error", "Invalid format. Use: Concept1 -> Concept2")
        return

    plt.figure(figsize=(8,6))
    pos = nx.spring_layout(G)
    nx.draw(G, pos, with_labels=True, node_size=3000, node_color="lightblue",
            font_size=10, font_weight="bold", arrows=True)
    plt.title("Concept Map")
    plt.show()

def clear_text():
    input_text.delete("1.0", tk.END)

# Main Window
root = tk.Tk()
root.title("Concept Map Visualizer")
root.geometry("600x400")

tk.Label(root, text="Enter Relationships (Format: Concept1 -> Concept2)", 
         font=("Arial", 11, "bold")).pack(pady=10)

input_text = tk.Text(root, height=12, wrap="word")
input_text.pack(fill="both", padx=10, pady=5)

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(button_frame, text="Generate Concept Map", 
          command=generate_map, bg="green", fg="white").pack(side="left", padx=5)

tk.Button(button_frame, text="Clear", 
          command=clear_text, bg="red", fg="white").pack(side="left", padx=5)

root.mainloop()