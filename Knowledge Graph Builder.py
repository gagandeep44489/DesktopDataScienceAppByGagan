import tkinter as tk
from tkinter import messagebox, filedialog
import networkx as nx
import matplotlib.pyplot as plt

def build_graph():
    text = input_text.get("1.0", tk.END).strip()
    
    if not text:
        messagebox.showwarning("Warning", "Please enter knowledge triples.")
        return

    G = nx.DiGraph()
    edges = []

    lines = text.split("\n")

    for line in lines:
        parts = line.split("|")
        if len(parts) == 3:
            subject = parts[0].strip()
            relation = parts[1].strip()
            obj = parts[2].strip()

            G.add_edge(subject, obj, label=relation)
            edges.append((subject, obj, relation))
        else:
            messagebox.showerror("Format Error", 
                                 "Each line must follow: Subject | Relationship | Object")
            return

    if len(G.nodes) == 0:
        messagebox.showerror("Error", "No valid triples found.")
        return

    plt.figure(figsize=(10, 7))
    pos = nx.spring_layout(G, k=0.5)

    nx.draw(G, pos, with_labels=True, node_size=3000, node_color="lightgreen",
            font_size=9, font_weight="bold", arrows=True)

    edge_labels = {(s, o): r for s, o, r in edges}
    nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, font_size=8)

    plt.title("Knowledge Graph")
    plt.tight_layout()
    plt.show()


def clear_text():
    input_text.delete("1.0", tk.END)


# GUI Setup
root = tk.Tk()
root.title("Knowledge Graph Builder")
root.geometry("700x450")

tk.Label(root, text="Enter Knowledge Triples (Subject | Relationship | Object)", 
         font=("Arial", 11, "bold")).pack(pady=10)

input_text = tk.Text(root, height=15, wrap="word")
input_text.pack(fill="both", padx=10, pady=5)

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(button_frame, text="Build Knowledge Graph", 
          command=build_graph, bg="green", fg="white").pack(side="left", padx=5)

tk.Button(button_frame, text="Clear", 
          command=clear_text, bg="red", fg="white").pack(side="left", padx=5)

root.mainloop()