import tkinter as tk
from tkinter import filedialog, messagebox
import matplotlib.pyplot as plt

def load_file():
    global sequence
    file_path = filedialog.askopenfilename(
        filetypes=[("Text Files", "*.txt"), ("FASTA Files", "*.fasta")]
    )
    if not file_path:
        return
    
    try:
        with open(file_path, "r") as file:
            lines = file.readlines()
            
            # Remove FASTA headers
            sequence = ""
            for line in lines:
                if not line.startswith(">"):
                    sequence += line.strip().upper()
        
        label_status.config(text="File Loaded Successfully")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def analyze_sequence():
    if not sequence:
        messagebox.showwarning("Warning", "Please load a DNA file first.")
        return

    a = sequence.count("A")
    t = sequence.count("T")
    g = sequence.count("G")
    c = sequence.count("C")

    total = len(sequence)
    gc_content = ((g + c) / total) * 100 if total > 0 else 0

    result_text = f"""
Sequence Length: {total}

Adenine (A): {a}
Thymine (T): {t}
Guanine (G): {g}
Cytosine (C): {c}

GC Content: {gc_content:.2f}%
    """

    text_result.delete(1.0, tk.END)
    text_result.insert(tk.END, result_text)

    # Plot Graph
    plt.figure()
    plt.bar(["A", "T", "G", "C"], [a, t, g, c])
    plt.title("Nucleotide Distribution")
    plt.xlabel("Nucleotides")
    plt.ylabel("Count")
    plt.show()


# GUI Setup
root = tk.Tk()
root.title("Genetic Data Visualizer")
root.geometry("600x550")

sequence = ""

tk.Label(root, text="Genetic Data Visualizer",
         font=("Arial", 16, "bold")).pack(pady=10)

tk.Button(root, text="Load DNA File",
          command=load_file,
          bg="blue", fg="white").pack(pady=10)

tk.Button(root, text="Analyze & Visualize",
          command=analyze_sequence,
          bg="green", fg="white").pack(pady=10)

label_status = tk.Label(root, text="")
label_status.pack()

text_result = tk.Text(root, height=15, width=70)
text_result.pack(pady=10)

root.mainloop()
