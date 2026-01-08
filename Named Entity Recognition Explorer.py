import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import spacy

class NERExplorerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Named Entity Recognition Explorer")
        self.root.geometry("1000x650")

        try:
            self.nlp = spacy.load("en_core_web_sm")
        except OSError:
            messagebox.showerror(
                "Model Missing",
                "spaCy language model not found.\nRun:\npython -m spacy download en_core_web_sm"
            )
            root.destroy()
            return

        self.create_widgets()

    def create_widgets(self):
        # Input Text Area
        input_frame = tk.Frame(self.root)
        input_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(input_frame, text="Input Text", font=("Arial", 11, "bold")).pack(anchor="w")

        self.text_input = tk.Text(input_frame, height=8, wrap=tk.WORD)
        self.text_input.pack(fill=tk.X)

        # Buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=8)

        tk.Button(button_frame, text="Load Text File", width=18, command=self.load_text).grid(row=0, column=0, padx=5)
        tk.Button(button_frame, text="Run NER Analysis", width=20, command=self.run_ner).grid(row=0, column=1, padx=5)
        tk.Button(button_frame, text="Clear", width=15, command=self.clear_all).grid(row=0, column=2, padx=5)

        # Output Frames
        output_frame = tk.Frame(self.root)
        output_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Highlighted Text
        text_frame = tk.LabelFrame(output_frame, text="Entity Highlighted Text")
        text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        self.output_text = tk.Text(text_frame, wrap=tk.WORD)
        self.output_text.pack(fill=tk.BOTH, expand=True)

        # Entity Table
        table_frame = tk.LabelFrame(output_frame, text="Extracted Entities")
        table_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)

        columns = ("Entity", "Label", "Description")
        self.entity_table = ttk.Treeview(table_frame, columns=columns, show="headings")

        for col in columns:
            self.entity_table.heading(col, text=col)
            self.entity_table.column(col, width=150)

        self.entity_table.pack(fill=tk.BOTH, expand=True)

        self.label_map = spacy.explain

    def load_text(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if file_path:
            with open(file_path, "r", encoding="utf-8") as file:
                content = file.read()
                self.text_input.delete(1.0, tk.END)
                self.text_input.insert(tk.END, content)

    def run_ner(self):
        text = self.text_input.get(1.0, tk.END).strip()
        if not text:
            messagebox.showwarning("Input Required", "Please enter or load text.")
            return

        doc = self.nlp(text)

        self.output_text.delete(1.0, tk.END)
        self.entity_table.delete(*self.entity_table.get_children())

        start = 0
        for ent in doc.ents:
            self.output_text.insert(tk.END, text[start:ent.start_char])
            tag = ent.label_

            self.output_text.insert(tk.END, ent.text, tag)
            self.output_text.tag_config(tag, background="lightyellow")

            description = self.label_map(ent.label_) or "N/A"
            self.entity_table.insert("", tk.END, values=(ent.text, tag, description))

            start = ent.end_char

        self.output_text.insert(tk.END, text[start:])

    def clear_all(self):
        self.text_input.delete(1.0, tk.END)
        self.output_text.delete(1.0, tk.END)
        self.entity_table.delete(*self.entity_table.get_children())

if __name__ == "__main__":
    root = tk.Tk()
    app = NERExplorerApp(root)
    root.mainloop()
