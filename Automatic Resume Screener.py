import csv
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import PyPDF2
except ImportError:
    PyPDF2 = None


resumes_data = []


def normalize_text(text):
    return re.sub(r"[^a-z0-9 ]+", " ", text.lower())


def extract_keywords(text):
    words = normalize_text(text).split()
    return {word for word in words if len(word) > 2}


def read_resume(file_path):
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".txt":
        with open(file_path, "r", encoding="utf-8", errors="ignore") as file:
            return file.read()

    if ext == ".pdf":
        if PyPDF2 is None:
            raise RuntimeError(
                "PyPDF2 is not installed. Install it to parse PDF resumes."
            )

        text_parts = []
        with open(file_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text_parts.append(page.extract_text() or "")

        return "\n".join(text_parts)

    raise RuntimeError(f"Unsupported file type: {ext}. Use .txt or .pdf")


def score_resume(resume_text, required_keywords, preferred_keywords):
    resume_words = extract_keywords(resume_text)

    required_matches = sorted(resume_words.intersection(required_keywords))
    preferred_matches = sorted(resume_words.intersection(preferred_keywords))

    required_score = (
        len(required_matches) / len(required_keywords) * 100
        if required_keywords
        else 100
    )
    preferred_score = (
        len(preferred_matches) / len(preferred_keywords) * 100
        if preferred_keywords
        else 100
    )

    final_score = (0.75 * required_score) + (0.25 * preferred_score)

    return {
        "required_score": required_score,
        "preferred_score": preferred_score,
        "final_score": final_score,
        "required_matches": ", ".join(required_matches) if required_matches else "-",
        "preferred_matches": ", ".join(preferred_matches) if preferred_matches else "-",
    }


def choose_resumes():
    files = filedialog.askopenfilenames(
        title="Select Resume Files",
        filetypes=[("Resume Files", "*.txt *.pdf")],
    )

    if files:
        selected_files_var.set(f"{len(files)} file(s) selected")
        selected_files_var.files = list(files)
    else:
        selected_files_var.set("No files selected")
        selected_files_var.files = []


def run_screening():
    global resumes_data

    files = getattr(selected_files_var, "files", [])
    if not files:
        messagebox.showerror("Error", "Please select at least one resume file.")
        return

    job_text = job_description_text.get("1.0", tk.END).strip()
    required_text = required_skills_entry.get().strip()
    preferred_text = preferred_skills_entry.get().strip()

    if not job_text and not required_text:
        messagebox.showerror(
            "Error",
            "Provide a job description or required skills before screening.",
        )
        return

    required_keywords = extract_keywords(job_text)
    required_keywords.update(extract_keywords(required_text))
    preferred_keywords = extract_keywords(preferred_text)
    preferred_keywords = preferred_keywords.difference(required_keywords)

    results_tree.delete(*results_tree.get_children())
    resumes_data = []

    for file_path in files:
        try:
            text = read_resume(file_path)
            scores = score_resume(text, required_keywords, preferred_keywords)

            row = {
                "name": os.path.basename(file_path),
                "path": file_path,
                "required_score": round(scores["required_score"], 2),
                "preferred_score": round(scores["preferred_score"], 2),
                "final_score": round(scores["final_score"], 2),
                "required_matches": scores["required_matches"],
                "preferred_matches": scores["preferred_matches"],
            }
            resumes_data.append(row)
        except Exception as err:
            resumes_data.append(
                {
                    "name": os.path.basename(file_path),
                    "path": file_path,
                    "required_score": 0,
                    "preferred_score": 0,
                    "final_score": 0,
                    "required_matches": "Error",
                    "preferred_matches": str(err),
                }
            )

    resumes_data.sort(key=lambda item: item["final_score"], reverse=True)

    for resume in resumes_data:
        results_tree.insert(
            "",
            tk.END,
            values=(
                resume["name"],
                f"{resume['required_score']}%",
                f"{resume['preferred_score']}%",
                f"{resume['final_score']}%",
            ),
        )

    messagebox.showinfo("Completed", f"Screened {len(resumes_data)} resume(s).")


def export_results():
    if not resumes_data:
        messagebox.showerror("Error", "No results to export.")
        return

    file_path = filedialog.asksaveasfilename(
        title="Export Results",
        defaultextension=".csv",
        filetypes=[("CSV Files", "*.csv")],
    )

    if not file_path:
        return

    with open(file_path, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(
            file,
            fieldnames=[
                "name",
                "path",
                "required_score",
                "preferred_score",
                "final_score",
                "required_matches",
                "preferred_matches",
            ],
        )
        writer.writeheader()
        writer.writerows(resumes_data)

    messagebox.showinfo("Success", "Results exported successfully.")


root = tk.Tk()
root.title("Automatic Resume Screener")
root.geometry("950x700")

header = tk.Label(
    root,
    text="Automatic Resume Screener",
    font=("Arial", 18, "bold"),
)
header.pack(pady=10)

frame = tk.Frame(root)
frame.pack(fill="x", padx=15)

selected_files_var = tk.StringVar(value="No files selected")
selected_files_var.files = []

file_button = tk.Button(
    frame,
    text="Select Resumes (.txt/.pdf)",
    command=choose_resumes,
    bg="#0a66c2",
    fg="white",
)
file_button.grid(row=0, column=0, sticky="w")

file_label = tk.Label(frame, textvariable=selected_files_var)
file_label.grid(row=0, column=1, padx=10, sticky="w")

job_label = tk.Label(root, text="Job Description", font=("Arial", 11, "bold"))
job_label.pack(anchor="w", padx=15, pady=(12, 4))

job_description_text = tk.Text(root, height=8, wrap="word", bg="#f7f9fc")
job_description_text.pack(fill="x", padx=15)

skills_frame = tk.Frame(root)
skills_frame.pack(fill="x", padx=15, pady=10)

required_label = tk.Label(skills_frame, text="Required Skills (comma/space separated):")
required_label.grid(row=0, column=0, sticky="w")

required_skills_entry = tk.Entry(skills_frame)
required_skills_entry.grid(row=1, column=0, sticky="ew", padx=(0, 8))

preferred_label = tk.Label(skills_frame, text="Preferred Skills (comma/space separated):")
preferred_label.grid(row=0, column=1, sticky="w")

preferred_skills_entry = tk.Entry(skills_frame)
preferred_skills_entry.grid(row=1, column=1, sticky="ew")

skills_frame.columnconfigure(0, weight=1)
skills_frame.columnconfigure(1, weight=1)

controls = tk.Frame(root)
controls.pack(fill="x", padx=15, pady=5)

screen_button = tk.Button(
    controls,
    text="Run Screening",
    command=run_screening,
    bg="green",
    fg="white",
)
screen_button.pack(side="left")

export_button = tk.Button(
    controls,
    text="Export Results CSV",
    command=export_results,
    bg="orange",
    fg="black",
)
export_button.pack(side="left", padx=10)

results_label = tk.Label(root, text="Ranked Results", font=("Arial", 11, "bold"))
results_label.pack(anchor="w", padx=15, pady=(10, 4))

columns = ("name", "required", "preferred", "final")
results_tree = ttk.Treeview(root, columns=columns, show="headings", height=15)
results_tree.heading("name", text="Resume")
results_tree.heading("required", text="Required Match")
results_tree.heading("preferred", text="Preferred Match")
results_tree.heading("final", text="Overall Score")
results_tree.column("name", width=420)
results_tree.column("required", width=140, anchor="center")
results_tree.column("preferred", width=140, anchor="center")
results_tree.column("final", width=140, anchor="center")
results_tree.pack(fill="both", expand=True, padx=15, pady=(0, 15))

root.mainloop()