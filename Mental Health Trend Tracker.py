import tkinter as tk
from tkinter import messagebox
import sqlite3
from datetime import datetime
import matplotlib.pyplot as plt
import pandas as pd

# ===============================
# DATABASE SETUP
# ===============================

conn = sqlite3.connect("mental_health.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT,
    mood INTEGER,
    stress INTEGER,
    sleep REAL,
    energy INTEGER,
    notes TEXT
)
""")
conn.commit()

# ===============================
# FUNCTIONS
# ===============================

def save_record():
    mood = mood_var.get()
    stress = stress_var.get()
    sleep = sleep_var.get()
    energy = energy_var.get()
    notes = notes_entry.get("1.0", tk.END)

    if not mood or not stress:
        messagebox.showerror("Error", "Mood and Stress are required")
        return

    date = datetime.now().strftime("%Y-%m-%d")

    cursor.execute("""
        INSERT INTO records (date, mood, stress, sleep, energy, notes)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (date, mood, stress, sleep, energy, notes))

    conn.commit()
    messagebox.showinfo("Saved", "Record Saved Successfully")

def show_trends():
    df = pd.read_sql_query("SELECT * FROM records", conn)

    if df.empty:
        messagebox.showwarning("No Data", "No records found")
        return

    df['date'] = pd.to_datetime(df['date'])
    df = df.sort_values('date')

    plt.figure()
    plt.plot(df['date'], df['mood'])
    plt.title("Mood Trend Over Time")
    plt.xlabel("Date")
    plt.ylabel("Mood Level")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

def check_risk():
    df = pd.read_sql_query("SELECT * FROM records ORDER BY date DESC LIMIT 5", conn)
    
    if len(df) < 5:
        return

    if all(df['mood'] <= 3):
        messagebox.showwarning("Alert",
                               "Low mood detected for 5 consecutive days.\nConsider professional support.")

# ===============================
# GUI
# ===============================

root = tk.Tk()
root.title("Mental Health Trend Tracker")
root.geometry("400x500")

tk.Label(root, text="Mood (1-10)").pack()
mood_var = tk.IntVar()
tk.Entry(root, textvariable=mood_var).pack()

tk.Label(root, text="Stress (1-10)").pack()
stress_var = tk.IntVar()
tk.Entry(root, textvariable=stress_var).pack()

tk.Label(root, text="Sleep Hours").pack()
sleep_var = tk.DoubleVar()
tk.Entry(root, textvariable=sleep_var).pack()

tk.Label(root, text="Energy (1-10)").pack()
energy_var = tk.IntVar()
tk.Entry(root, textvariable=energy_var).pack()

tk.Label(root, text="Notes").pack()
notes_entry = tk.Text(root, height=4)
notes_entry.pack()

tk.Button(root, text="Save Record", command=save_record).pack(pady=10)
tk.Button(root, text="Show Mood Trends", command=show_trends).pack(pady=10)
tk.Button(root, text="Check Risk", command=check_risk).pack(pady=10)

root.mainloop()