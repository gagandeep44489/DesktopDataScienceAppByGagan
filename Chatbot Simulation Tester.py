# Chatbot Simulation Tester - Desktop App in Python
# Features:
# - Simulate rule-based chatbot responses
# - User can type messages and see bot replies
# - Conversation history view
# - Reset conversation
# - Clean Tkinter-based desktop UI

import tkinter as tk
from tkinter import ttk, messagebox
import re

# ---------------- Chatbot Logic ----------------

def chatbot_response(user_input):
    user_input = user_input.lower().strip()

    patterns = {
        r"hi|hello|hey": "Hello! How can I assist you today?",
        r"how are you": "I am functioning properly. How can I help you?",
        r"what is your name": "I am a chatbot simulation tester.",
        r"help": "Sure. You can ask me general questions or test chatbot responses.",
        r"bye|exit|quit": "Goodbye! Have a great day."
    }

    for pattern, response in patterns.items():
        if re.search(pattern, user_input):
            return response

    return "I'm not sure how to respond to that. Please rephrase your question."

# ---------------- GUI Functions ----------------

def send_message():
    user_text = user_entry.get().strip()
    if not user_text:
        return

    chat_box.config(state=tk.NORMAL)
    chat_box.insert(tk.END, f"You: {user_text}\n")

    bot_reply = chatbot_response(user_text)
    chat_box.insert(tk.END, f"Bot: {bot_reply}\n\n")
    chat_box.config(state=tk.DISABLED)

    chat_box.see(tk.END)
    user_entry.delete(0, tk.END)


def reset_chat():
    chat_box.config(state=tk.NORMAL)
    chat_box.delete("1.0", tk.END)
    chat_box.config(state=tk.DISABLED)

# ---------------- Main Window ----------------

root = tk.Tk()
root.title("Chatbot Simulation Tester")
root.geometry("700x500")
root.resizable(False, False)

style = ttk.Style(root)
style.theme_use('clam')

# ---------------- Layout ----------------

main_frame = ttk.Frame(root, padding=15)
main_frame.pack(fill=tk.BOTH, expand=True)

header = ttk.Label(main_frame, text="Chatbot Simulation Tester", font=("Segoe UI", 14, "bold"))
header.pack(pady=5)

chat_box = tk.Text(main_frame, height=18, width=80, state=tk.DISABLED, wrap=tk.WORD)
chat_box.pack(pady=10)

entry_frame = ttk.Frame(main_frame)
entry_frame.pack(fill=tk.X)

user_entry = ttk.Entry(entry_frame, width=60)
user_entry.pack(side=tk.LEFT, padx=5)
user_entry.bind("<Return>", lambda event: send_message())

send_btn = ttk.Button(entry_frame, text="Send", command=send_message)
send_btn.pack(side=tk.LEFT, padx=5)

reset_btn = ttk.Button(main_frame, text="Reset Conversation", command=reset_chat)
reset_btn.pack(pady=10)

root.mainloop()
