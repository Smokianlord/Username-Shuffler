import pandas as pd
import random
import tkinter as tk
from tkinter import messagebox
import pyperclip
import os
import sys

# ---------------- RESOURCE HELPERS ----------------

def resource_path(relative_path):
    """ Get resource path for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def get_app_folder():
    """ Get folder where app (exe or py) is located """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

# ---------------- GLOBALS ----------------

usernames = []
result_text = ""

# ---------------- DATA LOADING ----------------

def load_excel(show_message=True):
    global usernames

    folder_path = get_app_folder()
    valid_files = [".xlsx", ".xls", ".csv"]

    files = [
        f for f in os.listdir(folder_path)
        if any(f.lower().endswith(ext) for ext in valid_files)
    ]

    if not files:
        messagebox.showerror(
            "Error",
            "No Excel or CSV file found in the folder"
        )
        return

    file_name = files[0]
    file_path = os.path.join(folder_path, file_name)

    if file_name.lower().endswith(".csv"):
        df = pd.read_csv(file_path, header=None)
    else:
        df = pd.read_excel(file_path, header=None)

    usernames = (
        df.astype(str)
        .values
        .flatten()
        .tolist()
    )

    usernames = [
        u.strip() for u in usernames
        if u and u.lower() != "nan"
    ]

    if show_message:
        clean_name = os.path.splitext(file_name)[0]
        messagebox.showinfo(
            "Loaded",
            f"Loaded {len(usernames)} usernames from {clean_name}"
        )

def reload_excel():
    load_excel(show_message=True)
    focus_input()

# ---------------- LOGIC ----------------

def shuffle_and_pick(event=None):
    global result_text

    if not usernames:
        messagebox.showerror("Error", "Excel not loaded")
        focus_input()
        return

    try:
        set_size = int(set_size_entry.get())
        if set_size <= 0:
            raise ValueError
        if set_size > len(usernames):
            messagebox.showerror(
                "Error",
                "Set size cannot be larger than total usernames"
            )
            focus_input()
            return
    except:
        messagebox.showerror("Error", "Enter a valid number")
        focus_input()
        return

    selected = random.sample(usernames, set_size)
    result_text = " ".join(selected)

    output_box.delete("1.0", tk.END)
    output_box.insert(tk.END, result_text)

    focus_input()

def copy_result():
    if result_text:
        pyperclip.copy(result_text)
        messagebox.showinfo("Copied", "Result copied to clipboard")
    else:
        messagebox.showerror("Error", "Nothing to copy")

    focus_input()

# ---------------- FOCUS ----------------

def focus_input():
    set_size_entry.focus_force()
    set_size_entry.icursor(tk.END)

# ---------------- UI ----------------

root = tk.Tk()
root.title("Username Picker")

# Window icon
root.iconbitmap(resource_path("icon.ico"))

root.geometry("560x400")
root.resizable(False, False)
root.configure(bg="#F8FAFC")

FONT_TITLE = ("Segoe UI", 16, "bold")
FONT_NORMAL = ("Segoe UI", 11)
FONT_OUTPUT = ("Consolas", 11)

PRIMARY = "#2563EB"
ACCENT = "#16A34A"
TEXT = "#111827"

tk.Label(
    root,
    text="Username Picker",
    font=FONT_TITLE,
    bg="#F8FAFC",
    fg=PRIMARY
).pack(pady=(18, 12))

controls = tk.Frame(root, bg="#F8FAFC")
controls.pack(pady=8)

tk.Label(
    controls,
    text="Set size",
    font=FONT_NORMAL,
    bg="#F8FAFC",
    fg=TEXT
).grid(row=0, column=0, padx=6)

set_size_entry = tk.Entry(
    controls,
    width=10,
    font=FONT_NORMAL,
    justify="center",
    relief="solid",
    bd=1
)
set_size_entry.grid(row=0, column=1, padx=6)
set_size_entry.bind("<Return>", shuffle_and_pick)

tk.Button(
    controls,
    text="Shuffle",
    font=FONT_NORMAL,
    bg=PRIMARY,
    fg="white",
    width=10,
    bd=0,
    command=shuffle_and_pick
).grid(row=0, column=2, padx=6)

tk.Button(
    controls,
    text="Reload",
    font=FONT_NORMAL,
    bg=ACCENT,
    fg="white",
    width=10,
    bd=0,
    command=reload_excel
).grid(row=0, column=3, padx=6)

output_box = tk.Text(
    root,
    height=5,
    font=FONT_OUTPUT,
    wrap="word",
    bg="white",
    fg=TEXT,
    relief="solid",
    bd=1
)
output_box.pack(padx=20, pady=18, fill=tk.X)

tk.Button(
    root,
    text="Copy",
    font=FONT_NORMAL,
    bg="#0EA5E9",
    fg="white",
    width=18,
    bd=0,
    command=copy_result
).pack(pady=(0, 20))

# ---------------- STARTUP ----------------

root.after(200, load_excel)
root.after(300, focus_input)

root.mainloop()
