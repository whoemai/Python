import tkinter as tk
import random
import string
import pyperclip

def generate_password():
    numberPassword = int(entry.get())
    password = ''.join(random.choices(string.ascii_letters + string.digits + string.punctuation, k=numberPassword))
    label.config(text=password)

def copy_password():
    pyperclip.copy(label["text"])

root = tk.Tk()
root.title("Password Generator")
root.geometry("400x150")

label = tk.Label(root, text="")
label.pack(side="top", fill="both", expand=True)

entry = tk.Entry(root, width=50)
entry.pack(side="top", fill="both", expand=True)
entry.insert(0, 20)

generate_button = tk.Button(root, text="Generate", command=generate_password)
generate_button.pack(side="left", padx=10)

copy_button = tk.Button(root, text="Copy", command=copy_password)
copy_button.pack(side="right", padx=10)

root.mainloop()