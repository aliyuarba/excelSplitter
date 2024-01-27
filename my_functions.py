import tkinter as tk
from tkinter import ttk

def start_progress():
    progress_var.set(0)  # Atur nilai awal progress bar
    for i in range(101):
        progress_var.set(i)  # Atur nilai progress bar pada setiap iterasi
        root.update_idletasks()  # Pembaruan tampilan agar progress bar terlihat
        progress_bar.update()  # Pembaruan tampilan progress bar
        root.after(50)  # Jeda 50 milidetik (simulasi pekerjaan)


root = tk.Tk()

progress_var = tk.DoubleVar()

progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.grid()


start_button = tk.Button(root, text="Mulai Pekerjaan", command=start_progress)
start_button.grid()

root.mainloop()