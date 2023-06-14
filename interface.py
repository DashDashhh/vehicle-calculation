import tkinter as tk
from main import main
import time

root = tk.Tk()

def click():
    main()
    label = tk.Label(root, text="script run!")
    label.pack()

button = tk.Button(root, text="click me", command=click)
button.pack()

root.mainloop()