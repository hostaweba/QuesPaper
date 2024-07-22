# Importing the TitleBar class from the titlebar module
from titlebar import TitleBar
import tkinter as tk

# Creating a tkinter window
root = tk.Tk()
root.geometry("400x300")

# Creating an instance of the TitleBar class
title_bar = TitleBar(root, title="My App")

# Adding some buttons to the title bar
title_bar.add_button("Maximize", lambda: root.attributes("-zoomed", True))
title_bar.add_button("Minimize", lambda: root.attributes("-zoomed", False))
title_bar.add_button("Close", root.destroy)

# Running the tkinter event loop
root.mainloop()
