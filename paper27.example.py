import tkinter as tk

class CustomTitleBar(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.master = master
        self.configure(bg="blue")  # Customize the background color
        self.pack(fill=tk.X)

        # Create minimize button
        self.minimize_button = tk.Button(self, text="-", command=self.master.iconify)
        self.minimize_button.pack(side=tk.LEFT)

        # Create maximize button
        self.maximize_button = tk.Button(self, text="□", command=self.toggle_maximize)
        self.maximize_button.pack(side=tk.LEFT)

        # Create close button
        self.close_button = tk.Button(self, text="X", command=self.master.destroy)
        self.close_button.pack(side=tk.RIGHT)

    def toggle_maximize(self):
        if self.master.wm_attributes("-zoomed"):
            self.master.wm_attributes("-zoomed", False)
        else:
            self.master.wm_attributes("-zoomed", True)

# Example usage
root = tk.Tk()
root.title("Custom Title Bar Example")
root.geometry("400x300")

title_bar = CustomTitleBar(root)
content_frame = tk.Frame(root)
content_frame.pack(fill=tk.BOTH, expand=True)
label = tk.Label(content_frame, text="Your content here")
label.pack(padx=20, pady=20)

root.mainloop()
