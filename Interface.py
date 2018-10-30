import tkinter as tk
from tkinter import filedialog

# Main interface class for program.
# Handles user input and displaying program output to user.

class interface:
    def __init__(self, master):
        self.master = master
        self.window = tk.Frame(self.master)
        #self.window.tk.geometry(self.window, '750x500')

        lbl = tk.Label(self.window, text='Email Path:')
        lbl.grid(column=0, row=0)

        path = tk.Entry(self.window, width=50)
        path.grid(column=1, row=0)

        # TODO: Open file dialog for workspace
        def clicked(): 
            self.set_entry_text(path, self.browse_file_dialog())
        btn_browse = tk.Button(self.window, text='Browse', command=clicked)
        btn_browse.grid(column=2, row=0)

        # TODO: Add Start button

        self.window.pack()

    def set_entry_text(self, entry, text):
        entry.delete(0, tk.END)
        entry.insert(0, text)
        return

    # Open file dialog
    def browse_file_dialog(self):
        r =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("msg files","*.msg"),("Excel files","*.xlsx*")))
        return r

    def clicked(self, func):
        func()
        return

# print string with line start symbol prefix
# print_status("Hello World")
#   :">> Hello World"
def print_status(string):
    LINE_START = '>> '
    print(LINE_START + string)