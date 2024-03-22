from tkinter import messagebox

class CustomException(Exception):
    def __init__(self, title, message):
        self.title = title
        self.message = message
        super().__init__(self.title,self.message)
        messagebox.showerror(self.title,self.message)
