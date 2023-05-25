from tkinter import messagebox

class CustomWarning(Exception):
	def __init__(self,title,message):
		self.title = title
		self.message = message
		super().__init__(self.title,self.message)
		messagebox.showwarning(self.title,self.message)