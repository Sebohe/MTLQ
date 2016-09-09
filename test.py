import tkinter as tk
from tkinter import ttk
import threading
import time
import os
import psutil

class App(tk.Tk):

	def __init__(self):
		tk.Tk.__init__(self)
		self.progressbar = ttk.Progressbar(self, orient='horizontal',
										   length=400, mode='determinate')
		self.progressbar.pack(padx=10, pady=10)

		initialMem = (psutil.virtual_memory()[3])
		GIGA = 7069642752 - 1000000000
		print (initialMem)
		print (GIGA)
		self.progressbar["value"] = GIGA




if __name__ == "__main__":
	app = App()
	app.mainloop()
