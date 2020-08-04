#Marcos Hung
import tkinter as tk
from tkinter import Tk
from tkinter import messagebox

class Pop_Up:
    def __init__(self, title,display_text):
        self.root = Tk()
        self.title = title
        self.display_text = display_text
        
    def quit(self):
        self.root.destroy()

    def make_window(self):
        self.root.withdraw()
        messagebox.showinfo(title=self.title, message= self.display_text) 
        self.root.destroy()


