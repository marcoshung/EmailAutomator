#Marcos Hung
from tkinter import *

class Pop_Up:
    def __init__(self, title,display_text):
        self.root = Tk()
        self.title = title
        self.display_text = display_text
        
    def quit(self):
        self.root.destroy()

    def make_window(self):
        self.root.title(self.title)
        t = Text(self.root, height = 2, width = 75)
        t.pack()
        t.insert(END, self.display_text)
        quit_button = Button(self.root, text="Done", width=10, height=2, command=self.quit)
        quit_button.pack()
        self.root.mainloop()


