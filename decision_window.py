
#Marcos Hung
from tkinter import *
#Hold response to whether the emails will be sent recurringly or not
class Decision_Window:
    def __init__(self):
        self.root = Tk()
        self.recurring = None

    def set_to_recur(self):
        self.recurring = True
        self.root.destroy()

    def set_to_one_time(self):
        self.recurring = False
        self.root.destroy()
    
    def make_window(self):
        self.root.title("Recurring or One Time")
        top = Frame(self.root)
        bottom = Frame(self.root)
        top.pack(side = TOP)
        bottom.pack(side = BOTTOM, fill=BOTH, expand=True)

        recurring_button = Button(self.root, text="Recurring", width=10, height=2, command= self.set_to_recur)
        one_time_button = Button(self.root, text = "One Time", width = 10, height = 2, command = self.set_to_one_time)

        recurring_button.pack(in_=bottom, side=LEFT)
        one_time_button.pack(in_=bottom, side=LEFT)

        t = Text(self.root, height = 1, width = 40)
        t.pack(in_ = top, side = TOP)
        t.insert(END, "Recurring or One Time Emails?")
        self.root.mainloop()