#Marcos Hung
from tkinter import *

class Email_Window:
    def __init__(self, display_text):
        self.root = Tk()
        self.display_text = display_text
        self.send_all = False
        self.skip_all = False
        self.send = False
    
    def do_nothing(self):
        print("Nothing")

    def send_email(self):
        self.send = True
        self.root.destroy()
        

    def dont_send(self):
        self.send = False
        self.root.destroy()

    def send_all_email(self):
        self.send_all = True
        self.root.destroy()
    def skip_all_email(self):
        self.skip_all = True
        self.root.destroy()

    def make_window(self):
        self.root.title("Email Preview")
        top = Frame(self.root)
        bottom = Frame(self.root)
        top.pack(side = TOP)
        bottom.pack(side = BOTTOM, fill=BOTH, expand=True)

        send_button = Button(self.root, text="Send", width=10, height=2, command=self.send_email)
        skip_button = Button(self.root, text="Skip", width=10, height=2, command=self.dont_send)
        skip_all_button = Button(self.root, text="Skip All", width=10, height=2, command=self.skip_all_email)
        send_all_button = Button(self.root, text = "Send All", width = 10, height = 2, command = self.send_all_email)

        send_all_button.pack(in_=bottom, side=BOTTOM)
        skip_all_button.pack(in_=bottom, side = BOTTOM)
        skip_button.pack(in_=bottom, side=BOTTOM)
        send_button.pack(in_=bottom, side=BOTTOM)

        t = Text(self.root, height = 40, width = 200)
        t.pack(in_ = top, side = TOP)
        t.insert(END, self.display_text)
        self.root.mainloop()