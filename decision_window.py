
#Marcos Hung
from tkinter import *
#Hold response to whether the emails will be sent recurringly or not
class Decision_Window:
    def __init__(self, title, display_text, button_name_one, button_name_two):
        self.root = Tk()
        #option is a boolean that will be true if the first button is selected and false if the second button is selected
        self.option = None
        self.title = title
        self.display_text = display_text
        self.button_one = button_name_one
        self.button_two = button_name_two

    def set_to_option1(self):
        self.option = True
        self.root.destroy()

    def set_to_option2(self):
        self.option = False
        self.root.destroy()
    
    def make_window(self):
        self.root.title(self.title)
        top = Frame(self.root)
        bottom = Frame(self.root)
        top.pack(side = TOP)
        bottom.pack(side = BOTTOM, fill=BOTH, expand=True)

        first_button = Button(self.root, text= self.button_one, width=10, height=2, command= self.set_to_option1)
        second_button = Button(self.root, text =self.button_two, width = 10, height = 2, command = self.set_to_option2)

        second_button.pack(in_=bottom, side=BOTTOM)
        first_button.pack(in_=bottom, side=BOTTOM)


        t = Text(self.root, height = 1, width = 75)
        t.pack(in_ = top, side = TOP)
        t.insert(END, self.display_text)
        self.root.mainloop()