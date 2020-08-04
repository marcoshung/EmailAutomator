#!/usr/bin/env python

#Marcos Hung
import smtplib, ssl
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
from email_window import Email_Window
from decision_window import Decision_Window
from pop_up import Pop_Up
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
import html2text
from tkinter import *
from tkinter import simpledialog
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from pathlib import Path

port = 465
context = ssl.create_default_context()

def get_subject_line():
    root = tk.Tk()
    root.withdraw()
    line= simpledialog.askstring(title="Subject Line", prompt="Enter the Subject Line You Want to Use for All the Emails") 
    root.destroy()
    return line


def get_email_info():
    root = tk.Tk()
    root.title("Pick The Excel File With the Email Information")
    t = Text(root, height = 2, width = 50)
    t.pack(in_ = root, side = TOP)
    t.insert(END, "Pick Your Excel File With the Email Information")
    root.update()
    #only accept excel files for email info
    email_info = askopenfilename(title = "Get Email Information", filetypes=[("Excel files", ".xlsx .xls")])
    root.destroy()
    return email_info

#message template that will be used. Can make edits if would like to change the text
def get_email_template():
    #keeps root tk window from opening
    root = tk.Tk()
    root.title("Pick The Email Template")
    t = Text(root, height = 2, width = 40)
    t.pack(in_ = root, side = TOP)
    t.insert(END, "Pick Your Email Template")
    root.update()

    #gets file name from user selection
    filename = askopenfilename(title = "Get Email Template", filetypes=[(".txt files", ".txt")])
    if(len(filename) == 0):
        return
    template_file  = open(filename, mode = 'r')

    #destroys root window to allow program to continue
    root.destroy()

    #parses the file contents
    html_template = template_file.read()
    template_file.close()

    return html_template

#reads email login creditionals in a specified credentials.txt file in same directory as this program

def read_creds():
    root = tk.Tk()
    root.title("Pick Your Creditional Text File")
    t = Text(root, height = 2, width = 40)
    t.pack(in_ = root, side = TOP)
    t.insert(END, "Pick Your Creditional Text File (.txt File)")
    root.update()
    filename = askopenfilename(title = "Get Creditional File", filetypes=[(".txt files", ".txt")])
    root.destroy()

    #returns empty usr and pass if no file was selected
    if len(filename) == 0:
        return ["",""]

    with open(filename, "r") as f:
        file = f.readlines()
        user = file[0].strip()
        passw = file[1].strip()

    return user, passw
#reads excel sheet based on predetermined information and guidelines
def read_email_info(filename):
    email_info = []
    df = pd.read_excel(filename)
    for i in range(len(df)):
        values = df.loc[i].to_numpy()
        email_info.append(values.tolist())
    return email_info


def main():
  
    decision = Decision_Window()
    decision.make_window()
    if(decision.recurring == None):
        Pop_Up("No Selection", "No Selection was made. Exiting Program now").make_window()
        return
    subject_line = get_subject_line()

    #user presses cancels
    if subject_line == None:
        exit_mes = Pop_Up("Program Exit", "Program will exit now")
        exit_mes.make_window()
        return 
    #creds
    sender, password = read_creds()
    if sender == "" or password == "":
        cred_error = Pop_Up("Creditional Error", "No creditionals found")
        cred_error.make_window()
        return
    #email content 
    excel_file_name = get_email_info()
    excel_email_info = read_email_info(excel_file_name)
    if excel_email_info == None:
        Pop_Up("Excel Error", "There was no data in the excel file").make_window()
        return

    email_template = get_email_template()
    
    if(email_template == None):
        Pop_Up("Template Error", "No Template Was Selected").make_window()
        return
    send_all = False

    print("Starting to send")

    #automator for sending emails
    for i in range(len(excel_email_info)):
        message = MIMEMultipart("alternative")
        message["Subject"] = subject_line
        message["From"] = sender
        #fills template with values
        email_info = excel_email_info[i]
        receiver = email_info[0]

        email_info.pop(0)
        message["To"] = receiver
        try:
            html_part = email_template % tuple(email_info)
        except:
            error_message = Pop_Up("Value Error", "The number of variables in your .txt file does not match the number of variables in your excel file")
            error_message.make_window()
            return

        #constructs email message body
        html_attachment = MIMEText(html_part, "html")
        message.attach(html_attachment)
        soup = BeautifulSoup(html2text.html2text(html_part), 'lxml')

        #this prevents the individual email preview from popping up
        if(not send_all):
            email_text_window = Email_Window(soup.text)
            email_text_window.make_window()

        if(email_text_window.skip_all):
            skipped_all_pop_up = Pop_Up("Complete", "All remaining emails will not send")
            skipped_all_pop_up.make_window()
            return

        if(email_text_window.send_all):
            send_all = True

        if(email_text_window.send or send_all):
            with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
                server.login(sender, password)
                server.sendmail(sender, receiver, message.as_string())

                #won't make a pop up for each individual bc its annoying
                if(not send_all):
                    pop_up = Pop_Up("Email Sent", "Email sent @" + receiver + "!")
                    pop_up.make_window()    
    pop_up = Pop_Up("Complete", "All Emails Sent!")
    pop_up.make_window()

if __name__ == "__main__":
    main()