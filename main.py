#!/usr/bin/env python

#Marcos Hung
import smtplib, ssl
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
from email_window import Email_Window
from decision_window import Decision_Window
from pop_up import Pop_Up
from user_input_pop_up import User_Input
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import tkinter as tk
import html2text
from tkinter import *
from tkinter import simpledialog
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import datetime
import time
import os

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

#emails that will be sent on a recurring basis
def get_recurring_info():
    valid = False
    #while loop allows user to retry entering input if they inputted date or time incorrectly
    while(not valid):
        start_date_window = User_Input("Enter Starting Date", "Enter Starting Date you want to send the emails in MM-DD-YYYY form.")
        start_date = start_date_window.make_window()
        try:
            datetime.datetime.strptime(start_date, '%m-%d-%Y')
            valid = True
        except:
            Pop_Up("Date Error", "Incorrect Date Format. Try Again").make_window()

    valid = False
    while(not valid):
        end_date_window = User_Input("Enter Ending Date", "Enter Ending Date you want to send the emails in MM-DD-YYYY form.")
        end_date = start_date_window.make_window()
        try:
            datetime.datetime.strptime(end_date, '%m-%d-%Y')
            valid = True
        except:
            Pop_Up("Date Error", "Incorrect Date Format. Try Again").make_window()  
    valid = False
    while(not valid):
        time_window = User_Input("Enter Time", "Enter Time you want to send the emails in HH:MM format.")
        time = time_window.make_window()
        try:
            datetime.datetime.strptime(time, '%H:%M')
            valid = True
        except:
             Pop_Up("Time Error", "Incorrect Time Format. Try Again").make_window()

#info that will be sent at one specific time
def get_scheduled_info():
    valid = False
    #while loop allows user to retry entering input if they inputted date or time incorrectly
    while(not valid):
        date_window = User_Input("Enter Date", "Enter Date you want to send the emails in MM-DD-YYYY format.")
        date = date_window.make_window()
        try:
            datetime.datetime.strptime(date, '%m-%d-%Y')
            valid = True
        except:
            Pop_Up("Date Error", "Incorrect Date Format. Try Again").make_window()
    valid = False
    while(not valid):
        time_window = User_Input("Enter Time", "Enter Time you want to send the emails in HH:MM format.")
        time = time_window.make_window()
        try:
            datetime.datetime.strptime(time, '%H:%M')
            valid = True
        except:
             Pop_Up("Time Error", "Incorrect Time Format. Try Again").make_window()
    date = datetime.datetime.strptime(date, '%m-%d-%Y') 
    time = datetime.datetime.strptime(time, '%H:%M')
    dt= datetime.datetime(date.year, date.month, date.day, time.hour, time.minute)
    if(dt < datetime.datetime.now()):
        Pop_Up("Date and Time Error", "This date and time has already past").make_window()
        return
    return dt

def get_attachment():
    root = tk.Tk()
    root.title("Pick Your Attachment")
    t = Text(root, height = 2, width = 40)
    t.pack(in_ = root, side = TOP)
    t.insert(END, "Pick Your Attachment (NOTE* This attachment will be attached to ALL emails. Do not use this feature to send customized attachments to individuals")
    root.update()
    attachment_name = askopenfilename(title = "Get Attachment")
    root.destroy()

    # Open PDF file in binary mode
    with open(attachment_name, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)
    part.add_header("Content-Disposition","attachment", filename= os.path.basename(attachment_name))
    if(not part):
        Pop_Up("Attachment Error", "File Error")
    return part

    
def main():
    #reccuring or not
    recurring = Decision_Window("Recurring or one Time", "Is this event Recurring or One Time?", "Recurring", "One Time")
    recurring.make_window()

    #user exits pop up
    if(recurring.option == None):
        Pop_Up("No Selection", "No Selection was made. Exiting Program now").make_window()
        return

    if(recurring.option):
        get_recurring_info()
    else:
         scheduled = Decision_Window("Scheduled or Not Scheduled?", "Do you want to send these emails at a certain time or immediately?", 
                                     "Scheduled", "Immediately")
         scheduled.make_window()
         if(scheduled.option == None):
            Pop_Up("Schedule Error", "No Selection was made. Exiting Program")
            return
         elif(scheduled.option):
             scheduled_time = get_scheduled_info()
             if(scheduled_time == None):
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

    #email template
    email_template = get_email_template()
    if(email_template == None):
        Pop_Up("Template Error", "No Template Was Selected").make_window()
        return
    
    #attachments
    attachments = []
    attachment_window = Decision_Window("Attachments", "Would you like to add an Attachment?", "Yes", "No")
    attachment_window.make_window()
    if(attachment_window.option):
        done = False
        while(not done):
            attachments.append(get_attachment())
            again = Decision_Window("Add Another?", "Would You like to add another attachment?", "Yes", "No")
            again.make_window()
            if(not again.option):
                done = True

    #automator for sending emails
    send_all = False
    send_scheduled = False
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
        if(not len(attachments) == 0):
            for attachment in attachments:
                message.attach(attachment)
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
                try:
                    server.login(sender, password)
                except:
                    cred_error = Pop_Up("Credential Error", "Credentials Incorrcet. Please check to make sure your credential file is correct")
                    cred_error.make_window()
                    return
                #recurring
                if(recurring.option):
                    pass
                #one time
                else:
                    #scheduled
                   if(scheduled.option):
                       #check if the first schedule has been sent. this is bc if the loop goes to the sleep method it will sleep again causing error
                       if(not send_scheduled):
                            time.sleep(scheduled_time.timestamp() - time.time())
                            send_scheduled = True
                       server.sendmail(sender, receiver, message.as_string())
                    #immediately
                   else:
                       server.sendmail(sender, receiver, message.as_string())

                #won't make a pop up for each individual if send all option was chosen bc its annoying
                if(not send_all):
                    pop_up = Pop_Up("Email Sent", "Email sent @" + receiver + "!")
                    pop_up.make_window() 
                server.quit()
    pop_up = Pop_Up("Complete", "All Emails Sent!")
    pop_up.make_window()

if __name__ == "__main__":
    main()