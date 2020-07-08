# -*- coding: UTF-8 -*-

import tkinter as tk
from tkinter import *
from tkinter import filedialog, StringVar
import json, os
import email, smtplib, ssl
import csv
import time
from email.mime.text import MIMEText 
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
from datetime import date
from email.mime.image import MIMEImage
from msvcrt import getch
import sys


path_of_a = []
open_p_o_a = []
header = []


def add_path_csv(event):
    filepath = filedialog.askopenfilename(initialdir="/", title="Select File", filetypes=(("csv", "*.csv"), ("all files", "*.*")))
    event.delete(0, END)
    event.insert(END, filepath)


def add_path_pdf(event):
    filepath = filedialog.askopenfilename(initialdir="/", title="Select File", filetypes=(("pdf", "*.pdf"), ("all files", "*.*")))
    event.insert(END, filepath)


def add_path_logo(event):
    filepath = filedialog.askopenfilename(initialdir="/", title="Select File", filetypes=(("png", "*.png"),("jpg", "*.jpg"), ("all files", "*.*")))
    event.delete(0, END)
    event.insert(END, filepath)


def save_to_file(_smtp, _smtp_port, _email_entry, _no_of_a, _csv_file):
    settings_dict = {}
    settings_dict['smtp_settings'] = {}
    settings_dict['file_settings'] = {}
    settings_dict['message_settings'] = {}
    settings_dict['img_settings'] = {}

    if _smtp is not None:
        x['smtp_settings']['smtp'] = _smtp
        settings_dict['smtp_settings']['smtp'] = _smtp
    else:
        settings_dict['smtp_settings']['smtp'] = x['smtp_settings']['smtp']

    if _smtp_port is not None:
        x['smtp_settings']['smtp_port'] = _smtp_port
        settings_dict['smtp_settings']['smtp_port'] = _smtp_port
    else:
        settings_dict['smtp_settings']['smtp_port'] = x['smtp_settings']['smtp_port']

    if _email_entry is not None:
        x['smtp_settings']['sender_email'] = _email_entry
        settings_dict['smtp_settings']['sender_email'] = _email_entry
    else:
        settings_dict['smtp_settings']['sender_email'] = x['smtp_settings']['sender_email']

    if len(path_of_a) > 0:
        settings_dict['file_settings']['file_name'] = []
        settings_dict['file_settings']['file_path'] = []
        for a in path_of_a:
            settings_dict['file_settings']['file_name'].append(os.path.basename(a.get()))
            settings_dict['file_settings']['file_path'].append(a.get())

    if _no_of_a is not None:
        if int(_no_of_a) > 0:
            
            x['file_settings']['file_numbers'] = int(_no_of_a)
            settings_dict['file_settings']['file_numbers'] = int(_no_of_a)
            create_file_path(x['file_settings']['file_numbers'])
            for i in range(len(path_of_a)):
                path_of_a[i].insert(END, settings_dict['file_settings']['file_path'][i])
        else:
            settings_dict['file_settings']['file_numbers'] = 0
            x['file_settings']['file_numbers'] = 0
    else:
        settings_dict['file_settings']['file_numbers'] = x['file_settings']['file_numbers']
    
    if _csv_file is not None:
        x['file_settings']['csv_path'] = _csv_file
        settings_dict['file_settings']['csv_path'] = _csv_file
        create_header(x)
        
    else:
        settings_dict['file_settings']['csv_path'] = x['file_settings']['csv_path']
    
    settings_dict['message_settings']['subject'] = email_subject.get()
    settings_dict['message_settings']['from'] = email_from.get()

    with open("html.txt", "w") as y:
        y.write(email_html.get("1.0", "end-1c"))

    settings_dict['message_settings']['html'] = "html.txt"

    settings_dict['img_settings']['img_file_path'] = email_logo.get()
    settings_dict['img_settings']['img_file_name'] = os.path.basename(email_logo.get())
    settings_dict['img_settings']['img_cid'] = f"{os.path.basename(email_logo.get().replace(' ', ''))}"
    
    with open('settings.txt', 'w', encoding='utf-8') as json_file:
        json.dump(settings_dict, json_file, ensure_ascii=False, indent=4)

    cid.delete(0, END)
    cid.insert(END, settings_dict['img_settings']['img_cid'])

    root.update()


def create_str(s, row):      #replace $monat$ $tag$ $jahr$ $name$ $surname$ in strings
    t = s.replace("$year$", year.get()).replace("$month$", month.get())

    for count, head in enumerate(header):
        t = t.replace(f"${head}$", row[count]) #TO DO
    return t


def send_mail():
    _smtp = x['smtp_settings']['smtp']
    _smtp_port = x['smtp_settings']['smtp_port']
    _email = x['smtp_settings']['sender_email']

    try:
        context = ssl.create_default_context()
        _password = password.get()
        server = smtplib.SMTP(_smtp, _smtp_port)
        server.starttls(context=context)                                            # secure tls connection
        server.login(_email, _password)
        log.insert(END, "\nlogin successful\n\n")
        root.update()
        print("\nlogin successful\n\n")
    except Exception as e:
        log.insert(END, f"\n{e}\n")
        root.update()
        print(e)        # print error code

    row_count = sum(1 for row in csv.reader(open(x['file_settings']['csv_path']))) - 1
    mail_count = 0 
    error_count = 0
    try:
        with open(x['file_settings']['csv_path']) as csvfile:       # open csv file
            reader = csv.reader(csvfile)
            while True:
                next(reader)
                #for name, surname, number, email in reader:         # iterate through rows
                for row in reader:
                    mail_count += 1
                    try:
                        message = MIMEMultipart("mixed")
                        message["Subject"] = create_str(email_subject.get(), row)
                        message["From"] = email_from.get()
                        message["To"] = row[len(row) - 1]

                        html = email_html.get("1.0", END) # html body from txt file
                        
                        part_html = MIMEText(create_str(html, row), "html") # create MIMEText html part
                        
                        if x['file_settings']['file_path'] != "":
                            message.attach(part_html) # attach html
                        
                        if x['img_settings']['img_file_path'] != "":
                            message = attach_img(message) # attach image from index in settings.txt list
                        
                        for i in range(int(x['file_settings']['file_numbers'])):       # attach pdf files from index in settings.txt list
                           message = attach_file(i, message, row)

                        #print(f"Sende Email {mail_count}/{row_count} an {name} {surname}")
                        log.insert(END, f"Sending email {mail_count}/{row_count} to {row[len(row) - 1]}\n")
                        root.update()
                        print(f"Sending email {mail_count}/{row_count} to {row[len(row) - 1]}\n")
                        
                        server.sendmail(_email, row[len(row) - 1], message.as_string())   # send composed emails

                    except Exception as e:
                        log.insert(END, f"Error in mail to {row[len(row) - 1]} with: {e}\n")
                        root.update()
                        print(f"Error in mail to {row[len(row) - 1]} with: {e}\n") # print out errors and continue with next row in .csv
                        error_count += 1
                        continue
                break
        if mail_count == row_count:
            log.insert(END, f"\nSending emails finished with {error_count} errors.")
            root.update()
            print(f"\nSending emails finished with {error_count} errors.") # confirm sending of all emails with number of errors.
        else: 
            log.insert(END, "\n\nUNEXPECTED ERROR")
            root.update()
            print("\n\nUNEXPECTED ERROR")
        server.quit()
    except Exception as e:
        log.insert(END, f"\n{e}\n")
        root.update()


def create_file_path(n):
    if len(path_of_a) > 0:
        for a in path_of_a:
            a.grid_forget()
        for b in open_p_o_a:
            b.grid_forget()

    for label in root.grid_slaves():
        if int(label.grid_info()["row"] == 0 and label.grid_info()["column"] == 3):
            label.grid_forget()

    del path_of_a[:]
    del open_p_o_a[:]

    if n > 0:
        fp_label = tk.Label(root, text="file paths")
        fp_label.grid(row=0, column=3, sticky=E)
        k = 0
        for i in range(n):
            k = int(i/4)
            l = i % 4
            path_of_a.append("")
            open_p_o_a.append("")
            path_of_a[i] = Entry(root)
            path_of_a[i].grid(row=0+l, column=4+k)
            open_p_o_a[i] = Button(root, text="open", command=lambda i=i: add_path_pdf(path_of_a[i]))
            open_p_o_a[i].grid(row=0+l, column=4+k, padx=(165, 0))


def attach_file(i, _message, row):     #attach file function
    
    file_name_ = create_str(os.path.basename(path_of_a[i].get()), row)
    
    file_path_ = create_str(path_of_a[i].get(), row)
    
    with open(file_path_, "rb") as attachment:
        file = MIMEBase("application", "pdf")
        file.set_payload(attachment.read())
        attachment.close()

    encoders.encode_base64(file)

    file.add_header("Content-Disposition", "attachment", filename= f"{file_name_}")
    _message.attach(file)
    return _message


def attach_img(_message):      #attach image function
    with open(email_logo.get(), "rb") as fp:
        msgImage = MIMEImage(fp.read())
        fp.close()
    msgImage.add_header("Content-ID", f"<{cid.get()}>")
    msgImage.add_header("Content-Disposition", "inline", filename=os.path.basename(email_logo.get()))
    _message.attach(msgImage)
    return _message


def settings_window(_file_loaded):
    settings = tk.Toplevel(root)
    settings.lift()
    settings.title("settings")

    # settings
    Label(settings, text="email").grid(row=1, column=0, sticky=E)
    email_entry = Entry(settings)
    email_entry.grid(row=1, column=1)

    # settings
    Label(settings, text="smtp").grid(row=2, column=0, sticky=E)
    smtp = Entry(settings)
    smtp.grid(row=2, column=1)

    # settings
    Label(settings, text="smtp port").grid(row=3, column=0, sticky=E)
    smtp_port = Entry(settings)
    smtp_port.grid(row=3, column=1)

    # settings
    Label(settings, text="csv file").grid(row=1, column=3, sticky=E)
    csv_file = Entry(settings)
    csv_file.grid(row=1, column=4)
    csv_file_open = Button(settings, text="open", command=lambda: add_path_csv(csv_file))
    csv_file_open.grid(row=1, column=5, sticky=W)

    # settings
    Label(settings, text="no of attach").grid(row=2, column=3, sticky=E)
    no_of_a = Entry(settings)
    no_of_a.grid(row=2, column=4)
    add = Button(settings, text="set", command=lambda: create_file_path(int(no_of_a.get())))
    add.grid(row=2, column=5)

    save = Button(settings, text="save", command=lambda: save_to_file(smtp.get(), smtp_port.get(), email_entry.get(), no_of_a.get(), csv_file.get()))
    save.grid(row=15, column=2, pady=(15, 0))

    if _file_loaded:
        email_entry.insert(END, x['smtp_settings']['sender_email']) # settings
        smtp.insert(END, x['smtp_settings']['smtp'])                # settings
        smtp_port.insert(END, x['smtp_settings']['smtp_port'])      # settings
        no_of_a.insert(END, x['file_settings']['file_numbers'])     # settings
        csv_file.insert(END, x['file_settings']['csv_path'])        # settings


def load_settings(loaded):
    if loaded:
        f = open("settings.txt", "r")     #open and load settings in json format
        _x = json.loads(f.read())
        _html = open(_x["message_settings"]["html"], "r").read()
        _file_numbers = _x["file_settings"]["file_numbers"]
        _file_loaded = True
        if _x['file_settings']['csv_path'] != "":
            create_header(_x)
        print("file loaded")
        f.close()
        return _x, _html, _file_numbers, _file_loaded
    else:
        _file_loaded = True

        settings_dict = {}
        settings_dict['smtp_settings'] = {}
        settings_dict['file_settings'] = {}
        settings_dict['message_settings'] = {}
        settings_dict['img_settings'] = {}

        settings_dict['smtp_settings']['smtp'] = ""
        settings_dict['smtp_settings']['smtp_port'] = ""
        settings_dict['smtp_settings']['sender_email'] = ""

        settings_dict['file_settings']['file_name'] = []
        settings_dict['file_settings']['file_path'] = []
        settings_dict['file_settings']['file_numbers'] = 0

        settings_dict['file_settings']['csv_path'] = ""

        settings_dict['message_settings']['subject'] = ""
        settings_dict['message_settings']['from'] = ""
        settings_dict['message_settings']['html'] = "html.txt"

        settings_dict['img_settings']['img_file_path'] = ""
        settings_dict['img_settings']['img_file_name'] = ""
        settings_dict['img_settings']['img_cid'] = ""
        
        with open('settings.txt', 'w', encoding='utf-8') as json_file:
            json.dump(settings_dict, json_file, ensure_ascii=False, indent=4)

        f = open("settings.txt", "r")     #open and load settings in json format
        _x = json.loads(f.read())
        f.close()

        if os.path.isfile(_x["message_settings"]["html"]):
            _html = open(_x["message_settings"]["html"], "r").read()
        else:
            _html = open("html.txt", "w")
            _html.close()
            _html = open("html.txt", "r").read()
                

        return _x, _html, settings_dict['file_settings']['file_numbers'], _file_loaded


def create_header(_x):
    del header[:]

    with open(_x['file_settings']['csv_path']) as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            for i in range(len(row) - 1):
                header.append(row[i])
            break
    Label(root, text=f"replaces $month$ with given month").grid(row=6, column=4, sticky=E)
    Label(root, text=f"replaces $year$ with given year").grid(row=7, column=4, sticky=E)

    for label in root.grid_slaves():
        if int(label.grid_info()["row"] >= 8 and label.grid_info()["column"] == 4):
            label.grid_forget()

    for i in range(len(header)):
        Label(root, text=f"replaces ${header[i]}$ with {header[i]}").grid(row=8+i, column=4, sticky=E)


root = tk.Tk()
root.title("Email script")

# year
Label(root, text="year").grid(row=0, column=0, sticky=E)
year = Entry(root)
year.grid(row=0, column=1)

# month
Label(root, text="month").grid(row=1, column=0, sticky=E)
month = Entry(root)
month.grid(row=1, column=1)

# email from
Label(root, text="from").grid(row=2, column=0, sticky=E)
email_from = Entry(root)
email_from.grid(row=2, column=1)

# email subject
Label(root, text="subject").grid(row=3, column=0, sticky=E)
email_subject = Entry(root)
email_subject.grid(row=3, column=1)

# attach logo
Label(root, text="attach logo").grid(row=4, column=0, sticky=E)
email_logo = Entry(root)
email_logo.grid(row=4, column=1)
open_logo = Button(root, text="open", command=lambda: add_path_logo(email_logo))
open_logo.grid(row=4, column=2, sticky=W)
Label(root, text="content image id").grid(row=4, column= 3, sticky=E)
cid = Entry(root)
cid.grid(row=4, column=4)

#html content
Label(root, text="html content").grid(row=5, column=0, pady=(15, 0), sticky=NE)
email_html = Text(root)
email_html.grid(row=5, column=1, columnspan=4, padx = (0, 15), pady=(15, 15))

#password enter
Label(root, text="password").grid(row=6, column=0, sticky=E)
password = Entry(root, show="*")
password.grid(row=6, column=1)

#send button
send = Button(root, text="send emails", command=send_mail)
send.grid(row=7, column=2)

#save button
save = Button(root, text="save", command=lambda: save_to_file(None, None, None, None, None))
save.grid(row=8, column=2, pady=(15, 0))

# settings button
open_settings = Button(root, text="open settings", command=lambda: settings_window(file_loaded))
open_settings.grid(row=9, column=2, pady=(15,0))

#log
Label(root, text="log").grid(row=5, column=6, sticky=E)
log = Text(root)
log.grid(row=5, column=7)

scrollb = Scrollbar(root, command = log.yview)
log['yscrollcommand'] = scrollb.set
scrollb.grid(row=1, column=8, rowspan=13)

if os.path.isfile("settings.txt"):
    x, html, file_numbers, file_loaded = load_settings(True)
else:
    x, html, file_numbers, file_loaded = load_settings(False)

#load setings
if file_loaded:
    email_from.insert(END, x['message_settings']['from'])
    email_subject.insert(END, x['message_settings']['subject'])
    email_html.insert(END, html)
    email_logo.insert(END, x['img_settings']['img_file_path'])
    create_file_path(file_numbers)
    for i in range(len(path_of_a)):
        path_of_a[i].insert(END, x['file_settings']['file_path'][i])
    cid.insert(END, x["img_settings"]["img_cid"].replace(" ", ""))

root.mainloop()