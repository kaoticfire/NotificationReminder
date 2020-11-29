from datetime import datetime, timedelta
import schedule
from win32com.client import Dispatch
from os import getenv
import tkinter
from tkinter import ttk
from operator import itemgetter
from logging import getLogger, basicConfig, Formatter, WARNING
from logging.handlers import RotatingFileHandler

__name__ = 'notification_system'


def logging(msg):
    access = 'a'
    filename = getenv('UserProfile') + r'\notifications.log'
    size = 50 * 1024
    basicConfig(WARNING)
    file_handler = RotatingFileHandler(filename, access, maxBytes=size,
                                       backupCount=2, encoding=None, delay=0)
    file_format = Formatter('%(asctime)s - %(levelname)s - %(message)s',
                            datefmt='%Y-%m-%d %H:%M:%S')
    file_handler.setLevel(WARNING)
    file_handler.setFormater(file_format)
    logger = getLogger(__name__)
    logger.addHandler(file_handler)
    logger.warning(msg)


def send_alert_email(site_id, priority):
    outlook = Dispatch('outlook.application')
    msg = outlook.CreateItem(0)
    msg.to = email_recipient.get()
    msg.Subject = site_id + ' ' + priority
    msg.HTMLBody = '<html><body><p>Time for another update!</p></body></html>'
    try:
        msg.Send()
    except Exception:
        logging('Problem sending email')
        pass


def check_alarms():
    for item in watching:
        if item[1] <= datetime.now():
            send_alert_email(item[0], item[2])
            item[1] = item[1] + timedelta(minutes=15)
            display_watched()


def run_schedule():
    schedule.run_pending()
    root.after(1000, run_schedule)


def add_site(a):
    entry = []
    site = site_entry.get()
    time_down = time_entry.get()
    priority = priority_choice.get()
    if site == '':
        return
    if time_down == '' and priority == '':
        logging('Bad timer statement')
        return
    if time_down != '':
        time_down = int(time_down)
    site_choice.set('')
    time_down_choice.set('')
    site_entry.focus()
    if priority == 'P1':
        alarm = datetime.now() + timedelta(minutes=30 - time_down)
    elif priority == 'P2':
        alarm = datetime.now() + timedelta(minutes=240 - time_down)
    elif priority == 'P3':
        alarm = datetime.now() + timedelta(minutes=480 - time_down)
    entry.append(site)
    entry.append(alarm)
    entry.append(priority)
    watching.append(entry)
    display_watched()


def display_watched():
    global watching
    site_list.delete(0, 'end')
    watching = sorted(watching, key=itemgetter(1))
    for item in watching:
        site_list.insert('end', item[0] + ' - ' + item[1].strftime('%H:%M') + ' - ' + item[2])


def remove_alarm():
    try:
        selection = int(site_list.curselection()[0])
        del watching[selection]
        site_list.delete(selection)
    except IndexError:
        logging('Problem removing selection')
        return


root = tkinter.Tk()
root.title('Notification Reminder')
root.geometry('390x260+200+200')

site_choice = tkinter.Variable(root)
priority_choice = tkinter.Variable(root)
time_down_choice = tkinter.Variable(root)
email_recipient = tkinter.Variable(root)

watching = []
email_recipients = ['someone@example.com', 'someoneelse@example.com']
email_recipients = sorted(email_recipients)

entry_frame = tkinter.Frame(root)
entry_frame.grid(row=0, column=0)
tkinter.Label(entry_frame, text='Site:').grid(row=0, column=0)
site_entry = ttk.Entry(entry_frame, textvariable=site_choice, width=8)
site_entry.grid(row=0, column=1)
tkinter.Label(entry_frame, text='Minutes Down:').grid(row=0, column=2)
time_entry = ttk.Entry(entry_frame, textvariable=time_down_choice, width=8)
time_entry.grid(row=0, column=3)
priority_choice.set('Priority')
priority = tkinter.OptionMenu(entry_frame, priority_choice, 'P1', 'P2', 'P3')
priority.grid(row=0, column=4)
addButton = tkinter.Button(entry_frame, text='Add', command=lambda: add_site(watching), default='active')
addButton.grid(row=0, column=5)

list_frame = tkinter.Frame(root)
list_frame.grid(row=1, column=0)
site_list = tkinter.Listbox(list_frame, width=50)
site_list.grid(row=0, column=0)
delete_button = tkinter.Button(list_frame, text='Remove', command=remove_alarm)
delete_button.grid(row=2, column=0)

mail_frame = tkinter.Frame(root)
mail_frame.grid(row=2, column=0)
mail_recipient = tkinter.OptionMenu(mail_frame, email_recipient, *email_recipients)
mail_recipient.grid(row=0, column=0)

schedule.every(30).seconds.do(check_alarms)

root.after(1000, run_schedule)

site_entry.bind('<Return>', add_site)
time_entry.bind('<Return>', add_site)
site_entry.focus()
root.mainloop()
