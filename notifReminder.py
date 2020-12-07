import tkinter as tk
import webbrowser as wb
from datetime import datetime, timedelta
from logging import getLogger, basicConfig, Formatter, WARNING
from logging.handlers import RotatingFileHandler
from operator import itemgetter
from os import getenv
from time import sleep

from schedule import every, run_pending
from win32com.client import Dispatch

__name__ = 'notification_system'


def logging(msg):
    access = 'a'
    filename = getenv('AppData') + r'\notifications.log'
    size = 50 * 1024
    basicConfig(level=WARNING)
    file_handler = RotatingFileHandler(filename, access, maxBytes=size,
                                       backupCount=2, encoding=None, delay=False)
    file_format = Formatter('%(asctime)s - %(levelname)s - %(message)s',
                            datefmt='%Y-%m-%d %H:%M:%S')
    file_handler.setLevel(WARNING)
    file_handler.setFormatter(file_format)
    logger = getLogger(__name__)
    logger.addHandler(file_handler)
    logger.warning(msg)
    sleep(1)
    wb.open(filename)


def send_alert_email(site_id, priority):
    outlook = Dispatch('outlook.application')
    msg = outlook.CreateItem(0)
    msg.to = 'someone@example.com'
    msg.Subject = site_id + ' ' + priority
    msg.HTMLBody = '<html><body><p>Time for another update!\nIt is important to' \
                   'SLA\'s and the customer\nthat an update be sent promptly!' \
                   '</p></body></html>'
    try:
        msg.Send()
    except Exception:
        logging('Problem sending email')
        pass


def check_alarms():
    for item in watching:
        if item[1] <= datetime.now():
            send_alert_email(item[0], item[2])
            item[1] = item[1] + timedelta(minutes=30)
            display_watched()


def run_schedule():
    run_pending()
    root.after(1000, run_schedule)


def add_site(_):
    entry = []
    site = site_entry.get()
    time_down = time_entry.get()
    priority = priority_choice.get()
    alarm = ''
    if site == '':
        return
    if time_down == '' or priority == '':
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


def focus_next(event):
    event.widget.tk_focusNext().focus()
    return ('break')


def change_mode(tog=[0]):
    tog[0] = not tog[0]
    frames = (entry_frame, list_frame)
    widgets = (site_label, time_label, site_entry, time_entry, priority_c,
               add_button, site_list, delete_button, priority_c['menu'],
               change_button)
    nbg = '#000000'
    nfg = '#66FFFF'
    dbg = '#FFFFFF'
    dfg = '#000000'

    if tog[0]:
        root.option_add('*Background', dbg)
        root.option_add('*Foreground', dfg)
        root.configure(background=dbg, highlightbackground=dbg, highlightcolor=dfg)

        for i in frames:
            i.configure(background=dbg, highlightbackground=dbg, highlightcolor=dfg)

        for i in widgets:
            i.configure(background=dbg, foreground=dfg)
    else:
        root.option_add('*Background', nbg)
        root.option_add('*Foreground', nfg)
        root.configure(background=nbg, highlightbackground=nbg, highlightcolor=nfg)

        for i in frames:
            i.configure(background=nbg, highlightbackground=nbg, highlightcolor=nfg)

        for i in widgets:
            i.configure(background=nbg, foreground=nfg)


root = tk.Tk()
root.title('Notification Reminder')
root.resizable(False, False)
root.focusmodel('active')
root.geometry('310x230+200+200')

site_choice = tk.Variable(root)
priority_choice = tk.Variable(root)
time_down_choice = tk.Variable(root)
email_recipient = tk.Variable(root)

watching = []

entry_frame = tk.Frame(root)
site_label = tk.Label(entry_frame, text='Site:')
site_entry = tk.Entry(entry_frame, textvariable=site_choice, width=8)
time_label = tk.Label(entry_frame, text='Time Down:')
time_entry = tk.Entry(entry_frame, textvariable=time_down_choice, width=8)
priority_choice.set('Priority')
priority_c = tk.OptionMenu(entry_frame, priority_choice, 'P1', 'P2', 'P3')
priority_c['highlightthickness'] = 0

entry_frame.grid(row=0, column=0)
site_label.grid(row=0, column=0, sticky='e')
site_entry.grid(row=0, column=1, padx=5)
time_label.grid(row=0, column=2)
time_entry.grid(row=0, column=3, padx=5)
priority_c.grid(row=0, column=4, padx=5)

list_frame = tk.Frame(root)
site_list = tk.Listbox(list_frame, width=50)
add_button = tk.Button(list_frame, text='Add', borderwidth=0.5,
                       command=lambda: add_site(watching))
delete_button = tk.Button(list_frame, text='Remove', borderwidth=0.5,
                          command=remove_alarm)
change_button = tk.Button(list_frame, text='Change', borderwidth=0.5,
                          command=change_mode)

list_frame.grid(row=1, column=0)
site_list.grid(row=0, column=0, columnspan=3)
add_button.grid(row=2, column=0, padx=5, pady=5)
delete_button.grid(row=2, column=1, padx=5, pady=5)
change_button.grid(row=2, column=2, padx=5, pady=5)

every(30).seconds.do(check_alarms)
root.after(1000, run_schedule)
site_entry.bind('<Tab>', focus_next)
site_entry.bind('<Return>', add_site)
time_entry.bind('<Return>', add_site)
site_entry.focus()

root.mainloop()
