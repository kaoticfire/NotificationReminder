# NotificationReminder
A modified version of the Site-Down-Acker

This software sends email reminders for issues that have arisen.  
The emails are sent for issues that have arisen and are based off  
criteria chosen by one of three preset intervals.

## Requirements
To user this software, python 3.8 or higher (32-bit version) is required  
along with two packages (pywin32 & schedule). These packages can be  
installed using the following command.
    “pip install <package_name>”
To see what version is installed simply open a command prompt and  
type “python –version”, if an error is received then python is not  
installed, if the version is less that, if an error is received then  
python is not installed, if the version is less than 3.8, python will need  
to be uninstalled and an updated version installed (as python does not  
have an in-place upgrade option). A configured outlook client is  
required for the software to send emails, and a windows operating system.

## Setup
To use this software, a line of code is to be edited. Open the “.py” file  
and search for the following line:
    “msg.to = ‘someone@example.com’”
Replace the email listed in the file with that of the desired email  
to receive notifications. The software can be displayed in a light or  
a dark theme.
