'''
Name: eventmailer.py
Authors: Roman Kapitoulski, Eric Russon, Maryam Bunama
Version: 0.0
Date: June 24, 2023 
Description: This is a script that sends pre-scheduled emails that are stored in an Excel file.
'''

import win32com.client as win32
import datetime, os

def Emailer(subject, body, recipient):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)    
    mail.To = recipient
    mail.Subject = subject
    mail.Body = body

def Scheduler(time):
    currentTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    sendTime = time

    if sendTime == currentTime:
        Emailer()
        print(f"Mail was sent at {currentTime}.")
    
    else:
        pass


excel = win32.gencache.EnsureDispatch('Excel.Application')
filepath = os.getcwd()
wb = excel.Workbooks.Open(f'{filepath}\eventlist.xlsx')

readEvents = wb.Worksheets('Events')
allEvents = readEvents.UsedRange
print(f'Data on selected sheet : {allEvents}')

readEmails = wb.Worksheets('MailingList')
allEmails = readEmails.UsedRange
print(allEmails)

# for event in allEvents:
#     print(event)
#     print()


#figure out how to split nested tuples into lists 
#append list info to Emailer()
#split date for correct formatting
