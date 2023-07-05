'''
Name: Script2-bedtimealarm.py
Authors: Roman Kapitoulski, Eric Russon, Maryam Bunama
Version: 1.2.1
Date: July 4, 2023
Description: This script has been made to run in the background and monitoring a kid's PC hours. Once he reaches the time limit,
he will see a pop-up dialog box with two options. If he chooses 'Yes', the computer will shut down after 5 mintues. If he chooses 'no',
he will be able to continue using the computer but an email will be sent to the parent. 
'''

import winsound # Module to play a Windows Alarm
import threading # Module that allows alarm to start at the same time as the dialog box
import win32com.client as win32 # Module to access excel configuration file and send email
import win32api # Module to use the popup dialog box
import win32gui # Module to display the popup dialog box
import win32con # Module that contains the constants for the dialog box options
import os, time, datetime, schedule

## EXCEPTION
# Built in exception to stop the scheduler if conditions are met
class DialogFinished(Exception):
    pass

## FUNCTIONS 
def playWindowsAlarm():
    # Set alarm name based on Registry editor key naming.
    alarmName = "Notification.Looping.Alarm"

    # Play alarm 
    winsound.PlaySound(alarmName, winsound.SND_ALIAS)

def formatDecimalTime(decimalTime):
    # Convert decimal time to hours and minutes
    hours = int(decimalTime * 24)
    minutes = int((decimalTime * 24 * 60) % 60)

    # Create a time object with the converted hours and minutes
    timeObj = datetime.time(hours, minutes)

    # Format the time object as HH:MM
    formattedTime = timeObj.strftime("%H:%M")
    return formattedTime

def dispatchExcelFile(filepath):
    # Dispatch the file and return the excel file and worksheet variables
    excelApp = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excelApp.Workbooks.Open(filepath)
    ws1 = wb.Worksheets(1)
    return ws1

def Emailer(parentEmail):
    # Dispatch outlook
    outlook = win32.Dispatch('outlook.application')

    # Create an email
    mail = outlook.CreateItem(0)

    # Setting email recipient, subject, and body
    mail.To = parentEmail
    mail.Subject = 'Child chose to ignore the alarm'
    mail.Body = f'You are recieveing this email because your child chose to ignore the alarm you previously set.'

    # Sends email to reciepient
    mail.Send()

def displayDialog(email):

    # Define the dialog box properties
    title = "Time is Up!"
    message = "You reached your PC time limit. Do you want to shut down?"

    # Create a thread for playing the alarm sound. This allows the alarm and the dialog box to play at the same time
    alarmThread = threading.Thread(target=playWindowsAlarm)
    alarmThread.start() # Start the alarm sound thread
    
    # Create the dialog box
    response = win32api.MessageBox(win32gui.GetForegroundWindow(), message, title, win32con.MB_ICONWARNING | win32con.MB_YESNO)
    alarmThread.join() # Join the alarm sound thread to wait for it to finish

    # Handle the user's response
    if response == win32con.IDYES: # User clicks 'Yes'
        print("Shutting down the computer in 5 minutes")
        # time.sleep (5) # FOR DEMO PURPOSES - wait for 5 seconds
        time.sleep(5 * 60) # Wait for 5 minutes
        os.system("shutdown /s /t 0") # Shut down the computer
        raise DialogFinished # Trigger the custom exception

    elif response == win32con.IDNO: # Users clicks 'No'
        print("Ignoring the alert.")
        print("Sending email...")
        Emailer(email) # Send the email
        raise DialogFinished # Trigger the custom exception

def runInBackground():
    filePath = f'{os.getcwd()}/alarmsettings.xlsx' # get excel settings filepath

    try:
        ws1 = dispatchExcelFile(filePath)  # Dispatch excel 
    except:
        print('Settings file was not found! Please make sure a settings file was created.')

    email = ws1.Range('E2').Value # Parent email address

    # Use veto time if it's available. Otherwise, use set time
    if ws1.Range('C2').Value == None:
        timeLimit = ws1.Range('A2').Value
    else: 
        timeLimit = ws1.Range('C2').Value

    # Format the time limit to hh:mm for the scheduler to work
    formattedTimeLimit = formatDecimalTime(timeLimit)

    # Schedule the displayDialog function to run at the specified time
    schedule.every().day.at(formattedTimeLimit).do(displayDialog, email)

    # Keep the program running
    while True:
        schedule.run_pending()
        time.sleep(1)

# Run program indefinitely until the dialog box pops up. If it does, the custom exception will execute and stop the program.
try:
    runInBackground()
except DialogFinished:
    print("Dialog finished, exiting script")