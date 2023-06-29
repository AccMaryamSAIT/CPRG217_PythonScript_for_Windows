'''
Name: Script3-scholarshipmailer.py
Authors: Roman Kapitoulski, Eric Russon, Maryam Bunama
Version: 1.0
Date: June 29, 2023 
Description: This is a script that sends an acceptance email to student who applied for the CPRG-217 scholarship.
It reads the student information in an Excel file applies it to custom fields in a Word document. Then, it sends
an email customized for each scholarship reciever.
'''

import win32com.client as win32 # Module that allows dispatching and working with Excel, Word, and Outlook for this script
import os, datetime 

## FUNCTIONS
def replaceVariablesInWord(content, variables):
    # Create a copy of the word document's content
    replacedContent = content

    # Replace the word document variable with the desired value relying on a dictionary
    for variable, value in variables.items():
        placeholderText = f'<{variable}>'
        replaceText = str(value)
        replacedContent = replacedContent.replace(placeholderText, replaceText)

    # Return the modified content
    return replacedContent

def decimalToTime(excelTimeDecimal):
    # Convert Excel decimal time to timedelta object
    timeDelta = datetime.timedelta(days=excelTimeDecimal)

    # Create a datetime object with today's date and the time delta
    baseDatetime = datetime.datetime.combine(datetime.date.today(), datetime.time())

    # Add the time delta to the base datetime
    resultDatetime = baseDatetime + timeDelta

    # Extract the time component from the result and format it as "hh:mm AM/PM"
    time = resultDatetime.time()
    formattedTime = time.strftime('%I:%M %p')

    return formattedTime 

# Emailer function with required values
def Emailer(body, recipient):
    # Dispatch outlook
    outlook = win32.Dispatch('outlook.application')

    # Create an email
    mail = outlook.CreateItem(0)

    # Setting email recipient, subject, and body
    mail.To = recipient
    mail.Subject = 'You have been accepted! - CPRG217 Scholarship'
    mail.Body = body

    # Sends email to reciepient
    mail.Send()

## FILE PATHS
# current working directory filepath. 
# Script must be in the same directory as the excel and word doc file used.
filepath = os.getcwd()

# Specify the path to the template file
excelPath = f'{os.getcwd()}/recievers.xlsx'
wordPath = f'{os.getcwd()}/scholarship.docx'

## RETRIEVE EXCEL FILE INFORMATION
# Dispatch excel 
excel = win32.gencache.EnsureDispatch('Excel.Application')

# Get working dirctory to open excel file. 
try: 
    wb = excel.Workbooks.Open(excelPath)

except Exception as e:
    error_message = f'An error occured: {str(e)}'
    print('-' * len(error_message), f'\n{error_message}\nFile not found.\n', '-' * len(error_message))
    exit()

# Read events sheet and assign to a variable
try:
    readEvents = wb.Worksheets('Events')
    allEvents = readEvents.UsedRange

except Exception as e:
    error_message = f'An error occured: {str(e)}'
    print('-' * len(error_message), f'\n{error_message}\nWorksheet not found.\n', '-' * len(error_message))
    exit()

## RETRIEVE WORD DOC INFORMATION
# Dispatch word
word = win32.gencache.EnsureDispatch('Word.Application')

# Read word document assign it to a variable
try:
    doc = word.Documents.Open(wordPath)
    docContent = doc.Content.Text

except Exception as e:
    error_message = f'An error occured: {str(e)}'
    print('-' * len(error_message), f'\n{error_message}\nFile not found.\n', '-' * len(error_message))
    exit()


## ASSIGN VARIABLES TO EXCEL INFORMATION
# Iterate through each row, collect, and format information.
i = allEvents.Rows.Count - 1
n = 0   
while i > 0:
    # Excel cell values that change to the next value with every loop
    FirstName = allEvents[7+n].Value
    LastName = allEvents[8+n].Value
    Date = allEvents[9+n].Value.strftime('%B %d,%Y') #Formatting the excel date to 'Month, dd, yyyy'
    Time = decimalToTime(allEvents[10+n].Value) 
    Location = allEvents[11+n].Value
    Email = allEvents[12+n].Value

    # Define the variables to replace
    variables = {
        'FirstName': FirstName,
        'LastName': LastName,
        'Date': Date,
        'Time': Time,
        'Location': Location,
        'Email': Email
    }

    # Replace the variables in the Word document
    emailBody = replaceVariablesInWord(docContent, variables)
    
    Emailer(emailBody, Email)

    # Commented out code allows printing of the email's body content.  
    # Allows for doublechecking content before emailing it. 
    # for line in emailBody.split('\r'):
    #     print(line)
    
    n += 6
    i -= 1


## Closing and quitting files
# Close the workbook and Excel application
wb.Close()
excel.Quit()

# Close and exit Word application
doc.Close()
word.Quit()