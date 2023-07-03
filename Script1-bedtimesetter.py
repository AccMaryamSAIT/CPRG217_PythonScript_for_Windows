'''
Name: Script1-bedtimesetter.py
Authors: Roman Kapitoulski, Eric Russon, Maryam Bunama
Version: 1.2
Date: July 3, 2023
Description: This script has been made for a parent to set an alarm for children when their bedtime approaches. 
It has the options to set the timer, view the time set, veto the timer (until a set time limit), save and exit, 
or save without exiting. The time format used corresponds to 'hh:mm AM/PM' throughout. The code won't accept vetoing the time 
if no time was previously set.
'''

import win32com.client as win32 # Module to create and edit excel file in this script
from tabulate import tabulate # Module that displays data in a table
import getpass, os, datetime 

## FUNCTIONS
def getPassword():
    # Prompt user to enter password
    password = getpass.getpass('Enter password: ')
    return password

def authenticate(password):
    # Preset password to compare against inputted password
    return password == '1'

def formatDecimalTime(decimalTime):
    # Convert decimal time to hours and minutes
    total_minutes = decimalTime * 24 * 60  # Convert to total minutes in a day
    hours = int(total_minutes // 60)  # Extract the whole number of hours
    minutes = int(total_minutes % 60)  # Extract the remaining minutes

    # Create a time object with the converted hours and minutes
    timeObj = datetime.time(hours, minutes)

    # Format the time object as %I:%M %p
    formattedTime = timeObj.strftime('%I:%M %p')
    
    return formattedTime

def createExcelFile():
    # Dispatch excel file and create a workbook
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Add()
    
    # Name the worksheet
    ws1 = wb.Worksheets(1)
    ws1.Name = 'Settings'

    # Assigning static values to header cells
    headerValues = {
    'A1': 'Time Limit',
    'B1': 'Veto Status',
    'C1': 'Veto Until',
    'D1': 'Date set',
    'E1': 'Parent Email' }

    for headerAddress, headerValue in headerValues.items():
        ws1.Range(headerAddress).Value = headerValue

    # Save with a preset filename
    wb.SaveAs(os.path.join(f'{os.getcwd()}/alarmsettings.xlsx'))
    wb.Close()    

def dispatchExcelFile(filepath):
    # Dispatch the file and return the excel file and worksheet variables
    excelApp = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excelApp.Workbooks.Open(filepath)
    ws1 = wb.Worksheets(1)
    return ws1

def setEmail(filepath):
    # Dispatch the file
    ws1 = dispatchExcelFile(filepath)

    # Input prompt
    print('\nRemember to save changes after setting your email.')
    email = input('Please input your email\n>>> ')

    # Adding input to excel file
    cellAddress = 'E2'
    cellValue = email
    ws1.Range(cellAddress).Value = cellValue

def setTime(filepath):
    # Dispatch excel file
    ws1 = dispatchExcelFile(filepath)

    # Input prompt
    print('\nInput time in the "hh:mm AM/PM" format.')
    timeLimit = input('>>> ')

    # Format time and print error if user inputs the wrong information
    try:
        timeObj = datetime.datetime.strptime(timeLimit, '%I:%M %p')
        timeFormatted = timeObj.strftime('%I:%M %p')
    except ValueError:
        print('Invalid time format. Please use the "hh:mm AM/PM" format.')
        return

    # Update cell values
    cellValues = {
    'A2': timeFormatted,
    'B2': False,
    'C2': None,
    'D2': datetime.datetime.now().strftime('%b/%d/%Y %I:%M %p')
    }

    # Add the data to the specified cells in excel
    for cell in cellValues.items():
        cellAddress = cell[0]
        cellValue = cell[1] 
        ws1.Range(cellAddress).Value = cellValue
    
    print('\nSetting time...')

def setVetoTime(filepath):
    # Dispatch excel file
    ws1 = dispatchExcelFile(filepath)

    # Print error if no time was previously set
    timeCell = ws1.Range('A2').Value
    if timeCell is None:
        print('\nOperation can\'t be performed. Please set the time limit first (Option: 1).')
        return

    print('\nInput veto time in the "hh:mm AM/PM" format')
    timeVeto = input('>>> ')

    # Format time and print error if user inputs the wrong information
    try:
        timeObj = datetime.datetime.strptime(timeVeto, '%I:%M %p')
        timeFormatted = timeObj.strftime('%I:%M %p')
    except ValueError:
        print('Invalid time format. Please use the "hh:mm AM/PM" format.')
        return

    # Update cell values
    cellValues = {
    'B2': True,
    'C2': timeFormatted,
    'D2': datetime.datetime.now().strftime('%M/%d/%Y %I:%M %p')
    }

    # Add the data to the specified cells in excel
    for cell in cellValues.items():
        cellAddress = cell[0]
        cellValue = cell[1] 
        ws1.Range(cellAddress).Value = cellValue
    
    print('\nVetoing time...')

def viewTimeSettings(filepath):   
    # Dispatch excel file
    ws1 = dispatchExcelFile(filepath)

    # Get the used range of the worksheet and convert it into a formatted Excel table
    usedRange = ws1.UsedRange
    tableData = usedRange.Value

    # Convert the table data to a list of rows
    rows = [list(row) for row in tableData]
    
    # Extract the header row and remove it from the rows list
    headers = rows.pop(0)

    # Only format the specified cells if time was inputted. 
    timeCell = ws1.Range('A2').Value
    if timeCell is None:
        pass
    else:
        # Format the time cell for time limit
        timeSetCell = rows[0][0]
        formattedTimeSetCell = formatDecimalTime(timeSetCell)
        rows[0][0] = formattedTimeSetCell

    timeCell = ws1.Range('C2').Value
    if timeCell is None:
       pass
    else:
        # Format the time cell for time veto 
        timeVetoCell = rows[0][2]
        formattedTimeVetoCell = formatDecimalTime(timeVetoCell)
        rows[0][2] = formattedTimeVetoCell


    # Display the table in the terminal
    print(tabulate(rows, headers=headers, tablefmt='fancy_grid'))

def saveExcelFile(filepath):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(filepath)
    wb.Close(SaveChanges=True) # Writes changes to excel file
    excel.Quit()

def exitExcelFileWithoutSaving(filepath):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(filepath)
    wb.Close(SaveChanges=False) # Ensures that no changes are written to excel file
    excel.Quit()

def menu():
    option = 0
    while option != 5:
        # Takes the filepath of the excel settings file
        filepath = f'{os.getcwd()}/alarmsettings.xlsx'

        # Menu printed out with list of options. Each option corresponds to a number
        print('\nWelcome to Bedtime Alarm Setter!')
        print('Select from the following options:')
        print('1- Set timer')
        print('2- View configurations')
        print('3- Veto timer')
        print('4- Set Email')
        print('5- Save and exit')
        print('6- Exit without saving')
        
        option = int(input('>>> '))
        if option == 1:
            setTime(filepath) # Function for setting the time. 
        elif option == 2:
            viewTimeSettings(filepath) # Function to display table of time settings
        elif option == 3:    
            setVetoTime(filepath) # Function for setting time to veto
        elif option == 4:       
            setEmail(filepath) # Function for setting the email
        elif option == 5:
            saveExcelFile(filepath) # Save excel file and exit
            print('\nSaved and exiting program...\n')
            exit()
        elif option == 6:
            exitExcelFileWithoutSaving(filepath) # Exiting program without saving changes
            print('\nExiting program...\n')
            exit()
        else:
            print('\nInvalid input. Please input the correct values.')

def main():
    # Takes the filepath of the excel settings file
    filepath = f'{os.getcwd()}/alarmsettings.xlsx'

    # Create a new file if program file doesn't exist or was deleted
    if os.path.isfile(filepath):
        pass
    else:  
        createExcelFile()

    # Assigns inputted password to a variable
    password = getPassword()

    # Authenticates password against the preset password
    if authenticate(password):  
        # Launches menu if authentication successful
        print('Authentication successful!')
        menu()
    else:
        print('Authentication failed. Access denied.')

# Launch the program
main()