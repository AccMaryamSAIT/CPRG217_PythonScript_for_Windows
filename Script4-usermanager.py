'''
Name: Script4-usermanager.py
Authors: Roman Kapitoulski, Eric Russon, Maryam Bunama
Version: 1.1
Date: July 14, 2023
Description: This script allows bulk user management by adding and deleting users from an excel file as well as being able to
view all users accounts on the system. As of now, the script offers no control over adding and deleting specific users. Users originally existing
on the system are unaffected by the deletion.
Note: Ensure that this program is run on CMD or Powershell as an administrator.
'''

import win32com.client as win32
import win32net
import win32netcon
import os

## FUNCTIONS

def dispatchExcelFile(filepath):
    # Dispatch excel file and open the workbook
    excelApp = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excelApp.Workbooks.Open(filepath)

    # Exctract the worksheet into a variable to allow retrieval of information
    ws1 = wb.Worksheets(1)

    # Return ws1 for usage in other functions. Return wb and excelApp to allow closing of the excel file after it served its purpose.
    return ws1, wb, excelApp

def retrieveUserInfo():
    # Get the file path of the excel file. Excel file must be in the same directory as the .py file.
    excelPath = f'{os.getcwd()}/Net Users.xlsx'

    # Call the dispatch function and assign the variables 
    ws1, wb, excelApp = dispatchExcelFile(excelPath)

    # Read only the filled in cells in the file
    ws1 = ws1.UsedRange

    # Strip our the header in order to only use the data
    i = ws1.Rows.Count - 1
    n = 0

    # Create empty lists to add the user information to.
    usernames = []
    passwords = []

    # Loop to extract data and append to the empty lists.
    while i > 0:
        # Excel cell values that change to the next value with every loop
        name = ws1[3+n].Value
        password = ws1[4+n].Value
        usernames.append(name)
        passwords.append(password)
        n += 2
        i -= 1

    # Close and quit the excel file.
    wb.Close()
    excelApp.Quit()

    # Return the populated lists
    return usernames, passwords
    
def createUsers():
    # Retreive data and assign variables to the populated lists
    usernames, passwords = retrieveUserInfo()
    
    # Iterate through each user and append information to a dictionary.
    for username, password in zip(usernames, passwords):  
        user_info = {
            # The following key-value pairs are required parameters for the NetUserAdd function
            'name': username,
            'password': password,
            'password_age': 0,
            'priv': win32netcon.USER_PRIV_USER,
            'home_dir': None,
            'comment': None,
            'flags': win32netcon.UF_SCRIPT,
            'script_path': None
        }
        
        try: 
            # Create each user with the information from the dictionary.
            win32net.NetUserAdd(None, 1, user_info)

            # Print a conformation if the process succeeded.
            print(f'User "{username}" created successfully.')

        except Exception as e:
            # Print an error if the process failed.
            print(f'Failed to create user {e}')
            
def deleteUsers():
    # Retrieve only the usernames list. The '_' is used to explicitly ignore the second variable
    usernames, _ = retrieveUserInfo()

    # Iterate through the list and retrieve usernames
    for username in usernames:
        try:
            # Delete each user in the list
            win32net.NetUserDel(None, username)

            # Print a conformation if the process succeeded.
            print(f'User "{username}" deleted successfully.')

        except Exception as e:
            # Print an error if the process failed.
            print(f'Failed to delete user: {e}')
        
def viewUsers():
    # Enumerate through the users and extract their information
    try:
        resumeHandle = 0
        while True:
            users, _, resumeHandle = win32net.NetUserEnum(None, 2, win32netcon.FILTER_NORMAL_ACCOUNT, resumeHandle) # The '_' is in place of the total number of users which we are not using
            for user in users:
                # Print each username and place a divider inbetween users.
                print(f'Username {user["name"]}')
                print(f'-----------------------------')
            if not resumeHandle:
                break
    except Exception as e:
        # Print an error if process failed.
        print(f'Failed to view users: {e}')

def Menu():
    # Use a while loop for the menu that will repeat until the user decides to exit out of it.
    x = False
    while x != True:
        print('Welcome to the bulk user manager!')
        print('Select from the following options or select 4 to exit.')
        print('1- View Users')
        print('2- Add Users')
        print('3- Delete Users')
        print('4- exit')
        option = int(input('>>>> '))
        if option == 1:
            viewUsers()
        elif option == 2:
            createUsers()
        elif option == 3:
            deleteUsers()
        elif option == 4:
            x = True
            exit()
        else:
            print('Invalid option selected.')

## PROGRAM STARTS

Menu()