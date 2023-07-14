'''
Name:usermanager.py
Authors: Roman Kapitoulski, Eric Russon, Maryam Bunama
Version: 1.0
Date: July 14, 2023
Description: Viewing, Adding, and Deleting Users on a PC
'''
import win32api
import win32net
import win32com.client as win32
import win32netcon
import os

def dispatchExcelFile(filepath):
    excelApp = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excelApp.Workbooks.Open(filepath)
    ws1 = wb.Worksheets(1)
    return ws1

def RetUserInfo():
    excelPath = f'{os.getcwd()}/Net Users.xlsx'

    ws1 = dispatchExcelFile(excelPath)
    ws1 = ws1.UsedRange
    i = ws1.Rows.Count - 1
    n = 0  
    while i > 0:
        # Excel cell values that change to the next value with every loop
        name = ws1[3+n].Value
        password = ws1[4+n].Value
        
        n += 2
        i -= 1
        return name, password

def create_user():
    username, password = RetUserInfo()
    print(username, password)
    user_info = {
        'name': username,
        'password': password
    }
    
    # try: 
    #     # win32net.NetUserAdd(None, 1, user_info)
    #     # print(f"User '{username}' created successfully.")
    # except Exception as e:
    #     print(f"Failed to create user {e}")
        
def delete_user():
    username, password = RetUserInfo()
    try:
        win32net.NetUserDel(None, username)
        print(f"User '{username}' deleted successfully.")
    except Exception as e:
        print(f"Failed to delete user: {e}")
        
def view_users():
    try:
        resume_handle = 0
        while True:
            users, total, resume_handle = win32net.NetUserEnum(None, 2, win32netcon.FILTER_NORMAL_ACCOUNT, resume_handle)
            for user in users:
                print(f"Username {user['name']}")
                print(f"Full Name: {user['full_name']}")
                print(f"------------------------")
            if not resume_handle:
                break
    except Exception as e:
        print(f"Failed to view users: {e}")

def Menu():
    x = 0
    while x !=3:
        print('1: Get User')
        print('2: Add User')
        print('3: Delete User')
        print('4: exit')
        option = int(input('Type option here:'))
        if option == 1:
            view_users()
        elif option == 2:
            create_user()
        elif option == 3:
            delete_user()
        elif option == 4:
            exit()
        else:
            print('Invalid option selected')

Menu()