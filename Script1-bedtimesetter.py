'''
Name: Script1-bedtimesetter.py
Authors: Roman Kapitoulski, Eric Russon, Maryam Bunama
Version: 0.0
Date: June 29, 2023
Description: This script has been made for a parent to set an alarm for children when their bedtime approaches. 
It has the options to set the timer, view the time set, veto the timer (for a period of time), save and exit, 
or save without exiting.
'''
option = 0
while option != 5:
    print('\nWelcome to Bedtime Alarm Setter!')
    print('Select from the following options:')
    print('1- Set timer')
    print('2- View configurations')
    print('3- Veto timer')
    print('4- Save and exit')
    print('5- Exit without saving')
    
    option = int(input('>>> '))
    if option == 1:
        # code to set the timer
        print('\nTimer set')
    elif option == 2:
        # code to view timer and configurations
        timer = 0 # TEMPORARY VALUE
        print(f'\nTimer set to: {timer}')
    elif option == 3:
        # code to veto
        print('\nTimer vetoed')
    elif option == 4:
        # code to write changes to excel file
        print('\nSaved and exiting program...\n')
        exit()
    elif option == 5:
        print('\nExiting program...\n')
        exit()
    else:
        print('\nInvalid input. Please input the correct values.')