'''
Name: script5-pcmonitor.py
Authors: Roman Kapitoulski, Eric Russon, Maryam Bunama
Version: 1.1 
Date: July 15, 2023
Description: Set limits and monitor GPU temperature, memory usage and disk usage.
Display a GUI error if limits reached.
Retrieve information about GPU, memory usage, and disk usage. 
'''

import psutil # Obtians information about Memory and Disk
import GPUtil # Obtains information about Nvidia GPUs
import win32gui # Contains Windows GUI elements
import win32con # Contains extra Windows GUI elements
import time # Time method to control how often to check for limits


def getGPUTemp():
    gpus = GPUtil.getGPUs()
    if gpus:
        # Return the first GPU temperature
        gpu = gpus[0]
        return gpu.temperature

    return None


def getMemoryUsage():
    memory = psutil.virtual_memory()
    return memory.percent


def getDiskUsage(diskPath):
    try:
        disk = psutil.disk_usage(diskPath)
    except FileNotFoundError:
        disk = psutil.disk_usage("/")
    return disk.percent


def showErrorMessage(message):
    result = win32gui.MessageBox(0, message, "System Monitor Alert", win32con.MB_ICONERROR | win32con.MB_OKCANCEL)
    # If user clicks cancel button then exit the program, otherwise if clicks OK the program will continue to monitor
    if result == win32con.IDCANCEL:
        exit(0)


def monitorLimits(gpuLimit, memoryLimit, diskPath, diskLimit):
    print("Monitoring initiated...")
    while True:
        gpuTemp = getGPUTemp()
        memoryUsage = getMemoryUsage()
        diskUsage = getDiskUsage(diskPath)

        # Display instructions on how the user should continue
        instructions = "Please make required changes to lower the limit and then click OK. " \
                       "To quit this program click Cancel."

        if gpuTemp is not None and gpuTemp > gpuLimit:
            showErrorMessage(f"GPU temperature exceeded {gpuLimit}°C! " + instructions)

        if memoryUsage > memoryLimit:
            showErrorMessage(f"Memory usage exceeded {memoryLimit}%! " + instructions)

        if diskUsage > diskLimit:
            showErrorMessage(f"Disk usage exceeded {diskLimit}%! " + instructions)

        print("GPU Temperature: " + str(gpuTemp) + "°C")
        print("Memory Usage: " + str(memoryUsage) + "%")
        print("Disk Usage: " + str(diskUsage) + "%")

        time.sleep(10)  # Wait for 10 seconds before checking limits again


def promptUser():
    # Try except blocks to ensure proper values are entered

    try:
        gpuLimit = int(input("Enter the GPU temperature limit in °C (Defaults 95°C): "))
    except ValueError:
        print("Invalid value. Entering default value")
        gpuLimit = 95

    try:
        memoryLimit = int(input("Enter the memory usage percentage limit (Defaults 85%): "))
    except ValueError:
        print("Invalid value. Entering default value")
        memoryLimit = 85

    try:
        diskPath = input("Enter the path of the disk (Default \"/\"): ")
    except ValueError:
        print("Invalid value. Entering default value")
        diskPath = "/"

    try:
        diskLimit = int(input("Enter the disk usage percentage limit (Default 95%): "))
    except ValueError:
        print("Invalid value. Entering default value")
        diskLimit = 95

    return gpuLimit, memoryLimit, diskPath, diskLimit


def main():
    # Introduction to the program
    print("Welcome to the MER System Monitoring Program")
    print("This program is intended to run on a Windows machine with a Nvidia GPU")

    # Prompt user to see if they have the requirements to run this program
    windows = input("Do you have a Windows machine? (y/n): ")
    while windows.lower() != "y" and windows.lower() != "n":
        print("Please enter y or n")
        windows = input("Do you have a Windows machine? (y/n): ")

    nvidia = input("Do you have a Nvidia GPU? (y/n): ")
    while nvidia.lower() != "y" and nvidia.lower() != "n":
        print("Please enter y or n")
        nvidia = input("Do you have a Nvidia GPU? (y/n): ")

    if windows != "y" and nvidia != "y":
        print("Sorry this program won't work for you")
        exit(0)

    print("Enter desired values, otherwise press enter for default value.")

    # Collect limit values
    gpuLimit, memoryLimit, diskPath, diskLimit = promptUser()

    # Begin monitoring values
    monitorLimits(gpuLimit, memoryLimit, diskPath, diskLimit)


# If the python script file name is main.py then run this program
if __name__ == '__main__':
    main()
