import os
import keyboard
from tkinter import Tk
from tkinter import filedialog


DIR = os.path.abspath("./")
TEMPDIR = os.path.abspath("./temp")
DELLOUTPUTDIR = os.path.abspath("./dellWarrantyOutput")
SCHOOLS = ('BBE', 'BES', 'BHS', 'BPE', 'CRE', 'EBW', 'EMS', 'GBE', 'GES', 'GHS', 'GWE', 'HBW', 'HES', 'HHS', 'HMS', 'ILE', 'JAE',
               'JWW', 'KDDC', 'LCE', 'LCH', 'LCM', 'LPE', 'MCE', 'MES', 'MHM', 'MTC', 'NBE', 'NSE', 'OES', 'PEM', 'PGE', 'PHS', 'PWM',
               'RSM', 'RTF', 'SCE', 'SCH', 'SCM', 'SMS', 'TWH', 'UES', 'VSE', 'WBE', 'WES', 'WFE', 'WHE', 'WHH', 'WHM', 'WHS', 'WMS')

def wait_For_Enter(message = "Press Enter to continue..."):
    print(message)
    keyboard.wait(hotkey='enter')
    pass

def check_School_Abbreviation(input):
    if input in SCHOOLS:
        return True
    else:
        return False
    
def gui_File_Path_Out(fileExtension):
    while True:
        Tk().withdraw()          
        fileName = filedialog.asksaveasfilename(defaultextension=fileExtension)
        if not fileName:
            print("Whoops! Looks like you didn't select a valid path. Please try again.")
        else:
            break
    return fileName

def gui_File_Path_In():
    while True:
        Tk().withdraw()
        fileName = filedialog.askopenfilename()
        if not fileName:
            print("Whoops! Looks like you didn't select a valid path. Please try again.")
        else:
            break
    return fileName

def cleanup_Files():
    def clear_Dir(dir):
        try:
            for file in os.listdir(dir):
                filePath = os.path.join(dir, file)
                if filePath:
                    os.remove(filePath)
        except OSError:
            print("Error occurred while deleting files.")
    directories = (TEMPDIR, DELLOUTPUTDIR)
    for i in directories:
        clear_Dir(i)