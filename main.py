import myExcel
import myWord
import myQRCode
import myMail
import myConfig
import myGUI
import myBusiness


import json
import time



import sys
from PySide2 import QtCore, QtWidgets
from PySide2.QtWidgets import QMainWindow, QWidget, QLabel, QLineEdit, QTextEdit
from PySide2.QtWidgets import QPushButton, QFileDialog
from PySide2.QtCore import QSize    


import myConfig
print("test")
from time import sleep
#sleep(15)

# from multiprocessing import Process
# import threading, 
import sys, os
import time

import pythoncom
from threading import Thread
import win32com.client as win32

def createExcel():
    pythoncom.CoInitialize()
    myBusiness.the_aim_of_the_program_with_delay()




if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWin = myGUI.MainWindow()
    mainWin.setFromConfigurationFile()
    # the aim of the program 
    mainWin.show()
    #myBusiness.the_aim_of_the_program_with_delay()


    thread = Thread(target = createExcel)
    thread.start()

    """
    thread = threading.Thread(target=myBusiness.the_aim_of_the_program_with_delay, args=())
    #thread.daemon = True                            # Daemonize thread
    thread.start()                                  # Start the execution
    #my_func()
    """
    
    #do stuff
    mainWin.setFromConfigurationFile()
    sys.exit( app.exec_() )

# C:/Python38/python.exe -m pip install auto-py-to-exe
# C:/Python38/python.exe -m auto-py-to-exe
# py -3.7 -m pyinstaller --noconfirm --onedir --console --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/main.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myBusiness.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myConfig.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myExcel.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myGUI.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myMail.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myQRCode.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myWord.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"
# pyinstaller --noconfirm --onedir --console --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/main.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myBusiness.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myConfig.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myExcel.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myGUI.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myMail.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myQRCode.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myWord.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"

# works
# pyinstaller --noconfirm --onedir --console --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"

# pyinstaller --noconfirm --onefile --console --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"

# pip freeze > requirements.txt
# ModuleNotFoundError: No module named 'Image'
#  pip install pillow


# pyinstaller --noconfirm --onedir --console --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"
# pyinstaller --onefile --onedir --console --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"

# pyinstaller.exe --onefile --console --upx-exclude vcruntime140.dll --upx-exclude ucrtbase.dll  --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"

"""
openpyxl

PySide2

multiprocessing

docx docxtpl mailmerge
qrcode pillow 

"""

# https://dev.to/eshleron/how-to-convert-py-to-exe-step-by-step-guide-3cfi

