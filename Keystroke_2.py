####Libraries ####
from typing import Text
import pyautogui as pag
import time
import pandas as pd
import pathlib
import os
import openpyxl
from datetime import datetime

datapath = "D:/04_MLC/Keystroke_Macros_Experiment" + "/" + "sample_data.xlsx"


""" Information on PyAutoGUI (https://pyautogui.readthedocs.io/en/latest/)
screenWidth, screenHeight = pyautogui.size() # Get the size of the primary monitor.
screenWidth, screenHeight 


currentMouseX, currentMouseY = pyautogui.position() # Get the XY position of the mouse.
currentMouseX, currentMouseY


pyautogui.moveTo(100, 150) # Move the mouse to XY coordinates.

pyautogui.click()          # Click the mouse.
pyautogui.click(100, 200)  # Move the mouse to XY coordinates and click it.
pyautogui.click('button.png') # Find where button.png appears on the screen and click it.

pyautogui.move(400, 0)      # Move the mouse 400 pixels to the right of its current position.
pyautogui.doubleClick()     # Double click the mouse.
pyautogui.moveTo(500, 500, duration=2, tween=pyautogui.easeInOutQuad)  # Use tweening/easing function to move mouse over 2 seconds.

pyautogui.write('Hello world!', interval=0.25)  # type with quarter-second pause in between each key
pyautogui.press('esc')     # Press the Esc key. All key names are in pyautogui.KEY_NAMES

with pyautogui.hold('shift'):  # Press the Shift key down and hold it.
        pyautogui.press(['left', 'left', 'left', 'left'])  # Press the left arrow key 4 times.
# Shift key is released automatically.

pyautogui.hotkey('ctrl', 'c') # Press the Ctrl-C hotkey combination.

pyautogui.alert('This is the message to display.') # Make an alert box appear and pause the program until OK is clicked.
"""
#Load Data from excel
df = pd.read_excel(datapath, engine='openpyxl')

df.loc[1,'ASIC Report'] ## Returns "Description of report"
df.loc[1,'Event 1001: Test legal checklist'] ## Returns "Commenced investigation"

date1 = df.loc[3,'Event 1001: Test legal checklist'].strftime('%d/%m/%Y') 

########### Start with PyAutoGUI

def press_rep(input,times):
        for i in range(0,times):
                pag.press(input)


screenWidth, screenHeight = pag.size()
#Find Location of webpage
time.sleep(2)
############INSTRUCTION################                                Move mouse to position of web browser with ASIC Breach Reporting
pag.click(pag.position())
pag.click(screenHeight/2,screenWidth/2)
#Page 1
press_rep('tab',8)
pag.press('space')
time.sleep(5)

#Page 2
press_rep('tab',8)
pag.press('space')
time.sleep(5)

#Page 3
press_rep('tab',3)
pag.write(date1)
##Enter Date