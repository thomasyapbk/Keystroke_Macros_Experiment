####Libraries ####
from sys import is_finalizing
from tkinter import Button
from typing import Text
import pyautogui as pag
import time
import pandas as pd
import pathlib
import os
import openpyxl
from datetime import datetime

datapath = "D:/04_MLC/Keystroke_Macros_Experiment" + "/" + "sample_data_keystrokes.xlsx"

#Load Data from excel
df = pd.read_excel(datapath, engine='openpyxl')

#df.loc[1,'ASIC Report'] ## Returns "Description of report"
#df.loc[1,'Event 1001: Test legal checklist'] ## Returns "Commenced investigation"

#date1 = df.loc[3,'Event 1001: Test legal checklist'].strftime('%d/%m/%Y') 




########### Start with PyAutoGUI
section_lag = int(pag.prompt('Please select the time lag per section. Default is 5.','Select time'))
key_lag = float(pag.prompt('Please select the time lag per keystroke. Default is 0.5.','Select time'))
pag.alert(text = 'After clicking the Ok button, move mouse over the webpage on the windows taskbar.')
def have_input(keystroke):
        for i in range(0,len(keystroke)):
                if keystroke[i] == "x":
                        return True
                else:
                        return False

def keystroke_mapping(keystroke, proposed_input):
        if have_input:
                for i in range(0,len(keystroke)):
                        keystroke_mapping_input(keystroke[i],proposed_input)
                        time.sleep(0.5)
                
        else:
                for j in range(0,len(keystroke)):
                        keystroke_mapping_no_input(keystroke[j])
                        time.sleep(0.5)
        

def keystroke_mapping_no_input(key):
        if key == "-":
                pag.press('tab')
        elif key == ".":
                pag.press('space')
        elif key == "u":
                pag.press('up')
        elif key == "d":
                pag.press('down')

def keystroke_mapping_input(key,data):
        if key == "-":
                pag.press('tab')
        elif key == ".":
                pag.press('space')
        elif key == "u":
                pag.press('up')
        elif key == "d":
                pag.press('down')
        elif key == "x":
                pag.write(str(data))



screenWidth, screenHeight = pag.size()
#Find Location of webpage
time.sleep(2)
############INSTRUCTION################                                Move mouse to position of web browser with ASIC Breach Reporting
pag.click(pag.position())
pag.click(screenHeight/2,screenWidth/2)
######## Start reading keystrokes and inputs
#for i in range(0,len(df)):
for i in range(0,len(df)):
        if i+1 < len(df):
                next_section = df.loc[i+1,"Section"]
        proposed_response = df.loc[i,"Proposed Response"]
        if isinstance(proposed_response,datetime):
            proposed_response = str(proposed_response.strftime('%d/%m/%Y'))
        keystroke = df.loc[i,"Keystrokes"]
        keystroke_mapping(keystroke, str(proposed_response))
        if not(pd.isna(next_section)):
                time.sleep(5)
pag.alert(text = 'The Script is finished. Please check the Review Page for any errors')

#
#for i in range(0,10):
#        proposed_response = df.loc[i,"Proposed Response"]
#        keystroke = df.loc[i,"Keystrokes"]
#        print( "The proposed response is : ", proposed_response)
#        print( "The Keystroke associated with the response is : ", keystroke)



#Add in pop up to say do not touch the keyboard and input for wait timing.
#I need an override
#Figure out a way to populate the next sheet