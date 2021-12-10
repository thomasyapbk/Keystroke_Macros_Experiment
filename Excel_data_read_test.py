####Libraries ####
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

print("The size of the dataframe is: " , df.size)
print("The number of rows in the dataframe is: ", len(df))
screenWidth, screenHeight = pag.size()
print (screenHeight/2,screenWidth/2)
""" ######## Start reading keystrokes and inputs
for i in range(0,0):
        proposed_response = df.loc[i,"Proposed Response"]
        if isinstance(proposed_response,datetime):
            proposed_response = proposed_response.strftime('%d/%m/%Y') 
        keystroke = df.loc[i,"Keystrokes"]
        print( "The proposed response is : ", proposed_response)
        print( "The Keystroke associated with the response is : ", keystroke)
        for j in range(0,len(keystroke)):
            print("The individual inputs are:  ", keystroke[j])
            """

