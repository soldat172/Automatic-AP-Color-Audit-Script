import openpyxl, pyautogui
from tkinter import filedialog, Tk
from openpyxl import Workbook, load_workbook

#Controls which column and row the script works from
excelNameColumnLetter = 'A' #column letter to read AP names from
excelNameRow = int('3')#Starting row in excel

excelFailColumnLetter = 'K' #column letter to read fails from
excelFailRow = int('3')#Starting row in excel

excelNameColumnNumber = str(excelNameColumnLetter) + str(excelNameRow)#combines column letter and row together
excelFailColumnNumber = str(excelFailColumnLetter) + str(excelFailRow)#combines column letter and row together

#Initiate lists and their variables
greenAPs=[]#initiates a list for green AP's ('1')
greenAPsListCount = int('0')#Keeps track how long the list is
greenAPsCount = 0 #Counts up till it hits the end of the list, stops Auto script from erroring out

redAPs=[]#initiates a list for red AP's ('2')
redAPsListCount = int('0')#Keeps track how long the list is
redAPsCount = 0 #Counts up till it hits the end of the list, stops Auto script from erroring out

orangeAPs=[]#initiates a list for orange AP's ('3')
orangeAPsListCount = int('0')#Keeps track how long the list is
orangeAPsCount = 0 #Counts up till it hits the end of the list, stops Auto script from erroring out

blueAPs=[]#initiates a list for blue AP's ('4')
blueAPsListCount = int('0')#Keeps track how long the list is
blueAPsCount = 0 #Counts up till it hits the end of the list, stops Auto script from erroring out

#GUI control
Tk().withdraw() # We don't want a full GUI, so keep the root window from appearing
excelName = filedialog.askopenfilename() # Show an "Open" dialog box and return the path to the selected file
book = load_workbook(excelName)#Opens Excel
sheet = book['Audit'] #Makes the Audit t

print('Start of Auto Script')
for x in range (0,100):
    excelFailColumnNumber = str(excelFailColumnLetter) + str(excelFailRow)
    excelNameColumnNumber = str(excelNameColumnLetter) + str(excelNameRow)
    specificCellValue = sheet[excelFailColumnNumber]
    specificCellValue2 = sheet[excelNameColumnNumber]
    apFail = specificCellValue.value
    apName = specificCellValue2.value
    
    if apFail == 1:
        greenAPs.append(apName)
        print(greenAPs[greenAPsListCount])
        greenAPsListCount += 1

    elif apFail == 2:
        redAPs.append(apName)
        print(redAPs[redAPsListCount])
        redAPsListCount += 1
        
    elif apFail == 3:
        orangeAPs.append(apName)
        print(orangeAPs[orangeAPsListCount])
        orangeAPsListCount += 1
        
    elif apFail == 4:
        blueAPs.append(apName)
        print(blueAPs[BlueAPsListCount])
        blueAPsListCount += 1

    else:
        print('End of Audit Sheets')
        break
            
    excelFailRow += 1
    excelNameRow += 1
    
pyautogui.hotkey('ctrl', 'f')
pyautogui.press(['tab','tab'])
pyautogui.press('down')
pyautogui.press('esc')


for x in greenAPs:
    temporaryValue = greenAPs[greenAPsCount]
    parentSiteName,idfNumber,apNumberNaked = temporaryValue.split('-') # Splits the ap name into 3 parts
    print(apNumberNaked)
    pyautogui.hotkey('ctrl', 'f')
    pyautogui.typewrite(apNumberNaked)
    pyautogui.press('enter')
    pyautogui.press(['esc','esc'])
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.hotkey('shift','tab','enter')
    pyautogui.hotkey('shift','tab','enter')
    pyautogui.hotkey('alt', 'h')#line coloring
    pyautogui.press('l')
    pyautogui.press(['down','down','down','down','down','down','down'])
    pyautogui.press('left')
    pyautogui.press('enter')
    pyautogui.hotkey('alt', 'h')#fill coloring
    pyautogui.press('i')
    pyautogui.press(['down','down','down','down','down','down','down'])
    pyautogui.press('left')
    pyautogui.press('enter')
    greenAPsCount += 1


