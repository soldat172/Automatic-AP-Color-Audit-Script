import openpyxl, pyautogui
from tkinter import filedialog, Tk
from openpyxl import Workbook, load_workbook

redAPs=[]
greenAPs=[]
orangeAPs=[]
blueAPs=[]

count1 = 0

greenAPsCount = int('0')
redAPsCount = int('0')
orangeAPsCount = int('0')
blueAPsCount = int('0')

excelNameColumnLetter = 'A'
excelFailColumnLetter = 'K'
excelNameRow = int('3')
excelFailRow = int('3')
excelNameColumnNumber = str(excelNameColumnLetter) + str(excelNameRow)
excelFailColumnNumber = str(excelFailColumnLetter) + str(excelFailRow)

Tk().withdraw() # We don't want a full GUI, so keep the root window from appearing
excelName = filedialog.askopenfilename() # Show an "Open" dialog box and return the path to the selected file
book = load_workbook(excelName)#Opens Excel
sheet = book.active #Makes the current sheet active

for x in range (0,100):
    excelFailColumnNumber = str(excelFailColumnLetter) + str(excelFailRow)
    excelNameColumnNumber = str(excelNameColumnLetter) + str(excelNameRow)
    specificCellValue = sheet[excelFailColumnNumber]
    specificCellValue2 = sheet[excelNameColumnNumber]
    apFail = specificCellValue.value
    apName = specificCellValue2.value
    
    if apFail == 1:
        print ('green')
        greenAPs.append(apName)
        print(greenAPs[greenAPsCount])
        greenAPsCount += 1

    elif apFail == 2:
        print ('red')
        redAPs.append(apName)
        print(redAPs[redAPsCount])
        redAPsCount += 1
        
    elif apFail == 3:
        print ('orange')
        orangeAPs.append(apName)
        print(orangeAPs[redAPsCount])
        orangeAPsCount += 1
        
    elif apFail == 4:
        print ('blue')
        blieAPs.append(apName)
        print(blueAPs[redAPsCount])
        blueAPsCount += 1

    else:
        print(apFail)
        print('Outside 1-4 value')
            
    excelFailRow += 1
    excelNameRow += 1
    
pyautogui.hotkey('ctrl', 'f')
pyautogui.press(['tab','tab'],interval=.05)
pyautogui.press('down')
pyautogui.press('esc')
    
for x in greenAPs:
    temporaryValue = greenAPs[count1]
    parentSiteName,idfNumber,apNumberNaked = temporaryValue.split('-') # Splits the ap name into 3 parts
    print(apNumberNaked)
    pyautogui.hotkey('ctrl', 'f')
    pyautogui.typewrite(apNumberNaked)
    pyautogui.press('enter')
    pyautogui.press('esc')
    pyautogui.hotkey('alt', 'h')
    pyautogui.press('l')
    pyautogui.press(['down','down','down','down','down','down','down'])
    pyautogui.press('left')
    pyautogui.press('enter')
    pyautogui.hotkey('alt', 'h')
    pyautogui.press('i')
    pyautogui.press(['down','down','down','down','down','down','down'])
    pyautogui.press('left')
    pyautogui.press('enter')
    pyautogui.press('esc')
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.hotkey('alt', 'h')
    pyautogui.press('l')
    pyautogui.press(['down','down','down','down','down','down','down'])
    pyautogui.press('left')
    pyautogui.press('enter')
    pyautogui.hotkey('shift', 'tab')
    pyautogui.hotkey('shift', 'tab')
    pyautogui.press('enter')
    pyautogui.hotkey('alt', 'h')
    pyautogui.press('l')
    pyautogui.press(['down','down','down','down','down','down','down'])
    pyautogui.press('left')
    pyautogui.press('enter')
    count1 += 1


