import openpyxl, pyautogui
from tkinter import filedialog, Tk
from openpyxl import Workbook, load_workbook

def open_audit_sheet (): #Opens a file explorer and returns path from chosen file
    global audit_sheet
    Tk().withdraw() # We don't want a full GUI, so keep the root window from appearing
    xlname = filedialog.askopenfilename() # Show an "Open" dialog box and return the path to the selected file
    audit_book = load_workbook(xlname)#Opens Excel
    audit_sheet = audit_book['Audit'] #grabs information from the "audit" sheet

def audit_sheet_sorter (): #reads all AP's from chosen Audit Sheet and categories them by their color
    global AP_red_list
    global AP_blue_list
    global AP_green_list
    global AP_orange_list
    global AP_red_count_list
    global AP_blue_count_list
    global AP_green_count_list
    global AP_orange_count_list
    
    xlname_col = 'A' #column letter to read AP names from
    xlname_col_fail = 'K' #column letter to read fails from
    xlname_row = int('3')#Starting row in excel
    xlname_row_fail = int('3')#Starting row in excel
    
    AP_red_list = []  #initiates a list for green AP's ('2')
    AP_blue_list = []  #initiates a list for green AP's ('4')
    AP_green_list = []  #initiates a list for green AP's ('1')
    AP_orange_list = []  #initiates a list for green AP's ('3')
    
    AP_red_count_list = int('0')#Keeps track how long the list is
    AP_blue_count_list = int('0')
    AP_green_count_list = int('0')
    AP_orange_count_list = int('0')
    
    print('Start of Auto Script')
    for x in range (0,15):
        xlname_col_fail_num = str(xlname_col_fail) + str(xlname_row_fail)
        xlname_col_num = str(xlname_col) + str(xlname_row)
        cell_value1 = audit_sheet[xlname_col_fail_num]
        cell_value2 = audit_sheet[xlname_col_num]
        AP_fail = cell_value1.value
        AP_ID = cell_value2.value
        
        if AP_fail == 1: #reads all green AP's
            AP_green_list.append(AP_ID)
            print(AP_green_list[AP_green_count_list])
            AP_green_count_list += 1

        elif AP_fail == 2: #reads all red AP's
            AP_red_list.append(AP_ID)
            print(AP_red_list[AP_red_count_list])
            AP_red_count_list += 1
            
        elif AP_fail == 3: #reads all orange AP's
            AP_orange_list.append(AP_ID)
            print(AP_orange_list[AP_orange_count_list])
            AP_orange_count_list += 1
            
        elif AP_fail == 4: #reads all blue AP's
            AP_blue_list.append(AP_ID)
            print(AP_blue_list[AP_blue_count_list])
            AP_blue_count_list += 1

        else:
            print('End of Audit Sheets')
            break   
        xlname_row_fail += 1
        xlname_row += 1

def search_options ():
    pyautogui.hotkey('ctrl', 'f')
    pyautogui.press(['tab','tab'])
    pyautogui.press('down')
    pyautogui.press('esc')

def auto_script():
    AP_green_count = 0
    for x in AP_green_list: #finds AP name in Visio and colors it
        AP_num = AP_green_list[AP_green_count] 
        AP_num_print = (AP_num[8:13]) if len(AP_num) > 5 else AP_num #Truncates site name, IDF, and AP letter.
        print(AP_num_print)
        pyautogui.hotkey('ctrl', 'f')
        pyautogui.typewrite(AP_num_print)
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
        AP_green_count += 1

open_audit_sheet ()
audit_sheet_sorter ()
search_options()
auto_script()
