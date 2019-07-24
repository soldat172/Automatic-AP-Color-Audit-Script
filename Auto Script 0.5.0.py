import openpyxl, pyautogui, tkinter, time
from tkinter import filedialog, messagebox, Tk
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
    for x in range (0,5): #range of APS being colored
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


def search_options ():  #sets visio to search all pages for AP IDs, must be done before coloring
    pyautogui.hotkey('ctrl', 'f')
    pyautogui.press(['tab','tab'])
    pyautogui.press('down')
    pyautogui.press('esc')


def auto_script_green():  #finds AP name in Visio and colors it GREEN,
    AP_green_count = 0
    for x in AP_green_list: 
        AP_num = AP_green_list[AP_green_count]
        AP_num_print = (AP_num[8:13]) if len(AP_num) > 5 else AP_num #Truncates site name, IDF, and AP letter.
        print(AP_num_print)
        pyautogui.hotkey('ctrl', 'f')
        pyautogui.typewrite(AP_num_print)
        pyautogui.press(['enter', 'esc','esc','tab','enter'])
        pyautogui.hotkey('shift','tab','enter')
        pyautogui.hotkey('shift','tab','enter')
        pyautogui.hotkey('alt', 'h')#line coloring
        pyautogui.press(['l','down','down','down','down','down','down','down','left','enter'])
        pyautogui.hotkey('alt', 'h')#fill coloring
        pyautogui.press(['i','down','down','down','down','down','down','down','left','enter'])
        AP_green_count += 1     #adds one to total green count


def auto_script_red():    #finds AP name in Visio and colors it RED, 
    AP_red_count = 0
    for x in AP_red_list: #finds AP name in Visio and colors it
        AP_num = AP_red_list[AP_red_count]
        AP_num_print = (AP_num[8:13]) if len(AP_num) > 5 else AP_num #Truncates site name, IDF, and AP letter.
        print(AP_num_print)
        pyautogui.hotkey('ctrl', 'f')
        pyautogui.typewrite(AP_num_print)
        pyautogui.press(['enter', 'esc','esc','tab','enter'])
        pyautogui.hotkey('shift','tab','enter')
        pyautogui.hotkey('shift','tab','enter')
        pyautogui.hotkey('alt', 'h')#line coloring
        pyautogui.press(['l','down','down','down','down','down','down','down','left','left','left','left','enter'])
        pyautogui.hotkey('alt', 'h')#fill coloring
        pyautogui.press(['i','down','down','down','down','down','down','down','left','left','left','left','enter'])
        AP_red_count += 1   #adds one to total red count


def auto_script_orange(): #finds AP name in Visio and colors it ORANGE,
    AP_orange_count = 0
    for x in AP_orange_list: #finds AP name in Visio and colors it
        AP_num = AP_orange_list[AP_orange_count]
        AP_num_print = (AP_num[8:13]) if len(AP_num) > 5 else AP_num #Truncates site name, IDF, and AP letter.
        print(AP_num_print)
        pyautogui.hotkey('ctrl', 'f')
        pyautogui.typewrite(AP_num_print)
        pyautogui.press(['enter', 'esc','esc','tab','enter'])
        pyautogui.hotkey('shift','tab','enter')
        pyautogui.hotkey('shift','tab','enter')
        pyautogui.hotkey('alt', 'h')#line coloring
        pyautogui.press(['l','down','down','down','down','down','down','down','left','left','left','enter'])
        pyautogui.hotkey('alt', 'h')#fill coloring
        pyautogui.press(['i','down','down','down','down','down','down','down','left','left','left','enter'])
        AP_orange_count += 1   # adds one to total orange count


def auto_script_blue():   #finds AP name in Visio and colors it BLUE,
    AP_blue_count = 0
    for x in AP_blue_list: 
        AP_num = AP_blue_list[AP_blue_count]
        AP_num_print = (AP_num[8:13]) if len(AP_num) > 5 else AP_num #Truncates site name, IDF, and AP letter.
        print(AP_num_print)
        pyautogui.hotkey('ctrl', 'f')
        pyautogui.typewrite(AP_num_print)
        pyautogui.press(['enter', 'esc','esc','tab','enter'])
        pyautogui.hotkey('shift','tab','enter')
        pyautogui.hotkey('shift','tab','enter')
        pyautogui.hotkey('alt', 'h')#line coloring
        pyautogui.press(['l','down','down','down','down','down','down','down','right','enter'])
        pyautogui.hotkey('alt', 'h')#fill coloring
        pyautogui.press(['i','down','down','down','down','down','down','down','right','enter'])
        AP_blue_count += 1      #adds one to total blue count


class majority(tkinter.Frame):  #POP-UP GUI for choosing majority of Sheets color
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):   #buttons in the GUI
        #self.title.pack(text="""Choose AP Color Majority:""", justify = Tk.LEFT, padx = 20)

        self.option1 = tkinter.Button(self, fg = "green")    #Button, colored green
        self.option1["text"] = "Green"                      #button, named green
        self.option1["command"] = self.output1              #when button pressed, execute output1
        self.option1.pack(side="top")                       #position selt at top most position

        self.option2 = tkinter.Button(self, text = "Red", fg = "red", command = self.output2)  #better formatted buttons
        self.option2.pack(side="top")

        self.option3 = tkinter.Button(self, text = "Orange", fg = "orange", command = self.output3)
        self.option3.pack(side="top")

        self.option4 = tkinter.Button(self, text = "Blue", fg = "blue", command = self.output4)
        self.option4.pack(side="top")

        self.quit = tkinter.Button(self, text="QUIT", fg="red", command=self.master.destroy)
        self.quit.pack(side="bottom", pady = 30)

    def output1(self):
        print("GREEN AP Majority...")
        print("QUICK! You have 5 seconds to click into your visio file!!!\n")
        time.sleep(5)
        auto_script_red()
        auto_script_orange()
        auto_script_blue()
        exit(0)

    def output2(self):
        print("RED AP Majority...")
        print("QUICK! You have 5 seconds to click into your visio file!!!\n")
        time.sleep(5)
        auto_script_green()
        auto_script_blue()
        auto_script_orange()
        exit(0)

    def output3(self):
        print("Orange AP Majority...")
        print("QUICK! You have 5 seconds to click into your visio file!!!\n")
        time.sleep(5)
        auto_script_green()
        auto_script_red()
        auto_script_blue()
        exit(0)

    def output4(self):
        print("Blue AP Majority...")
        print("QUICK! You have 5 seconds to click into your visio file!!!\n")
        time.sleep(5)
        auto_script_green()
        auto_script_red()
        auto_script_orange()
        save_new_sheet()
        exit(0)

def save_new_sheet():   #Pop-up message box, asks if user wants to save sorted date to new excell file/sheet
    result = messagebox.askyesno("Visio AP Coloring Tool","Do you want to save data in new Excell?")
    print(result)
    if result == True:
        print("data saved in 'New File'")

    else:
        pass

open_audit_sheet()
audit_sheet_sorter()
search_options()

root = tkinter.Tk()
app = majority(master=root)
app.master.title("Auto AP Colorer!!!")
app.master.minsize(100,150)
app.mainloop()

