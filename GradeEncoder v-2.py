print('**CKC GRADING SYSTEM AUTO ENCODER**\n\n##### Please wait while the application is still loading... \n')

# library imports
import os
import shutil
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox as msg
import sys
import threading
from tkinter import filedialog
import pyautogui as macro
import pandas
import time

#Select grading program main window
select_Grade = tk.Tk()
select_Grade.title("Select Grading Program")
select_Grade.geometry('300x300+540+250')


#Selection Functions
def college():
    global grading_program
    grading_program = "College"
    MainWindow()

def basicEd():
    global grading_program
    grading_program = "Basic Ed"
    MainWindow()

def About():
    about = tk.Toplevel()

    about.title('Auto Encoder - created by: Sir Clark (ABOUT)')
    about.geometry('300x200+500+190')
    about.resizable(False, False)

    about_label = tk.Label(about, text="About Sir Clark", font='sans 12 bold')
    about_label.pack(pady=20, padx=20, anchor='nw')
    about_label = tk.Label(about, text="I am a Math teacher and\na programming hobbyist.\n\nemail: jerryclarkc0420@gmail.com", font='sans 10')
    about_label.pack(pady=10, padx=10, anchor='nw')

def Help():
    help = tk.Toplevel()

    help.title('Auto Encoder - created by: Sir Clark (GUIDE)')
    help.geometry('620x450+10+10')
    # help.resizable(False, False)

    help_label = tk.Label(help, text="INSTRUCTIONS FOR USAGE", font='sans 12 bold')
    help_label.pack(pady=10)

    instructions = """** Before anything else, make sure that you are using the Excel format given 
    by Sir Clark. 

1.) Log in to the Grading Program you are going to encode and choose the class and subject you would like to input the raw score.

2.) Enter the TOTAL RAW SCORES.

3.) Run "Auto Encoder" that is located on your Desktop as Adminstrator (skip this part if the program is already running after the installation). 
#Help: right-click "Auto Encoder" app and click "Run as Administrator"

4.) Choose the Grading Program (Basic Ed or College).

5.) Click "OPEN EXCEL" button then locate and select Excel file of your grading 
sheet.

6.) Input the number of students in that class. Then, select the term in the dropdown
(e.g. Prelim, Midterm, Finals)

7.) Click "START AUTO ENCODING". Then, go back to the Basic Ed/College grading program and CLICK THE FIRST ENTRY of the "WW" or "QU" and wait for the auto encoder to start.

NOTE: Once the encoding process has started, DO NOT PRESS KEYS OR MOVE 
YOUR MOUSE. If the encoding messed up, move your mouse on the UPPER-RIGHT CORNER of the screen to abort the process, after that, close the program.

Enjoy being productive!!!!

-- Sir Clark"""

    v=tk.Scrollbar(help, orient='vertical')
    v.pack(side=tk.RIGHT, fill='y')

    help_text = tk.Text(help, font='Georgia 12', padx=5, pady=5, yscrollcommand=v.set)
    help_text.insert(tk.END, instructions)
    help_text.pack(pady=10, padx=10)
    v.config(command=help_text.yview)
    help_text.bind("<Key>", lambda e: "break")

def MainWindow():
    # --- main window---    
    select_Grade.withdraw()

    window = tk.Toplevel()

    window.title('Auto Encoder - created by: Sir Clark')
    window.geometry('443x350+600+200')
    window.resizable(False, False)

    if grading_program == "Basic Ed":
        window.configure(bg='#b8c7fa')
        window_bg = '#b8c7fa'
    elif grading_program == "College":
        window.configure(bg='#bdffb4')
        window_bg = '#bdffb4'

    #Menu
    menubar = tk.Menu(window, background='#ffec82', foreground='black', activebackground='#6a92f7', activeforeground='black')  
    file = tk.Menu(menubar, tearoff=0)  
    file.add_command(label="Go Back to Selection Window", command=lambda: [window.withdraw(), select_Grade.deiconify()])
    menubar.add_cascade(label="Select Grading Program", menu=file)  

    help = tk.Menu(menubar, tearoff=0)
    help.add_command(label="Guide in Using this Auto Encoder", command=Help)  
    menubar.add_cascade(label="Help", menu=help) 

    about = tk.Menu(menubar, tearoff=0)
    about.add_command(label="About the Creator", command=About)  
    menubar.add_cascade(label="About", menu=about) 

    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Main Program ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    class Redirect(): #nothing special here in this class, it makes the code work hehe

        def __init__(self, widget, autoscroll=True):
            self.widget = widget
            self.autoscroll = autoscroll

        def write(self, text):
            self.widget.insert('end', text)
            if self.autoscroll:
                self.widget.see("end")  # autoscroll

    #get the count
    inputLabel = tk.Label(window, text="How many students in this class?:", font='sans 8 bold')
    inputLabel.configure(bg=window_bg)
    inputLabel.place(x=15, y=120)
    inputText = tk.Text(window, height = 1, width = 4)
    inputText.place(x=210, y=120)

    #get excel data
    def openFile():
        window.filename = filedialog.askopenfilename(initialdir="C:/", title="Select Your Grades Excel File", filetypes=(("Excel File", "*.xlsx"),))
        global excel_file
        excel_file = window.filename
        excelLable = tk.Label(window, text=excel_file)
        excelLable.configure(bg='#f0ce73')
        excelLable.place(x=110, y=95)
    getExcel = tk.Button(window, text="OPEN EXCEL", command=openFile, fg='white', bg='#01700d', font=("Helvetica", 9))
    getExcel.place(x=17, y=91)

    #get term 
    current_var = tk.StringVar()
    combobox = ttk.Combobox(window, textvariable=current_var, width=13, state='readonly')
    combobox.place(x=320, y=120)
    combobox['values'] = ('Prelim-autoen', 'Midterm-autoen', 'Finals-autoen')

    inputLabel2 = tk.Label(window, text="Select Term:", font='sans 8 bold')
    inputLabel2.configure(bg=window_bg)
    inputLabel2.place(x=244, y=120)

    # ----------Mainfunctions-----------

    # this is function called when Start button is clicked
    def run():
        try:
            #copy the original excel to current dir and rename to temp_data.xlsx
            src = str(excel_file)
            des = str(os.path.abspath(os.path.basename(excel_file)))
            rnm = str(os.path.abspath("temp_data.xlsx"))

            try:
                shutil.copyfile(src, des)
                os.rename(des, rnm)
            except:
                os.remove("temp_data.xlsx")
                shutil.copyfile(src, des)
                os.rename(des, rnm)

            #start the encoding process
            threading.Thread(target=encode_function).start()
        except:
            msg.showerror('Error!', 'Something is not right!\n\nMake sure that you did the following before starting:\n\n-  Selected the Excel file.\n-  Typed the correct number of students.\n-  Selected the appropriate period/term.\n\n-- Sir Clark')    

    #Encoding Function
    def encode_function():

        print('Intitializing...')

        try:
            macro.FAILSAFE = True #put mouse pointer to upper-right corner of the screen to abort process

            current_term = combobox.get() #this will retrieve values from the combobox 
            data = pandas.read_excel('temp_data.xlsx', sheet_name=current_term)
            count = 0 
            student_count_input = inputText.get(1.0, "end-1c")
            student_count = int(student_count_input)  - 1

            

            if grading_program == "Basic Ed": #this is based on the selection made in the selection window line 22-31

                print('Encoding Program: Basic Ed\nClass size: ' + inputText.get(1.0, "end-1c") + '\nTerm: '+ combobox.get())
                time.sleep(1)
                print('\nStarting encoding in ...')
                time.sleep(0.5)
                print('############# 5 #############')
                time.sleep(1)
                print('############# 4 #############')
                time.sleep(1)
                print('############# 3 #############')
                time.sleep(1)
                print('############# 2 #############')
                time.sleep(1)
                print('############# 1 #############')
                time.sleep(2)
                print('Encoding the grades automatically ...')
                print('')
                print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
                
                for column in data['id'].tolist():
                    
                    data = data.fillna(0) # it will type 0 of the cell is empty

                    macro.typewrite(str(int(data['WW'][count])))
                    time.sleep(0.1)
                    macro.press('tab')
                    time.sleep(0.1)
                    macro.typewrite(str(int(data['PT'][count])))
                    time.sleep(0.1)
                    macro.press('tab')
                    time.sleep(0.1)
                    macro.typewrite(str(int(data['QA'][count])))
                    time.sleep(0.1)
                    #next entry
                    macro.press('tab', presses = 2)
                    time.sleep(0.1)
                    macro.press('enter')
                    time.sleep(0.1)
                    macro.press('tab', presses = 2)
                
                    if count == student_count:
                        break        
                    count += 1

                os.remove('temp_data.xlsx') #generated temp_data.xlsx will be removed once done encoding
                print('Finished encoding the grades!')
                prompt = msg.showinfo('From Sir Clark (Basic Ed Encoding)','Raw scores are encoded successfully.\n\nHave a nice productive day!\n\n\nClick "OK" to go back to the selection screen.')
                if prompt == "ok":
                    window.destroy()
                    select_Grade.deiconify()

            elif grading_program == "College":

                print('Encoding Program: Basic Ed\nClass size: ' + inputText.get(1.0, "end-1c") + '\nTerm: '+ combobox.get())
                time.sleep(1)
                print('Starting encoding in ...')
                time.sleep(0.5)
                print('############# 5 #############')
                time.sleep(1)
                print('############# 4 #############')
                time.sleep(1)
                print('############# 3 #############')
                time.sleep(1)
                print('############# 2 #############')
                time.sleep(1)
                print('############# 1 #############')
                time.sleep(2)
                print('Encoding the grades automatically ...')
                print('')
                print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')

                for column in data['id'].tolist():

                    data = data.fillna(0) # it will type 0 of the cell is empty
                    #Edit this for the college
                    macro.typewrite(str(int(data['QU'][count])))
                    time.sleep(0.1)
                    macro.press('tab')
                    time.sleep(0.1)
                    macro.typewrite(str(int(data['PER'][count])))
                    time.sleep(0.1)
                    macro.press('tab')
                    time.sleep(0.1)
                    macro.typewrite(str(int(data['OUT'][count])))
                    time.sleep(0.1)
                    macro.press('tab')
                    time.sleep(0.1)
                    macro.typewrite(str(int(data['TE'][count])))
                    time.sleep(0.1)
                                      
                    #next entry
                    macro.press('tab')
                    time.sleep(0.1)
                    macro.press('enter')
                    time.sleep(0.1)
                    macro.press('tab', presses = 3)
                                    
                    if count == student_count:
                        break        
                    count += 1

                os.remove('temp_data.xlsx') #generated temp_data.xlsx will be removed once done encoding

                prompt = msg.showinfo('From Sir Clark (College Encoding)','Raw scores are encoded successfully.\n\nHave a nice productive day!\n\n\nClick "OK" to go back to the selection screen.')
                if prompt == "ok":
                    window.destroy()
                    select_Grade.deiconify()

        except:
            msg.showerror('Error!', 'Something is not right!\n\nMake sure that you did the following before starting:\n\n-  Selected the Excel file.\n-  Typed the correct number of students.\n-  Selected the appropriate period/term.\n\n-- Sir Clark')    
            print("Theres is an error while initializing.")
    #  - Label - 
    head = tk.Label(window, text="CKC GRADING SYSTEM AUTO ENCODER " +"(" + grading_program + ")")
    head.config(font=('Bahnschrift Light', 12, 'bold'), bg=window_bg)
    head.pack(pady=2)

    Label = tk.Label(window, text="Note: Once Started, you have 5 seconds to click the first\ninput box in the CKC Grading System program.")
    Label.configure(bg=window_bg)
    Label.pack()

    Label2 = tk.Label(window, text="DO NOT PRESS ANY KEYS OR MOVE YOUR MOUSE WHILE ENCODING!")
    Label2.configure(bg=window_bg)
    Label2.pack()

    # - Log Textbox -
    text = tk.Text(window, width=50, height=10, bg='black', fg='#a8a9e4')
    # text.pack()
    text.place(x=20, y=145)
    text.bind("<Key>", lambda e: "break")

    old_stdout = sys.stdout    #This will print the console log
    sys.stdout = Redirect(text)

    # button
    btn = tk.Button(window, text="START AUTO ENCODING", command=lambda: [run(), ], fg='black', bg='#ffd900', font='sans 10 bold')
    btn.place(x=255, y=315)
    # btn.pack()

    def exit_MainWindow():
        window.destroy()
        select_Grade.deiconify()

    copyright = tk.Label(window, text="© Jerry Clark Ian Cabuntucan, 2022", font='sans 7 bold')
    copyright.configure(bg=window_bg)
    copyright.place(x=5, y=325)

    window.protocol("WM_DELETE_WINDOW", exit_MainWindow) #this will exit the program through this top level window
    window.config(menu=menubar)
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ End Main Program ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
def exit_program():
    select_Grade.quit()

#Select Window
logo = tk.PhotoImage(file = "ckc logo.png")
tk.Label(select_Grade, text = 'Click Me !', image = logo).pack(pady=10)

select_Label = tk.Label(select_Grade, text="Which Program are you encoding to?", font=("Helvetica", 12, "bold"))
select_Label.pack(pady='10')
select_btn_bed = tk.Button(select_Grade, text="Basic Ed", command=basicEd, fg='white', bg='#4463fc', font=("Helvetica", 11, "bold"))
select_btn_bed.pack(fill="x", padx=10)
select_btn_college = tk.Button(select_Grade, text="College", command=college, fg='white', bg='#2c9b1e', font=("Helvetica", 11, "bold"))
select_btn_college.pack(fill="x", padx=10, pady=10)
exit_btn = tk.Button(select_Grade, text="Exit", command=exit_program, fg='white', bg='#db3a3a', font=("Helvetica", 9, "bold"))
exit_btn.pack(anchor='e', padx=10, pady=5)

select_Grade.mainloop()