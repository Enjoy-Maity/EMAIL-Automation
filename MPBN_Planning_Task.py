import traceback                                # Importing traceback for the traceback of the exception
import tkinter as tk                            # Importing the Tkinter Module for Developing the GUI with alias.
from tkinter import *                           # Importing all the modules and methods available in Tkinter.
from tkinter import filedialog, messagebox      # Importing Filedialog and messagebox modules from tkinter module to browse files and show messages.
import tkinter.ttk as ttk                       # Importing ttk for tkinter styles.
from PIL import ImageTk, Image                  # Importing ImageTk, Image modules from Pillow(PIL) Module to handle GIF and Images.
import sys                                      # Importing the sys module to create method for exitting the application.
import subprocess                               # Importing subprocess module to run cmd commands.
import Planning_sheet_creater                   # Importing Planning Sheet Creater Module.
import circle_Email_Automation_Task             # Importing the Circle Email Automation Task Module.
import pandas as pd                             # Importing the pandas with pd alias to work with excel files.
import Email_package_generator                  # Importing Email_package_generator for creation of Email-Package in the workbook.
import interdomain_KPIs_Data_Prep_Task          # Importing Interdomain KPIs Data Prep Task Module.
import interdomain_KPIs_Mail_Comm_Task          # Importing the interdomain kpi mail communication Module.
import evening_mail_task                        # Importing evening mail task module.
import circle_reply_task                        # Importing circle reply task module.
from datetime import datetime, timedelta        # Importing datetime and timedelta module to get today's maintenance date

# Creating EmptyString Exception Class inheriting the Default Exception for raising and handling custom made empty string exception.
class EmptyString(Exception):
    def __init__(self, msg):
        self.msg = msg
        super().__init__(self.msg)
        messagebox.showerror(" Empty String Not Allowed", self.msg)

# Creating RegionHandlerException Exception Class inheriting the Default Exception for raising and handling custom made region handler exception exception.
class RegionHandlerException(Exception):
    def __init__(self, msg):
        self.msg = msg
        super().__init__(self.msg)
        messagebox.showerror(" Exception Occurred", self.msg)

# Creating EveningTaskException Exception Class inheriting the Default Exception for raising and handling custom made Evening Task exception.
class EveningTaskException(Exception):
    def __init__(self, msg):
        self.msg = msg
        super().__init__(self.msg)
        messagebox.showerror("  Exception Occurred", self.msg)

# Creating ContainsInteger Exception Class inheriting the Default Exception for raising and handling custom made exception for input fields with integer values.
class ContainsInteger(Exception):
    def __init__(self, msg):
        self.msg = msg
        messagebox.showerror("  Integer Not Allowed", self.msg)

# Creating FileNotSelected Exception Class inheriting the Default Exception for raising and handling custom made exception for file not selected.
class FileNotSelected(Exception):
    def __init__(self, msg, title):
        self.msg = msg
        self.title = title
        super().__init__(self.msg, self.title)
        messagebox.showerror(self.title, self.msg)

# Creating CustomException Exception Class inheriting the Default Exception for raising and handling custom made exceptions.
class CustomException(Exception):
    def __init__(self, msg, title):
        self.msg = msg
        self.title = title
        super().__init__(self.msg, self.title)
        messagebox.showerror(self.title, self.msg)

# Creating CustomWarning Exception Class inheriting the Default Exception for raising and handling custom made exception for handling custom warning.
class CustomWarning(Exception):
    def __init__(self, msg, title):
        self.msg = msg
        self.title = title
        super().__init__(self.msg, self.title)
        messagebox.showwarning(self.title, self.msg)

# Creating the Class for our GUI Application.
class App(tk.Tk):
    # Constructor Method(Function) for the Main GUI Window.
    def __init__(self, main_win):
        self.empty_string_list = []                                         # Declaration & Initialization of list for empty user input fields used in the Application to raise exception.
        self.integer_string_list = []                                       # Declaration & Initialization of list for user input fields with integer used in the Application to raise exception.
        self.main_win = main_win                                            # Setting the reference of object main win value to the main win value passed, in our case root GUI Window .
        self.main_win_flag = -1                                             # Setting the Main win Flag to be negative for detecting whether any Child GUI has been called while hiding the main GUI window.
        self.main_win.geometry("1080x701")                                  # Setting the default dimensions of the Main GUI Window.
        self.main_win.maxsize(1080, 701)                                    # Setting the maximum dimensions of the Main GUI Window.
        self.main_win.minsize(1080, 701)                                    # Setting the minimum dimensions of the Main GUI Window.
        self.main_win.iconbitmap("./images/ericsson-blue-icon-logo.ico")    # Setting the Icon to be shown on the title bar.
        self.main_win.title("   MPBN Planning Task")                        # Setting the Title of the Main GUI Window.
        
        '''
            Checking if the main_win_flag value is 0 or not. If the value is non-zero then the GIF Frame to shown is the first frame of the GIF,
            else any Child GUI was called at any point of time while running the application, and the Main GUI Window was Minimized or hidden,
            so the last frame of the GIF which was shown is fetched through the self.frame_idx variable and is sent to the 'update' method as argument
            so that the GIF with appropriate Frame number is fetched and shown.
        '''
        if self.main_win_flag == 0:         
            self.update(self.frame_idx.get())
        else:
            self.style = ttk.Style()

            # Setting the theme style settings to be used in the GUI.
            self.style.theme_use("vista")
            self.style.theme_settings("vista", {
                "TButton": {
                    "configure": {"padding": 2,
                                  "font": "Ericsson_Hilda 10 bold"},

                },
                "TMenubutton": {
                    "configure": {"font": "Ericsson_Hilda 14"},
                }
            }
            )

            self.main_win.bind("<Return>", self.file_browser_func)      # Binding the Enter key to call the file browser function to browse the MPBN Planning Task workbook.
            self.main_win.bind("<Escape>", self.main_win_quit)          # Binding the Escape key to quit the application.
            self.main_win.bind("<Alt-F4>",self.main_win_quit)           # Binding the Alt+F4 key to quit the application.
            self.task_running = 0
            self.task_module_running = ""

            # Fetching the Background Image of the Application to self.main_win_background_img variable.
            self.main_win_background_img = ImageTk.PhotoImage(
                Image.open("./images/MPBN PLANNING TASK_3_1.png"))
            
            '''
                Creating a canvas to hold the dynamic and static GUI components(GIF Frame and the Background Image) and positioning it over
                the main GUI Window.
            '''
            self.main_win_canvas = Canvas(
                self.main_win, width=1082, height=701, bd=0, highlightthickness=0, relief="ridge")
            self.main_win_canvas.grid(row=0, column=0, sticky=NW)

            # Getting the list of all the image frames of the GIF file.
            self.frames = [PhotoImage(file="./images/AI-transparent-automation-III.gif",
                                      format="gif -index %i" % (i)) for i in range(31)]

            # Setting the background image of the Canvas.
            self.main_win_canvas.create_image(
                0, 0, image=self.main_win_background_img, anchor="nw")
            
            # Declaring and Initializing Variable for file browser path.
            self.file_browser_file = ""
            
            # Creating the button for browsing the file by calling the Method(Function) for browsing the MPBN Planning Task Workbook.
            self.file_browser_btn = ttk.Button(
                self.main_win, text="Browse", command=lambda: self.file_browser_func(1))

            # Creating the Entry for the File Browser path to be selected.
            self.file_browser_entry = ttk.Entry(
                self.main_win, width=40, font=("Ericsson Hilda", 13))

            # Declaring and Initializing Variable for ITSM file browser path. 
            self.itsm_file_browser_file = ""
            
            # Creating the button for browsing the file by calling the Method(Function) for browsing the ITSM Raw Report.
            self.itsm_file_browser_btn = ttk.Button(
                self.main_win, text="Browse", command=lambda: self.itsm_file_browser_func(1))
            
            # Creating the Entry for the File Browser path to be selected.
            self.itsm_file_browser_entry = ttk.Entry(
                self.main_win, width=40, font=("Ericsson Hilda", 13))

            # Creating button for the planning sheet creation task
            self.planning_sheet_creater_task_btn = ttk.Button(
                self.main_win, text = "         Planning Sheet Preparation       ", command=self.planning_sheet_creater_task_func
            )
            
            # Creating button for the Circle Email- Automation task
            self.circle_email_automation_task_btn = ttk.Button(
                self.main_win, text = "         Circle Mail Communication         ", command=self.circle_email_automation_task_func_surity)

            # Creating button for the Email Package Data Preparation task
            self.email_package_prep_btn = ttk.Button(
                self.main_win, text="         Email Package Preparation      ", command=self.email_package_prep_func)

            # Creating button for the Interdoman KPIs Data Preparation task
            self.interdomain_kpis_data_prep_btn = ttk.Button(
                self.main_win, text="   Interdomain KPIs Data Preparation   ", command=self.interdomain_kpis_data_prep_func)

            # Creating button for the Interdomain KPIs Mail Communication task
            self.interdomain_kpis_mail_communication_btn = ttk.Button(
                self.main_win, text="Interdomain KPIs Mail Communication", command=lambda: self.interdomain_kpis_mail_communication_func(1))

            # Creating button for the Evening message task
            self.evening_task_btn = ttk.Button(
                self.main_win, text=" Email Package & Evening Message ", command=lambda: self.evening_task_func(1))

            # Creating button for the Executor Mail communication task
            self.executor_mail_communication_btn = ttk.Button(
                self.main_win, text="        Executor Mail Communication      ", command= self.executor_mail_communication)

            ##################################################### Status setter variables #####################################################################################################################
            self.planning_sheet_creater_task_status = StringVar(self.main_win_canvas)
            self.planning_sheet_creater_task_status.set("")

            self.circle_email_automation_task_status = StringVar(
                self.main_win_canvas)
            self.circle_email_automation_task_status.set("")

            self.email_package_prep_task_status = StringVar(
                self.main_win_canvas)
            self.email_package_prep_task_status.set("")

            self.interdomain_kpis_data_prep_task_status = StringVar(
                self.main_win_canvas)
            self.interdomain_kpis_data_prep_task_status.set("")
            self.interdomain_kpis_data_prep_task_completed = 0

            self.interdomain_kpis_mail_communication_status = StringVar(
                self.main_win_canvas)
            self.interdomain_kpis_mail_communication_status.set("")

            self.evening_task_status = StringVar(self.main_win_canvas)
            self.evening_task_status.set("")

            self.executor_mail_communication_status = StringVar(
                self.main_win_canvas)
            self.executor_mail_communication_status.set("")

            # List of colors used in labels, #FF00000 stands for 'Red', #00FF00 stands for 'Green', #FFFFFF stands for 'Whites'
            self.color = ["#FF0000", "#00FF00", "#FFFFFF"]
            
            self.planning_sheet_creater_task_color_get = StringVar(self.main_win_canvas)
            self.planning_sheet_creater_task_color_get.set("")

            self.circle_email_automation_task_color_get = StringVar(self.main_win_canvas)
            self.circle_email_automation_task_color_get.set("")

            self.email_package_prep_color_get = StringVar(self.main_win_canvas)
            self.email_package_prep_color_get.set("")

            self.interdomain_kpis_data_prep_color_get = StringVar(self.main_win_canvas)
            self.interdomain_kpis_data_prep_color_get.set("")

            self.interdomain_kpis_mail_communication_color_get = StringVar(self.main_win_canvas)
            self.interdomain_kpis_mail_communication_color_get.set("")

            self.evening_task_color_get = StringVar(self.main_win_canvas)
            self.evening_task_color_get.set("")

            self.executor_mail_communication_color_get = StringVar(self.main_win_canvas)
            self.executor_mail_communication_color_get.set("")

        ########################################################## Status checker flags #######################################################################################################################
            self.planning_sheet_creater_task_status_checker_flag            = 0
            self.circle_email_automation_status_checker_flag                = 0
            self.email_package_prep_task_status_checker_flag                = 0
            self.interdomain_kpis_data_prep_status_checker_flag             = 0
            self.interdomain_kpis_mail_communication_status_checker_flag    = 0
            self.evening_task_status_checker_flag                           = 0
            self.executor_mail_communication_status_checker_flag            = 0
            self.update(1)

        self.change_responsible_text_file_lines = open("./change_responsible.txt","r")
    
        # List of all the users.
        self.acceptable_change_responsible = self.change_responsible_text_file_lines.readlines()
        
        for i in range(0,len(self.acceptable_change_responsible)):
            self.acceptable_change_responsible[i] = self.acceptable_change_responsible[i].strip()

        self.change_responsible_text_file_lines.close()

        # Calling the Method to get the user name via GUI.
        self.get_sender_name(self.acceptable_change_responsible)
    
    # Creating the GUI for getting the User name in the variable named sender.
    def get_sender_name(self,acceptable_change_responsible):
        self.sender_win = Toplevel(self.main_win)                           # Creating the Child GUI Window of the main Window GUI.
        self.main_win.withdraw()                                            # Making the Main Window GUI to hide when the Child GUI window appears.
        self.sender_win.title(" Please Select Your Name To Proceed")        # Setting the title of the GUI Window.
        self.sender_win.iconbitmap("./images/ericsson-blue-icon-logo.ico")  # Setting the Icon to be shown at the title bar
        self.sender_win.geometry("600x150")                                 # Setting the GUI dimensions.
        self.sender_win.minsize(600, 150)                                   # Setting the minimum GUI dimensions.
        self.sender_win.maxsize(600, 150)                                   # Setting the maximum GUI dimensions.    
        self.sender_win.bind("<Escape>", self.sender_win_quit)              # Setting the window to get destroyed when Escape Button is Clicked.
        self.sender_win.bind("<Return>", self.submit_sender_name)           # Setting the Enter key to click the Button for submitting the file browser.
        self.sender_win.protocol(
            "WM_DELETE_WINDOW", lambda: self.sender_win_quit(1))            # Setting protocol for the condition, when the user deliberately closes the child GUI window.
        
        self.sender_win.grab_set()                                          # Focussing on the child GUI Window when the child GUI window appears on the screen.

        # Setting the background image of the child GUI window.
        self.sender_win_background = ImageTk.PhotoImage(
            Image.open("./images/MPBN PLANNING TASK_3_2.png"))
        self.sender_win_canvas = Canvas(
            self.sender_win, width=600, height=150, bd=0, highlightthickness=0, relief="ridge")
        self.sender_win_canvas.grid(row=0, column=0, sticky=NW)

        # Declaring and Initializing String Variable for taking the User Input from the Dropdown list.
        self.sender_win_entry_var = StringVar()
        self.sender_win_entry_var.set("Select Your Name!")
        
        # Creating the dropdown list of available users.
        self.sender_win_entry = ttk.OptionMenu(
            self.sender_win, self.sender_win_entry_var, *acceptable_change_responsible, style="TMenubutton",direction='flush')
        self.sender_win_entry['menu'].config(font=("Ericsson Hilda", 14))
        self.sender_win_entry.config(width=20)
        
        self.sender_win_entry.focus_force()                     # Forcing the Focus on the sender window entry when the Sender Name Child GUI

        # Creating the Button for the Submission of the entry.
        self.sender_win_btn = ttk.Button(
            self.sender_win, text="Submit", command=lambda: self.submit_sender_name(1))
        
        # Creating a button for adding the list of the change reponsible
        self.user_addition_button = ttk.Button(self.sender_win, text = "Add User", command = lambda:self.add_user_func(1))

        # Creating a button for deleting a user
        self.user_deletion_button = ttk.Button(self.sender_win, text = "Remove User", command = lambda:self.delete_user_func(1))

        # Setting the background image over the Child GUI.
        self.sender_win_canvas.create_image(
            0, 0, image=self.sender_win_background, anchor="nw")
        
        # Creating the text label for the Child GUI.
        self.sender_win_canvas.create_text(
            295, 30, text="Please Select your name to Proceed or No to Exit", fill="#FFFFFF", font=("Ericsson Hilda", 19, "bold"))
        
        # Creating the window for the drop down list of the users available and the submit button.
        self.sender_win_canvas.create_window(
            60, 70, anchor="nw", window=self.sender_win_entry)
        self.sender_win_canvas.create_window(
            422, 85, window=self.sender_win_btn)
        
        self.sender_win_canvas.create_window(60,105,anchor = "nw", window=self.user_addition_button)

        self.sender_win_canvas.create_window(378,105,anchor = "nw", window=self.user_deletion_button)

        # Calling the Child GUI Window in a loop until any Interruption event is occured.
        self.sender_win.mainloop()

    # Method(Function) for adding user name
    def add_user_func(self,event):
        self.add_user_win = Toplevel(self.sender_win)
        self.sender_win.withdraw()
        self.add_user_win.bind("<Escape>", self.add_user_quit)
        self.add_user_win.bind("<Alt-F4>", self.add_user_quit)
        self.add_user_win.protocol("WM_DELETE_WINDOW",lambda:self.add_user_quit(1))
        self.add_user_win.geometry("600x150")
        self.add_user_win.minsize(600,150)
        self.add_user_win.maxsize(600,150)
        self.add_user_win.title("   Add User Name")
        self.add_user_win.iconbitmap("./images/ericsson-blue-icon-logo.ico")

        self.add_user_func_canvas = Canvas(self.add_user_win,width = 600, height = 150, relief = "ridge",highlightthickness=0,bd = 0)
        self.add_user_func_canvas.grid(column=0, row=0, sticky=NW)

        self.add_user_win_bg = ImageTk.PhotoImage(
            Image.open("./images/MPBN PLANNING TASK_3_2.png"))
        self.add_user_func_canvas.create_image(0,0,image = self.add_user_win_bg,anchor = "nw")

        self.user_add_entry = ttk.Entry(self.add_user_win,width=30,font=('Ericsson Hilda',12,'normal'))
        self.user_add_submit= ttk.Button(self.add_user_win,text="Submit", command = lambda:self.change_responsible_text_editor(1,'add'))
        self.user_add_submit.bind("<Return>",lambda:self.change_responsible_text_editor(1,'add'))
        self.user_add_entry.focus_force()
        self.add_user_func_canvas.create_text(285,40,text = "Enter the Name you want to add!",fill = "#FFFFFF",font = ('Ericsson Hilda',19,'bold'))
        self.add_user_func_canvas.create_window(200,80,window = self.user_add_entry)
        self.add_user_func_canvas.create_window(400,79,window = self.user_add_submit)
        
    
    # Method for destruction of the add_user_window
    def add_user_quit(self,event):
        self.add_user_win.destroy()
        self.sender_win.destroy()
        self.get_sender_name(self.acceptable_change_responsible)

    # Method(Fuction) for deleting user name
    def delete_user_func(self,event):
        self.delete_user_win = Toplevel(self.sender_win)
        self.sender_win.withdraw()
        self.delete_user_win.bind("<Escape>", self.delete_user_quit)
        self.delete_user_win.bind("<Alt-F4>", self.delete_user_quit)
        self.delete_user_win.protocol("WM_DELETE_WINDOW",lambda:self.delete_user_quit(1))
        self.delete_user_win.geometry("600x150")
        self.delete_user_win.minsize(600,150)
        self.delete_user_win.maxsize(600,150)
        self.delete_user_win.title("    Remove User Name")
        self.delete_user_win.iconbitmap("./images/ericsson-blue-icon-logo.ico")

        self.delete_user_func_canvas = Canvas(self.delete_user_win,width = 600, height = 150, relief = "ridge",highlightthickness=0,bd = 0)
        self.delete_user_func_canvas.grid(column=0, row=0, sticky=NW)

        self.delete_user_win_bg = ImageTk.PhotoImage(
            Image.open("./images/MPBN PLANNING TASK_3_2.png"))
        self.delete_user_func_canvas.create_image(0,0,image = self.delete_user_win_bg,anchor = "nw")

        self.user_delete_entry = ttk.Entry(self.delete_user_win,width=30,font=('Ericsson Hilda',12,'normal'))
        self.user_delete_entry.focus_force()
        self.user_delete_submit= ttk.Button(self.delete_user_win,text="Submit", command = lambda:self.change_responsible_text_editor(1,'delete'))
        self.user_delete_submit.bind("<Return>", lambda:self.change_responsible_text_editor(1,'delete'))
        self.delete_user_func_canvas.create_text(285,40,text = "Enter the Name you want to remove!",fill = "#FFFFFF",font = ('Ericsson Hilda',19,'bold'))
        self.delete_user_func_canvas.create_window(200,80,window = self.user_delete_entry)
        self.delete_user_func_canvas.create_window(400,79,window =self.user_delete_submit)

    # Method for destruction of the add_user_window
    def delete_user_quit(self,event):
        self.delete_user_win.destroy()    
        self.sender_win.destroy()
        self.get_sender_name(self.acceptable_change_responsible)

    # Method for updation of the change responsible text
    def change_responsible_text_editor(self,event,task):
        file_read_for_change_responsible = open("./change_responsible.txt")
        
        change_responsible_lines = file_read_for_change_responsible.readlines()

        for i in range(0,len(change_responsible_lines)):
            change_responsible_lines[i] = change_responsible_lines[i].strip()
        
        match task:
            case 'add':
                self.user_add_entry_var = str(self.user_add_entry.get()).strip()
                if(len(self.user_add_entry_var) == 0):
                    messagebox.showerror("  Empty Field!","Kindly enter name you want to add, not empty space!")
                    self.user_add_entry.delete(0,END)
                else:
                    change_responsible_lines.insert(-1,self.user_add_entry_var)
            case 'delete':
                self.user_delete_entry_var = str(self.user_delete_entry.get()).strip()
                if((len(self.user_delete_entry_var) == 0) or (self.user_delete_entry_var.strip().upper() == "NO") or (self.user_delete_entry_var.strip().upper() == "SELECT YOUR NAME!")):
                    messagebox.showerror("  Empty Field!","Kindly enter name you want to remove, not empty space or 'NO'!")
                    self.user_delete_entry.delete(0,END)
                else:
                    if(self.user_delete_entry_var.strip() in change_responsible_lines):
                        change_responsible_lines.remove(self.user_delete_entry_var)
                    else:
                        messagebox.showerror("  Name Not found!","The Name that has been entered is not present in the acceptable change responsible list, Kindly Check!")
                        self.user_delete_entry.delete(0,END)

            case _:
                pass

        file_read_for_change_responsible.close()
        file_write_for_change_responsible = open("./change_responsible.txt","w")
        text_to_be_entered = ""
        for i in range(0,len(change_responsible_lines)):
            text_to_be_entered = f"{text_to_be_entered}{change_responsible_lines[i]}\n"
        file_write_for_change_responsible.write(text_to_be_entered)
        file_write_for_change_responsible.close()

        change_responsible_text_file_lines = open("./change_responsible.txt","r")
    
        # List of all the users.
        self.acceptable_change_responsible = change_responsible_text_file_lines.readlines()
        
        for i in range(0,len(self.acceptable_change_responsible)):
            self.acceptable_change_responsible[i] = self.acceptable_change_responsible[i].strip()

        change_responsible_text_file_lines.close()

        match task:
            case 'add':
                messagebox.showinfo("   Task Successful","User Name successfully added!")
                self.add_user_quit(1)
            case 'delete':
                messagebox.showinfo("   Task Successful","User Name successfully removed!")
                self.delete_user_quit(1)

    # Method(Function) for updating the GUI Components (GIF Frame and the Background Image).
    def update(self, ind):
        self.frame = self.frames[ind]                               # Getting the frame of the GIF to be shown.
        ind += 1                                                    # Increamenting the frame number of the GIF to get the next GIF frame.
        
        '''
            Creating the Integer Variable for getting the frame number of the GIF to be shared between different methods so that when 
            the task is completed the GIF frame which was shown before clicking on the task button can be shown again, and the GIF doesn't start from beginning.
        '''
        self.frame_idx = IntVar()                                   
        self.frame_idx.set(ind)                                     # Setting the Value of the Integer Variable of Frame Index to get the GIF Frame Number
        self.main_win_canvas.delete("all")                          # Deleting all the GUI Components (GIF Frame and the Background Image) to show the smooth transition of the GIF Frame from one frame to another.

        # Setting the GIF Image Frame onto GUI along with the Background Image.
        self.main_win_canvas.create_image(
            0, 0, image=self.main_win_background_img, anchor="nw")
        self.main_win_canvas.create_image(
            870, 2, image=self.frame, anchor="nw")

        '''
            Creating the labels for selection of ITSM Raw Report and the MPBN Planning Sheet workbook.
        '''
        self.main_win_canvas.create_text(220, 194, text="Select ITSM RAW Report :-", 
                                         fill="#FFFFFF", font=("Ericsson Hilda ExtraBold",21,"bold underline"))
        self.main_win_canvas.create_text(220, 247, text="Select Daily Planning Sheet :-",
                                         fill="#FFFFFF", font=("Ericsson Hilda ExtraBold", 21, "bold underline"))
        
        '''
            Creating window in the canvas for ITSM Raw Report CSV browser along with it's label in front of the entry indicating that 
            this is the entry where the CSV needs to be selected.
        '''
        self.main_win_canvas.create_window(
            420, 182, anchor="nw", window=self.itsm_file_browser_entry)
        self.main_win_canvas.create_window(
            840, 180, anchor="nw", window=self.itsm_file_browser_btn)
        
        '''
            Creating window in the canvas for MPBN Planning sheet workbook browser along with it's label in front of the entry indicating that 
            this is the entry where the workbook needs to be selected.
        '''
        self.main_win_canvas.create_window(
            420, 236, anchor="nw", window=self.file_browser_entry)
        self.main_win_canvas.create_window(
            840, 234, anchor="nw", window=self.file_browser_btn)

        # Creating the Window for the Planning Sheet creater button to be shown over the Canvas.
        self.main_win_canvas.create_window(
            90, 298, anchor="nw", window = self.planning_sheet_creater_task_btn)
        self.main_win_canvas.create_text(145, 330, anchor="nw", text = self.planning_sheet_creater_task_status.get(
        ), fill = self.planning_sheet_creater_task_color_get.get(), font=("Ericsson Hilda ExtraBold", 15, "bold"))

        # Creating the Window for the Circle Email Automation button to be shown over the Canvas.
        self.main_win_canvas.create_window(
            385, 298, anchor="nw", window = self.circle_email_automation_task_btn)
        self.main_win_canvas.create_text(440, 330, anchor="nw", text = self.circle_email_automation_task_status.get(
        ), fill = self.circle_email_automation_task_color_get.get(), font = ("Ericsson Hilda ExtraBold", 15, "bold"))

        # Creating the Window for the Email package prep button to be shown over the Canvas.
        self.main_win_canvas.create_window(
            700, 298, anchor="nw", window = self.email_package_prep_btn)
        self.main_win_canvas.create_text(730, 330, anchor="nw", text = self.email_package_prep_task_status.get(
        ), fill = self.email_package_prep_color_get.get(), font=("Ericsson Hilda ExtraBold", 15, "bold"))

        # Creating the Window for the Interdomain Kpi Mail button to be shown over the Canvas.
        self.main_win_canvas.create_window(
            90, 378, anchor="nw", window = self.interdomain_kpis_data_prep_btn)
        self.main_win_canvas.create_text(145, 410, anchor="nw", text = self.interdomain_kpis_data_prep_task_status.get(
        ), fill = self.interdomain_kpis_data_prep_color_get.get(), font=("Ericsson Hilda ExtraBold", 15, "bold"))

        # Creating the Window for the Evening message task button to be shown over the Canvas.
        self.main_win_canvas.create_window(
            385, 378, anchor="nw", window=self.interdomain_kpis_mail_communication_btn)
        self.main_win_canvas.create_text(440, 410, anchor="nw", text = self.interdomain_kpis_mail_communication_status.get(
        ), fill = self.interdomain_kpis_mail_communication_color_get.get(), font = ("Ericsson Hilda ExtraBold", 15, "bold"))  

        # Creating the Window for the Evening message task button to be shown over the Canvas.
        self.main_win_canvas.create_window(
            700, 378, anchor="nw", window=self.evening_task_btn)
        self.main_win_canvas.create_text(730, 410, anchor="nw", text = self.evening_task_status.get(
        ), fill = self.evening_task_color_get.get(), font = ("Ericsson Hilda ExtraBold", 15, "bold"))

        # Creating the Window for the Executor Circle mail communication button to be shown over the Canvas.
        self.main_win_canvas.create_window(
            385, 458, anchor="nw", window=self.executor_mail_communication_btn)
        self.main_win_canvas.create_text(440, 490, anchor="nw", text = self.executor_mail_communication_status.get(
        ), fill = self.executor_mail_communication_color_get.get(), font = ("Ericsson Hilda ExtraBold", 15, "bold"))

        # Solves the flickering problem when the frame gets updated by updating the idle tasks along with the GUI Components to behave in the intended way.
        self.main_win_canvas.update_idletasks()
        
        # Setting the GIF Frame Number value back to 1 when the last frame is reached so that the endless loop continues until the Application is running.
        if ind == 31:
            ind = 1

        # Checking if the main_win_flag is 0 or not.
        # If the Main Win Flag is not 0, that means the Child GUI window is called when the last frame of the GUI was being shown.
        if self.main_win_flag != 0:
            '''
                The Tkinter after method is used to trigger a function after a certain amount of time in case of using sleep(). In our case we are calling the update method after
                certain frame time, when the Child GUI is called and sending the frame value at which the next GIF frame should start from, in this case the last frame was already 
                reached so the next frame which should be shown was already selected in the 'ind' variable which is sent as an argument to the 'update' method.
            '''
            self.main_win.after(31, self.update, ind)

        # If the Main Win Flag is 0, that means the Child GUI window is called when frame other than the last frame of the GUI was being shown.
        if self.main_win_flag == 0:
            '''
                In this case the Frame Value is something in between the first & last frame of the GIF when the Child GUI was called so we are just using 'after' method 
                at the 'ind' frame and the next frame which will be shown next is already selected in the 'ind' variable and is sent as an argument to the 'update' method.
            '''
            self.main_win.after(ind, self.update, ind)         
    
    # Method(Function) for browsing the MPBN Planning Task Workbook.
    def file_browser_func(self, event):
        # Deleting the previous entry in the entry box where the path of the file is given.
        self.file_browser_entry.delete(0, END)
        
        # Creating a browser window for searching and selecting the file.
        self.mystring = filedialog.askopenfilename(initialdir="C:\\", title="  Select the worksheet", filetypes=(
            ("Excel Files (.xlsx)", "*.xlsx"), ("Excel Files (.xls)", "*.xls"), ("All Files", "*.*")))
        
        # Setting the File Browser Entry value to the path selected.
        self.file_browser_entry.insert(0, self.mystring)

        # Setting the value of the File Browser variable to the path selected so that the path selected can be used inter-methodly.
        self.file_browser_file = self.mystring
    
    # Method(Function) for browsing the ITSM Report CSV file.
    def itsm_file_browser_func(self, event):
        # Deleting the previous entry in the entry box where the path of the file is given.
        self.itsm_file_browser_entry.delete(0, END)

        # Creating a browser window for searching and selecting the file.
        self.mystring = filedialog.askopenfilename(initialdir="C:\\", title="  Select report csv file", filetypes=(
            ("CSV Files (.csv)", "*.csv"), ("All Files", "*.*")))
        
        # Setting the File Browser Entry value to the path selected.
        self.itsm_file_browser_entry.insert(0, self.mystring)

        # Setting the value of the ITSM File Browser variable to the path selected so that the path selected can be used inter-methodly.
        self.itsm_file_browser_file = self.mystring
    
    # Method(Function) for calling the module for Planning Sheet creation from the report csv.
    def planning_sheet_creater_task_func(self):
        if(self.task_running == 0):
        # Checking the status of the planning sheet creation whether it's done or not.
            if (self.planning_sheet_creater_task_status_checker_flag == 0):
                try:
                    self.task_running = 1
                    self.task_module_running = "Planning Sheet Preparation"

                    # Setting the task status label to 'In Progress' and setting it's color.
                    self.planning_sheet_creater_task_color_get.set(self.color[2])
                    self.planning_sheet_creater_task_status.set(" In Progress ")

                    # Checking if the workbook for the MPBN Planning Sheet is selected or not
                    if (len(self.file_browser_file) == 0):
                        # Raising the Exception for file not being selected.
                        raise FileNotSelected(
                            " Please Select the MPBN Planning Workbook first!", "File Not Selected")
                    
                    # Checking if the ITSM Raw Report CSV file is selected or not.
                    if (len(self.itsm_file_browser_file) == 0):
                        # Raising the Exception for file not being selected.
                        raise FileNotSelected(
                            " Please Select the ITSM Raw Report first!", "File Not Selected")

                    else:
                        # Calling the method of the module for planning sheet creation from the Raw Report CSV and getting the return value of the 
                        # status of the Task in status flag.
                        self.planning_sheet_creater_task_status_flag = Planning_sheet_creater.planning_sheet_creater(
                            self.itsm_file_browser_file, self.file_browser_file, self.sender)

                        # Checking if the status of the task is successful or not.
                        if (self.planning_sheet_creater_task_status_flag == "Successful"):
                            # Setting the label for task to successful.
                            self.planning_sheet_creater_task_status.set(
                                " Successful ")
                            
                            # Setting the color of the Successful label
                            self.planning_sheet_creater_task_color_get.set(
                                self.color[1])
                            
                            # Setting the status checker flag of the task to 1 indicating that this task has been successfully created
                            # and need not to run this task again.
                            self.planning_sheet_creater_task_status_checker_flag = 1
                            self.task_running = 0
                            self.task_module_running = ""

                        # If the status flag is Unsuccessful then the label for the task is set to Unsuccessful and it's color is set red.
                        if (self.planning_sheet_creater_task_status_flag == "Unsuccessful"):
                            self.planning_sheet_creater_task_status.set(
                                " Unsuccessful ")
                            self.planning_sheet_creater_task_color_get.set(
                                self.color[0])
                            self.planning_sheet_creater_task_status_checker_flag = 0
                            self.task_running = 0
                            self.task_module_running = ""

                # Handling the Exception for file being not selected and setting the label to unsuccessful along with it's color.
                except FileNotSelected:
                    self.planning_sheet_creater_task_color_get.set(self.color[0])
                    self.planning_sheet_creater_task_status_checker_flag = 0
                    self.planning_sheet_creater_task_status.set(" Unsuccessful ")
                    self.task_running = 0
                    self.task_module_running = ""

                # Handling any other Exception and setting the label to unsuccessful along with it's color.
                except Exception as error:
                    messagebox.showerror(" Exception Occured", error)
                    self.planning_sheet_creater_task_color_get.set(self.color[0])
                    self.planning_sheet_creater_task_status_checker_flag = 0
                    self.planning_sheet_creater_task_status.set(" Unsuccessful ")
                    self.task_running = 0
                    self.task_module_running = ""

            else:
                self.task_running = 0
                self.task_module_running = ""
                # Raising the Custom warning in case the task is already successfuly completed.
                raise CustomWarning("  Planning Sheet Creation Task Already Successfully Completed!", " Task Already Done")
        else:
            messagebox.showwarning("    Another task is running!",f"{self.task_module_running} is already running, Please Wait Patiently!")
        
    # Method(Function) for checking the surity from the user that if he wants to continue or not.
    def circle_email_automation_task_func_surity(self):
        # Checking if another module is not running
        if(self.task_running == 0):
            if (self.circle_email_automation_status_checker_flag == 0):
                self.task_running = 1
                self.task_module_running = "Circle Mail Communication"

                # Taking the response from the User.
                self.circle_email_automation_task_surity_check = messagebox.askyesno(
                    "  Circle Mail Confirmation", "Do you want to proceed for Email Communication for Tonight Planned Circles ?")
                
                # If the respose is positive the task is done, else the label for task status is set to Unsuccessful along with it's color.
                if (self.circle_email_automation_task_surity_check):
                    self.circle_email_automation_task_func()
                    self.task_running = 0
                    self.task_module_running = ""
                
                else:
                    self.task_running = 0
                    self.task_module_running = ""
                    self.circle_email_automation_task_status.set(
                                    " Unsuccessful ")
                    self.circle_email_automation_task_color_get.set(
                        self.color[0])
                    self.circle_email_automation_status_checker_flag = 0
                    self.task_running = 0
                    self.task_module_running = ""
        
            else:
                self.task_running = 0
                self.task_module_running = ""
                # Raising the Custom warning in case the task is already successfuly completed.
                raise CustomWarning("  Circle Automation Task Already Successfully Completed!", " Task Already Done")

        else:
            messagebox.showwarning("    Another Task is running!",f"{self.task_module_running} is already running, Please Wait Patiently!")

    # Method(Function) for Circle Email Automation Task.
    def circle_email_automation_task_func(self):
        try:
            # Setting the task status label to 'In Progress' and setting it's color.
            self.circle_email_automation_task_color_get.set(self.color[2])
            self.circle_email_automation_task_status.set(" In Progress ")

            # Checking if the workbook for the MPBN Planning Sheet is selected or not
            if (len(self.file_browser_file) == 0):
                # Raising the Exception for file not being selected.
                raise FileNotSelected(
                    " Please Select the MPBN Planning Workbook first!", "File Not Selected")

            else:
                # Calling the method of the module for circle email automation from the MPBN Planning sheet workbook and getting the return value of the 
                # status of the Task in status flag.
                self.circle_email_automation_status_flag = circle_Email_Automation_Task.fetch_details(
                    self.sender, self.file_browser_file)

                # Checking if the status of the task is successful or not.
                if (self.circle_email_automation_status_flag == "Successful"):
                    # Setting the label for task to successful.
                    self.circle_email_automation_task_status.set(
                        " Successful ")
                    
                    # Setting the color of the Successful label
                    self.circle_email_automation_task_color_get.set(
                        self.color[1])

                    # Setting the status checker flag of the task to 1 indicating that this task has been successfully created
                    # and need not to run this task again.
                    self.circle_email_automation_status_checker_flag = 1
                    self.task_running = 0
                    self.task_module_running = ""

                # If the status flag is Unsuccessful then the label for the task is set to Unsuccessful and it's color is set red.
                if (self.circle_email_automation_status_flag == "Unsuccessful"):
                    self.circle_email_automation_task_status.set(
                        " Unsuccessful ")
                    self.circle_email_automation_task_color_get.set(
                        self.color[0])
                    self.circle_email_automation_status_checker_flag = 0
                    self.task_running = 0
                    self.task_module_running = ""

        # Handling the Exception for file being not selected and setting the label to unsuccessful along with it's color.
        except FileNotSelected:
            self.circle_email_automation_task_color_get.set(self.color[0])
            self.circle_email_automation_status_checker_flag = 0
            self.circle_email_automation_task_status.set(" Unsuccessful ")
            self.task_running = 0
            self.task_module_running = ""

        # Handling any other Exception and setting the label to unsuccessful along with it's color.
        except Exception as error:
            messagebox.showerror(" Exception Occured", error)
            self.circle_email_automation_task_color_get.set(self.color[0])
            self.circle_email_automation_status_checker_flag = 0
            self.circle_email_automation_task_status.set(" Unsuccessful ")
            self.task_running = 0
            self.task_module_running = ""

        
    # Method(Function) for Email-Package Preparation.
    def email_package_prep_func(self):
        if(self.task_running == 0):
        # Checking the status of the email package prep task whether it's done or not.
            if (self.email_package_prep_task_status_checker_flag == 0):
                try:
                    self.task_running = 1
                    self.task_module_running = "Email Package Preparation"

                    # Setting the task status label to 'In Progress' and setting it's color.
                    self.email_package_prep_color_get.set(self.color[2])
                    self.email_package_prep_task_status.set(" In Progress ")

                    # Checking if the workbook for the MPBN Planning Sheet is selected or not
                    if (len(self.file_browser_file) == 0):
                        # Raising the Exception for file not being selected.
                        raise FileNotSelected(
                            " Please Select the MPBN Planning Workbook first!", "File Not Selected")

                    else:
                        # Calling the method of the module for circle email automation from the MPBN Planning sheet workbook and getting the return value of the 
                        # status of the Task in status flag.
                        self.email_package_status_flag = Email_package_generator.email_package_sheet_creater(
                            self.file_browser_file)

                        # Checking if the status of the task is successful or not.
                        if (self.email_package_status_flag == "Successful"):
                            # Setting the label for task to successful.
                            self.email_package_prep_task_status.set(
                                " Successful ")
                            
                            # Setting the color of the Successful label
                            self.email_package_prep_color_get.set(
                                self.color[1])

                            # Setting the status checker flag of the task to 1 indicating that this task has been successfully created
                            # and need not to run this task again.
                            self.email_package_prep_task_status_checker_flag = 1
                            
                            self.task_running = 0
                            self.task_module_running = ""

                        # If the status flag is Unsuccessful then the label for the task is set to Unsuccessful and it's color is set red.
                        if (self.email_package_status_flag == "Unsuccessful"):
                            self.email_package_prep_task_status.set(
                                " Unsuccessful ")
                            self.email_package_prep_color_get.set(
                                self.color[0])
                            self.email_package_prep_task_status_checker_flag = 0
                            
                            self.task_running = 0
                            self.task_module_running = ""

                # Handling the Exception for file being not selected and setting the label to unsuccessful along with it's color.
                except FileNotSelected:
                    self.email_package_prep_color_get.set(self.color[0])
                    self.email_package_prep_task_status_checker_flag = 0
                    self.email_package_prep_task_status.set(" Unsuccessful ")
                    self.task_running = 0
                    self.task_module_running = ""

                # Handling any other Exception and setting the label to unsuccessful along with it's color.
                except Exception as error:
                    messagebox.showerror(" Exception Occured", error)
                    self.email_package_prep_color_get.set(self.color[0])
                    self.email_package_prep_task_status_checker_flag = 0
                    self.email_package_prep_task_status.set(" Unsuccessful ")
                    self.task_running = 0
                    self.task_module_running = ""

            else:
                self.task_running = 0
                self.task_module_running = ""
                # Raising the Custom warning in case the task is already successfuly completed.
                raise CustomWarning("  Email-Package Preparation Task Already Successfully Completed!", " Task Already Done")
        
        else:
            messagebox.showwarning("    Another task is running!",f"{self.task_module_running} is already running, Please Wait Patiently!")

    # Method(Function) for Preparing the Interdomain KPIs data
    def interdomain_kpis_data_prep_func(self):
        # Checking If another task is running or not
        if (self.task_running == 0):
            # Checking the status of the Interdomain kpis data preparation whether it's done or not.
            if (self.interdomain_kpis_data_prep_status_checker_flag == 0):
                # Setting the task status label to 'In Progress' and setting it's color.
                self.interdomain_kpis_data_prep_color_get.set(self.color[2])
                self.interdomain_kpis_data_prep_task_status.set(" In Progress ")
                self.task_running = 1
                self.task_module_running = "Interdomain KPIs Data Preparation"

                try:
                    # Checking if the workbook for the MPBN Planning Sheet is selected or not.
                    if (len(self.file_browser_file) == 0):
                        # Raising the Exception for file not being selected
                        raise FileNotSelected(
                            " Please Select the MPBN Planning Excel Workbook first!", "File Not Selected")

                    else:
                        # Calling the method of the module for Interdomain Kpis data preparation from the MPBN Planning sheet workbook and getting the return value of the 
                        # status of the Task in status flag.
                        self.interdomain_kpis_data_prep_status_flag = interdomain_KPIs_Data_Prep_Task.paco_cscore(
                            self.sender, self.file_browser_file)

                        # Checking if the status of the task is successful or not.
                        if (self.interdomain_kpis_data_prep_status_flag == 'Successful'):
                            # Setting the color of the Successful label
                            self.interdomain_kpis_data_prep_color_get.set(
                                self.color[1])
                            
                            # Setting completion status of interdomain_kpis_data_prep_task to 1 for further usage  with other methods.
                            self.interdomain_kpis_data_prep_task_completed = 1
                            
                            # Setting the status checker flag of the task to 1 indicating that this task has been successfully created
                            # and need not to run this task again.
                            self.interdomain_kpis_data_prep_status_checker_flag = 1
                            
                            # Setting the label for task to successful.
                            self.interdomain_kpis_data_prep_task_status.set(
                                " Successful ")
                            
                            self.task_running = 0
                            self.task_module_running = ""

                        # If the status flag is Unsuccessful then the label for the task is set to Unsuccessful and it's color is set red.
                        elif (self.interdomain_kpis_data_prep_status_flag == 'Unsuccessful'):
                            self.interdomain_kpis_data_prep_color_get.set(
                                self.color[0])
                            self.interdomain_kpis_data_prep_task_completed = 0
                            self.interdomain_kpis_data_prep_status_checker_flag = 0
                            self.interdomain_kpis_data_prep_task_status.set(
                                " Unsuccessful ")
                            
                            self.task_running = 0
                            self.task_module_running = ""

                # If the status flag is Unsuccessful then the label for the task is set to Unsuccessful and it's color is set red.
                except FileNotSelected:
                    self.interdomain_kpis_data_prep_color_get.set(self.color[0])
                    self.interdomain_kpis_data_prep_status_checker_flag = 0
                    self.interdomain_kpis_data_prep_task_status.set(
                        " Unsuccessful ")
                    self.task_running = 0
                    self.task_module_running = ""
                
                # Handling any other Exception and setting the label to unsuccessful along with it's color.
                except Exception as error:
                    messagebox.showerror(" Exception Occured", error)
                    self.interdomain_kpis_data_prep_color_get.set(self.color[0])
                    self.interdomain_kpis_data_prep_status_checker_flag = 0
                    self.interdomain_kpis_data_prep_task_status.set(" Unsuccessful ")
                    self.task_running = 0
                    self.task_module_running = ""

            else:
                self.task_running = 0
                self.task_module_running = ""
                # Raising the Custom warning in case the task is already successfuly completed.
                raise CustomWarning(
                    " Interdomain KPIs Data Prep Task Already Successfully Completed", " Task Already Done")

        else:
            messagebox.showwarning("    Another task is running!",f"{self.task_module_running} is already running, Please Wait Patiently!")

    # Method(Function) for Interdomain KPIs Mail communication
    def interdomain_kpis_mail_communication_func(self, event):
        if(self.task_running == 0):
            if (self.interdomain_kpis_mail_communication_status_checker_flag == 0):
                self.interdomain_kpis_mail_communication_surity_check = messagebox.askyesno("   Interdomain Mail Communication Confirmation","Do you to proceed with Interdomain Mail Communication for tonight planned CRs?")

                if(self.interdomain_kpis_mail_communication_surity_check):
                    self.task_running = 1
                    self.task_module_running = "Interdomain KPIs Mail Communication"

                    self.region_handler_names_win = Toplevel(self.main_win)

                    # Hiding the main GUI window.
                    if self.main_win.state() == "normal":
                        self.main_win.withdraw()

                    # Checking the interdomain data prep task is cmpleted or not.
                    if (self.interdomain_kpis_data_prep_task_completed == 1):

                        # Creating New Child GUI for taking MPBN Planning Spocs' names
                        self.region_handler_names_win.geometry("440x300")
                        self.region_handler_names_win.minsize(440, 300)
                        self.region_handler_names_win.maxsize(440, 300)
                        self.region_handler_names_win.iconbitmap(
                            "./images/ericsson-blue-icon-logo.ico")
                        self.region_handler_names_win.title(
                            "   Names for (PAN INDIA) MPBN Planning SPOC's ")
                        self.region_handler_names_win.bind(
                            "<Escape>", self.region_handler_names_win_quit)

                        self.region_handler_names_win_background = ImageTk.PhotoImage(
                            Image.open("./images/MPBN PLANNING TASK_3_3.png"))
                        self.region_handler_names_win_canvas = Canvas(
                            self.region_handler_names_win, width=440, height=300, bd=0, highlightthickness=0, relief="ridge")
                        self.region_handler_names_win_canvas.grid(
                            row=0, column=0, sticky=NW)
                        self.region_handler_names_win_canvas.create_image(
                            0, 0, image=self.region_handler_names_win_background, anchor="nw")

                        # Enteries for the MPBN Planning Spoc names
                        self.north_and_west_region_entry = ttk.Entry(
                            self.region_handler_names_win_canvas, width=40, font=("Ericsson Hilda", 13))
                        self.region_handler_names_win_canvas.create_text(
                            10, 20, anchor="nw", text="Please Enter Name Of North and West Region Planner", fill="#FFFFFF", font=("Ericsson Hilda", 13, "bold"))
                        self.region_handler_names_win_canvas.create_window(
                            10, 65, anchor="nw", window=self.north_and_west_region_entry)
                        self.north_and_west_region_entry.focus_force()

                        self.region_handler_names_win_canvas.create_text(
                            10, 120, anchor="nw", text="Please Enter Name Of South and East Region Planner", fill="#FFFFFF", font=("Ericsson Hilda", 13, "bold"))
                        self.east_region_and_south_region_entry = ttk.Entry(
                            self.region_handler_names_win_canvas, width=40, font=("Ericsson Hilda", 13))
                        self.region_handler_names_win_canvas.create_window(
                            10, 165, anchor="nw", window=self.east_region_and_south_region_entry)

                        # Getting the names of the MPBN Planning Spocs from user input in the above enteries.
                        self.north_and_west_region = self.north_and_west_region_entry.get()
                        self.east_region_and_south_region = self.east_region_and_south_region_entry.get()

                        # Creating the Submit button for the user to submit the names.
                        self.region_handler_names_win_canvas_submit = ttk.Button(
                            self.region_handler_names_win, text="Submit", command=lambda: self.interdomain_kpis_mail_commmunication_starter_func(1))
                        self.region_handler_names_win_canvas.create_window(
                            380, 270, anchor="se", window=self.region_handler_names_win_canvas_submit)
                        self.region_handler_names_win.bind(
                            "<Return>", self.interdomain_kpis_mail_commmunication_starter_func)

                        self.region_handler_names_win.protocol(
                            "WM_DELETE_WINDOW", lambda: self.region_handler_names_win_quit(1))

                        # Checking if the main GUI Window is hidden or not, if hidden, making it reappear and destroying the Child GUI Window.
                        if self.region_handler_names_win.state() != "normal":
                            if self.main_win.state() != "normal":
                                self.main_win_flag = 0
                                self.main_win.deiconify()
                            self.region_handler_names_win.destroy()

                    else:
                        # Setting the label for task to be "Unsuccessful" and destroying the child GUI(although it will be so fast that user won't be able to see the child GUI window to ever appear.)
                        self.interdomain_kpis_mail_communication_color_get.set(
                            self.color[0])
                        self.interdomain_kpis_mail_communication_status.set(
                            ' Unsuccessful ')
                        self.region_handler_names_win.destroy()

                        if self.main_win.state() != "normal":
                            self.main_win_flag = 0
                            self.main_win.deiconify()
                        self.interdomain_kpis_mail_communication_status_checker_flag = 0

                        self.task_running = 0
                        self.task_module_running = ""

                        # Raising the custom made exception for running the interdomain data prep task first.
                        raise CustomException(
                            "Please! Run Interdomain KPIs Data Prep task First!", "   Task Unsuccessful")

                    self.region_handler_names_win.mainloop()

                else:
                    self.task_running = 0
                    self.task_module_running = ""
                    self.interdomain_kpis_mail_communication_status_checker_flag = 0
                    self.interdomain_kpis_mail_communication_color_get.set(
                        self.color[0])
                    self.interdomain_kpis_mail_communication_status.set(
                        ' Unsuccessful ')
            else:
                # Raising Exception when the task is already done successfully.
                self.task_running = 0
                self.task_module_running = ""
                self.interdomain_kpis_mail_communication_status_checker_flag = 0
                raise CustomWarning(
                    " Interdomain KPIs mail Communication Task Already Successfully Completed!", " Task Already Done")
        
        else:
            messagebox.showwarning("    Another task is running!",f"{self.task_module_running} is already running, Please Wait Patiently!")    
    
    # Method(Function) for interdomain mail communication starter function that takes all the required values and then calls the main method which does the job.
    def interdomain_kpis_mail_commmunication_starter_func(self, event):
        # Getting the entry values given by user input.
        self.east_region_and_south_region = self.east_region_and_south_region_entry.get()
        self.north_and_west_region = self.north_and_west_region_entry.get()

        # Setting the label of the task to 'In Progress' along with it's color.
        self.interdomain_kpis_mail_communication_color_get.set(self.color[2])
        self.interdomain_kpis_mail_communication_status.set(" In Progress ")

        # Emptying the list for appending fields with empty strings and integer values.
        self.new_empty_string_list = []
        self.new_integer_string_list = []

        try:
            # Checking if the MPBN Planning Workbook is not selected.
            if (len(self.file_browser_file) == 0):
                self.interdomain_kpis_mail_communication_status_checker_flag = 0
                # Raising custom made exception for the case where the MPBN Planning Workbook is not selected.
                raise FileNotSelected(
                    " Please Select the MPBN Planning Excel Workbook first!", "File Not Selected")
            '''
                Checking if the required enteries contain empty strings or strings containing integers, if no then the module method is 
                called and if yes, then custom made exception is raised.
            '''
            if (len(self.north_and_west_region) > 0) and (len(self.east_region_and_south_region) > 0):
                if (not (any(c.isdigit() for c in self.north_and_west_region))) and (not (any(c.isdigit() for c in self.east_region_and_south_region))):
                    self.main_win_flag = 0
                    
                    # Destroying the child GUI window and making the main GUI window to reappear. 
                    self.main_win.deiconify()
                    self.region_handler_names_win.destroy()
                    
                    # Calling the Module method to complete the task and setting the task label along with suitable color indicating that the task is successfully completed.
                    self.interdomain_starter_func_task_status_checker_flag = interdomain_KPIs_Mail_Comm_Task.paco_cscore(
                        self.sender, self.file_browser_file, self.north_and_west_region, self.east_region_and_south_region)
                    
                    if(self.interdomain_starter_func_task_status_checker_flag == 'Successful'):
                        self.interdomain_kpis_mail_communication_color_get.set(
                            self.color[1])
                        self.interdomain_kpis_mail_communication_status_checker_flag = 1
                        self.interdomain_kpis_mail_communication_status.set(
                            " Successful ")
                        self.task_running = 0
                        self.task_module_running = ""
                    
                    if (self.interdomain_starter_func_task_status_checker_flag == 'Unsuccessful'):
                        self.interdomain_kpis_mail_communication_color_get.set(
                            self.color[0])
                        self.interdomain_kpis_mail_communication_status_checker_flag = 1
                        self.interdomain_kpis_mail_communication_status.set(
                            " Unsuccessful ")
                        self.task_running = 0
                        self.task_module_running = ""

            if (any(c.isdigit() for c in self.north_and_west_region)):
                self.new_integer_string_list.append(
                    "North & West Region Handler")

            if (any(c.isdigit() for c in self.east_region_and_south_region)):
                self.new_integer_string_list.append(
                    "South & East Region Handler")

            if (len(self.north_and_west_region) == 0):
                self.new_empty_string_list.append(
                    "North & West Region Handler")

            if (len(self.east_region_and_south_region) == 0):
                self.new_empty_string_list.append(
                    "South & East Region Handler")

            if (len(self.new_integer_string_list) > 0) and (len(self.new_empty_string_list) == 0):
                self.interdomain_kpis_mail_communication_color_get.set(
                    self.color[0])
                self.interdomain_kpis_mail_communication_status_checker_flag = 0
                self.interdomain_kpis_mail_communication_status.set(
                    " Unsucessful ")
                
                raise RegionHandlerException(
                    f"Please Enter Valid Name/s, Fields with Numbers are not allowed \nField/Fields with Number: {','.join(self.new_integer_string_list)}")
            
            if (len(self.new_empty_string_list) > 0) and (len(self.new_integer_string_list) == 0):
                self.interdomain_kpis_mail_communication_color_get.set(
                    self.color[0])
                self.interdomain_kpis_mail_communication_status_checker_flag = 0
                self.interdomain_kpis_mail_communication_status.set(
                    " Unsucessful ")

                # Raising Custom made exception
                raise RegionHandlerException(
                    f"Please Enter valid Name/s, Empty Strings are not allowed\nEmpty Field/Fields: {','.join(self.new_empty_string_list)}")
            
            if (len(self.new_empty_string_list) > 0) and (len(self.new_integer_string_list) > 0):
                self.interdomain_kpis_mail_communication_color_get.set(
                    self.color[0])
                self.interdomain_kpis_mail_communication_status_checker_flag = 0
                self.interdomain_kpis_mail_communication_status.set(
                    " Unsucessful ")
                
                # Raising Custom made exception.
                raise RegionHandlerException(
                    f"Please Enter Valid Names, Empty Strings and Numbers are not allowed \nEmpty Field: {','.join(self.new_empty_string_list)} \nField with Number: {','.join(self.new_integer_string_list)}")

        # Handling Custom made exceptions and other exceptions that are not handled by the custom made exceptions.
        except FileNotSelected:
            self.interdomain_kpis_mail_communication_color_get.set(
                self.color[0])
            self.interdomain_kpis_mail_communication_status.set(
                ' Unsuccessful ')
            self.task_running = 0
            self.task_module_running = ""

        except RegionHandlerException:
            self.new_empty_string_list = []
            self.new_integer_string_list = []
            self.north_and_west_region_entry.focus_force()
            self.interdomain_kpis_mail_communication_color_get.set(
                self.color[0])
            self.interdomain_kpis_mail_communication_status.set(
                ' Unsuccessful ')
            self.task_running = 0
            self.task_module_running = ""

        except Exception as error:
            messagebox.showerror(" Exception Occured", error)
            self.interdomain_kpis_mail_communication_color_get.set(
                self.color[0])
            self.interdomain_kpis_mail_communication_status.set(
                ' Unsuccessful ')
            self.task_running = 0
            self.task_module_running = ""

    # Method(Function) for Creating the evening task message.
    def evening_task_func(self, event):
        if (self.task_running == 0):
                # Checking the status of the evening task whether it's done or not.
                if (self.evening_task_status_checker_flag == 0):
                    try:
                        self.task_running = 1
                        self.task_module_running = "Email Package & Evening Message"

                        if (len(self.file_browser_file) == 0):
                            # Raising the Exception for file not being selected
                            raise FileNotSelected(
                                " Please Select the MPBN Planning Excel Workbook first!", "File Not Selected")

                        else:
                            # Creating a variable to check whether the email package is created or not.
                            self.interdomain_kpis_data_prep_creation_status_flag = 0
                            self.workbook = pd.ExcelFile(self.file_browser_file)
                            self.worksheet_names = self.workbook.sheet_names

                            # Finding the Email Package from the workbook and reading it in pandas.
                            for sheet in self.worksheet_names:
                                if (sheet == 'Email-Package'):
                                    self.worksheet = pd.read_excel(
                                        self.workbook, sheet)
                                    self.worksheet['Execution Date'] = pd.to_datetime(self.worksheet['Execution Date'], format = "%m/%d/%Y")
                                    self.worksheet['Execution Date'] =  self.worksheet['Execution Date'].dt.strftime("%m/%d/%Y")
                                    
                                    # Getting Today's maintenance date.
                                    tomorrow = datetime.now() + timedelta(1)
                                    tomorrow = tomorrow.strftime("%m/%d/%Y")
                                    self.worksheet = self.worksheet[self.worksheet['Execution Date'] == tomorrow]
                                    
                                    # If there's data present in the worksheet then changing the value of the email package sheet creation status.
                                    if (len(self.worksheet) > 0):
                                        self.interdomain_kpis_data_prep_creation_status_flag = 1
                                        break

                            if (self.interdomain_kpis_data_prep_creation_status_flag == 0):
                                self.task_running = 0
                                self.task_module_running = ""

                                # Raising the custom made exception for the case when the email package sheet is not created or empty.
                                raise CustomException(
                                    'Kindly Click the Button for Interdomain Kpi Data Prep First!', 'Email-Package Worksheet Empty')

                            else:
                                '''
                                    Creating a child GUI Window to get the required inputs on Night Shift Lead Name, Buffer/Auditor/Trainer Name, and 
                                    Resource on Automation Name.
                                '''
                                self.evening_task_win = Toplevel(self.main_win)
                                if self.main_win.state() == 'normal':
                                    self.main_win.withdraw()
                                self.evening_task_win.iconbitmap(
                                    './images/ericsson-blue-icon-logo.ico')
                                self.evening_task_win.title(
                                    "   Please Enter The Names to Proceed")
                                self.evening_task_win.geometry("600x550")
                                self.evening_task_win.minsize(600, 550)
                                self.evening_task_win.maxsize(600, 550)
                                self.evening_task_win.bind(
                                    "<Escape>", self.evening_task_func_quit)

                                # Setting the background image of the child GUI window.
                                self.evening_task_background = ImageTk.PhotoImage(
                                    Image.open("./images/MPBN PLANNING TASK_3_4.png"))
                                self.evening_task_win_canvas = Canvas(
                                    self.evening_task_win, height=550, width=600, bd=0, highlightthickness=0, relief="ridge")
                                self.evening_task_win_canvas.grid(
                                    row=0, column=0, sticky=NW)
                                self.evening_task_win_canvas.create_image(
                                    0, 0, image=self.evening_task_background, anchor="nw")

                                '''
                                    Creating entry blocks for taking user input for Night Shift Lead Name, Buffer/Auditor/Trainer Name, and 
                                    Resource on Automation Name.
                                '''
                                self.evening_task_win_canvas.create_text(
                                    10, 20, anchor="nw", text="Please Enter Night Shift Lead Name", fill="#FFFFFF", font=("Ericsson Hilda", 18, "bold"))
                                self.evening_task_win_canvas_night_shift_lead_entry = ttk.Entry(
                                    self.evening_task_win_canvas, width=40, font=("Ericsson Hilda", 15))
                                self.evening_task_win_canvas.create_window(
                                    10, 70, anchor="nw", window=self.evening_task_win_canvas_night_shift_lead_entry)

                                self.evening_task_win_canvas.create_text(
                                    10, 150, anchor="nw", text="Please Enter Buffer/Auditor/Trainer Name", fill="#FFFFFF", font=("Ericsson Hilda", 18, "bold"))
                                self.evening_task_win_canvas_buffer_auditor_trainer_entry = ttk.Entry(
                                    self.evening_task_win_canvas, width=40, font=("Ericsson Hilda", 15))
                                self.evening_task_win_canvas.create_window(
                                    10, 200, anchor="nw", window=self.evening_task_win_canvas_buffer_auditor_trainer_entry)

                                self.evening_task_win_canvas.create_text(
                                    10, 280, anchor="nw", text="Please Enter Resource on Automation Name", fill="#FFFFFF", font=("Ericsson Hilda", 18, "bold"))
                                self.evening_task_win_canvas_resource_on_automation_entry = ttk.Entry(
                                    self.evening_task_win_canvas, width=40, font=("Ericsson Hilda", 15))
                                self.evening_task_win_canvas.create_window(
                                    10, 310, anchor="nw", window=self.evening_task_win_canvas_resource_on_automation_entry)

                                # Creating submit button calling the driver method after taking all the valid user inputs.
                                self.evening_task_submit_btn = ttk.Button(
                                    self.evening_task_win, text="Submit", command=lambda: self.evening_task_func_starter(1))
                                self.evening_task_win_canvas.create_window(
                                    580, 520, window=self.evening_task_submit_btn, anchor="se")

                                # Focussing on the Night Shift Lead entry when the child GUI Window appears.
                                self.evening_task_win_canvas_night_shift_lead_entry.focus_force()

                                # Setting protocol for the window destruction.
                                self.evening_task_win.protocol(
                                    "WM_DELETE_WINDOW", lambda: self.evening_task_func_quit(1))
                                
                                # Binding the enter key to the driver method.
                                self.evening_task_win.bind(
                                    "<Return>", self.evening_task_func_starter)

                                # Making the main GUI window to reappear, while destroying the child GUI.
                                if self.evening_task_win.state() != "normal":
                                    if self.main_win.state() != "normal":
                                        self.main_win_flag = 0
                                        self.main_win.deiconify()
                                    self.evening_task_win.destroy()

                                # Creating an endless loop until the user presses the submit button or the Enter key or any external interruption occurs.
                                self.evening_task_win.mainloop()

                    
                
                    #  Handling Exceptions for the task.
                    except FileNotSelected:
                        self.evening_task_color_get.set(self.color[0])
                        self.evening_task_status_checker_flag = 0
                        self.evening_task_status.set(' Unsuccessful ')
                        self.task_running = 0
                        self.task_module_running = ""
                    
                    except Exception as error:
                        messagebox.showerror(" Exception Occured", error)
                        self.evening_task_color_get.set(self.color[0])
                        self.evening_task_status_checker_flag = 0
                        self.evening_task_status.set(' Unsuccessful ')
                        self.task_running = 0
                        self.task_module_running = ""

                else:
                    self.task_running = 0
                    self.task_module_running = ""

                    # Raising custom made exception for the condition when the task has already been done.
                    raise CustomWarning(
                        "Evening Task Already Successfully Completed!", " Task Already Done")
        else:
            messagebox.showwarning("    Another task is running!",f"{self.task_module_running} is already running, Please Wait Patiently!")

    # Method(Function) for quitting the evening message task while destroying the child GUI Window.
    def evening_task_func_quit(self, event):
        self.task_running = 0
        self.task_module_running = ""
        self.evening_task_win.withdraw()
        self.evening_task_color_get.set(self.color[0])
        self.evening_task_status.set(' Unsuccessful ')
        self.main_win_flag = 0
        self.main_win.deiconify()
        self.evening_task_win.destroy()

    # Method(Function) for starting the evening task.
    def evening_task_func_starter(self, event):
        # Getting all the required entry fields from the user via String variables.
        self.night_shift_lead = self.evening_task_win_canvas_night_shift_lead_entry.get()
        self.buffer_auditor_trainer = self.evening_task_win_canvas_buffer_auditor_trainer_entry.get()
        self.resource_on_automation = self.evening_task_win_canvas_resource_on_automation_entry.get()
        
        # Setting the status to be "In Progress" and ,coloring it white.
        self.evening_task_color_get.set(self.color[2])
        self.evening_task_status.set(' In Progress ')

        # List for containing fields that are left empty by the user or contains integer.
        self.empty_string_list = []
        self.integer_string_list = []

        try:
            # Checking that whether night shift lead, buffer auditor trainer and the resource on automation conatain empty string or integer.
            # If no, then going forward with the task.
            if (len(self.night_shift_lead) > 0) and (len(self.buffer_auditor_trainer) > 0) and (len(self.resource_on_automation) > 0):
                if (not (any(c.isdigit() for c in self.night_shift_lead))) and (not (any(c.isdigit() for c in self.buffer_auditor_trainer))) and (not (any(c.isdigit() for c in self.resource_on_automation))):
                    
                    # Setting the main win flag to 0 for GIF frame, destroying the child GUI window and making the main GUI window to reappear.
                    self.main_win_flag = 0
                    self.main_win.deiconify()
                    self.evening_task_win.destroy()

                    # Calling the required method from the module with sufficient arguments to do the task and getting the return value from the metod in a flag variable.
                    self.evening_mail_task_status_flag = evening_mail_task.evening_task(
                        self.sender, self.night_shift_lead, self.buffer_auditor_trainer, self.resource_on_automation, self.file_browser_file)

                    # Checking the flag variable for setting the label for the task along with it's suitable color.
                    if (self.evening_mail_task_status_flag == 'Successful'):
                        self.evening_task_color_get.set(self.color[1])
                        self.evening_task_status_checker_flag = 1
                        self.evening_task_status.set(' Successful ')
                        self.task_running = 0
                        self.task_module_running = ""

                    if (self.evening_mail_task_status_flag == 'Unsuccessful'):
                        self.evening_task_color_get.set(self.color[0])
                        self.evening_task_status_checker_flag = 0
                        self.evening_task_status.set(' Unsuccessful ')
                        self.task_running = 0
                        self.task_module_running = ""

            # Checking whether the name enteries contain empty strings(no input), or integers and raising custom made exceptions accordingly.
            if (len(self.night_shift_lead) == 0):
                self.empty_string_list.append("Night Shift Lead")

            if (any(c.isdigit() for c in self.night_shift_lead)):
                self.integer_string_list.append("Night Shift Lead")

            if (len(self.buffer_auditor_trainer) == 0):
                self.empty_string_list.append("Buffer/Auditor/Trainer")
            if (any(c.isdigit() for c in self.buffer_auditor_trainer)):
                self.integer_string_list.append("Buffer/Auditor/Trainer")

            if (len(self.resource_on_automation) == 0):
                self.empty_string_list.append("Resource on Automation")
            if (any(c.isdigit() for c in self.resource_on_automation)):
                self.integer_string_list.append("Resource On Automation")

            if (len(self.empty_string_list) > 0) and (len(self.integer_string_list) == 0):
                self.evening_task_color_get.set(self.color[0])
                self.evening_task_status_checker_flag = 0
                self.evening_task_status.set(' Unsuccessful ')
                raise EveningTaskException(
                    f"Please Enter Valid Names, Empty Strings are not allowed\nEmpty Field/Fields: {','.join(self.empty_string_list)}")
            
            if (len(self.empty_string_list) == 0) and (len(self.integer_string_list) > 0):
                self.evening_task_color_get.set(self.color[0])
                self.evening_task_status_checker_flag = 0
                self.evening_task_status.set(' Unsuccessful ')
                raise EveningTaskException(
                    f"Please Enter Valid Names, Numbers are not allowed\nField/Fields with Numbers: {','.join(self.integer_string_list)}")
            
            if (len(self.empty_string_list) > 0) and (len(self.integer_string_list) > 0):
                self.evening_task_color_get.set(self.color[0])
                self.evening_task_status_checker_flag = 0
                self.evening_task_status.set(' Unsuccessful ')
                raise EveningTaskException(
                    f"Please Enter Valid Names, Empty Strings and Numbers are not allowed\n Empty Field/Fields: {','.join(self.empty_string_list)}\nField/Fields with Numbers: {','.join(self.integer_string_list)}")

        # Handling the Evening Task Exception 
        except EveningTaskException:
            self.empty_string_list = []
            self.integer_string_list = []
            self.evening_task_color_get.set(self.color[0])
            self.evening_task_status_checker_flag = 0
            self.evening_task_win_canvas_night_shift_lead_entry.focus_force()
            self.evening_task_status.set(' Unsuccessful ')
            self.task_running = 0
            self.task_module_running = ""

        # Handling any other Exception and setting the label to unsuccessful along with it's color.
        except Exception as error:
            messagebox.showerror(" Exception Occured", error)
            self.evening_task_color_get.set(self.color[0])
            self.evening_task_status_checker_flag = 0
            self.evening_task_status.set(' Unsuccessful ')
            self.task_running = 0
            self.task_module_running = ""

    # Method(Function) for sending the execution mail communication
    def executor_mail_communication(self):
        if(self.task_running == 0):
            # Checking the status of the circle email automation task whether it's done or not.
            if (self.executor_mail_communication_status_checker_flag == 0):
                try:
                    self.task_running = 1
                    self.task_module_running = "Executor Mail Communication"

                    # Setting the task status label to 'In Progress' and setting it's color.
                    self.executor_mail_communication_color_get.set(self.color[2])
                    self.executor_mail_communication_status.set(" In Progress ")

                    # Checking if the workbook for the MPBN Planning Sheet is selected or not
                    if (len(self.file_browser_file) == 0):
                        # Raising the Exception for file not being selected.
                        raise FileNotSelected(
                            " Please Select the MPBN Planning Workbook first!", "File Not Selected")

                    else:
                        # Calling the method of the module for circle email automation from the MPBN Planning sheet workbook and getting the return value of the 
                        # status of the Task in status flag.

                        self.executor_mail_communication_status_flag = circle_reply_task.circle_reply_task(
                            self.sender, self.file_browser_file)
                        
                        print(self.executor_mail_communication_status_flag)

                        # Checking if the status of the task is successful or not.
                        if (self.executor_mail_communication_status_flag == "Successful"):
                            # Setting the label for task to successful.
                            self.executor_mail_communication_status.set(
                                " Successful ")
                            
                            # Setting the color of the Successful label
                            self.executor_mail_communication_color_get.set(
                                self.color[1])

                            # Setting the status checker flag of the task to 1 indicating that this task has been successfully created
                            # and need not to run this task again.
                            self.executor_mail_communication_status_checker_flag = 1
                            self.task_running = 0
                            self.task_module_running = ""

                        # If the status flag is Unsuccessful then the label for the task is set to Unsuccessful and it's color is set red.
                        if (self.executor_mail_communication_status_flag == "Unsuccessful"):
                            self.executor_mail_communication_status.set(
                                " Unsuccessful ")
                            self.executor_mail_communication_color_get.set(
                                self.color[0])
                            self.executor_mail_communication_status_checker_flag = 0
                            self.task_running = 0
                            self.task_module_running = ""
                        

                # Handling the Exception for file being not selected and setting the label to unsuccessful along with it's color.
                except FileNotSelected:
                    self.executor_mail_communication_color_get.set(self.color[0])
                    self.executor_mail_communication_status_checker_flag = 0
                    self.executor_mail_communication_status.set(" Unsuccessful ")
                    self.task_running = 0
                    self.task_module_running = ""


                # Handling any other Exception and setting the label to unsuccessful along with it's color.
                except Exception as error:
                    messagebox.showerror(" Exception Occured", f"{traceback.format_exc()}\n\n{error}")
                    self.executor_mail_communication_color_get.set(self.color[0])
                    self.executor_mail_communication_status_checker_flag = 0
                    self.executor_mail_communication_status.set(" Unsuccessful ")
                    self.task_running = 0
                    self.task_module_running = ""


            else:
                self.task_running = 0
                self.task_module_running = ""
                # Raising the Custom warning in case the task is already successfuly completed.
                raise CustomWarning("  Executor Mail Communication Task Already Successfully Completed!", " Task Already Done")

        else:
            messagebox.showwarning("    Another task is running!",f"{self.task_module_running} is already running, Please Wait Patiently!")

    # Method(Function) for submitting the User Name.
    def submit_sender_name(self, event):
        # Getting the User name from string variable
        self.sender = str(self.sender_win_entry_var.get()).strip()

        # Checking if the user name is not selected.
        if (self.sender.strip() == "Select your Name!"):
            # Raising the exception for the situation when the user name is not selected.
            raise EmptyString("Please select your name to proceed!")

        elif (self.sender.strip() == "No"):
            sys.exit(0)  # exiting the program

        else:
            # Unhiding the main GUI Window
            self.main_win.deiconify()
            # Destroying the Child Sender GUI Window.
            self.sender_win.destroy()
    
    # Method(Function) to quit the Main GUI Window.
    def main_win_quit(self, event):
        if(self.main_win.state() == "normal"):
            self.main_win.destroy()
            self.exit(0)
        else:
            sys.exit(0)

    # Method(Function) to quit the sender GUI.
    def sender_win_quit(self, event):
        sys.exit(0)

    # Method(Function) for quuitting the region-handler child GUI and setting the task status label to Unsuccessful along with red color.
    def region_handler_names_win_quit(self, event):
        self.region_handler_names_win.withdraw()
        self.interdomain_kpis_mail_communication_color_get.set(self.color[0])
        self.interdomain_kpis_mail_communication_status.set(' Unsuccessful ')
        self.main_win_flag = 0
        self.main_win.deiconify()
        self.region_handler_names_win.destroy()
        self.task_running = 0
        self.task_module_running = ""

# Main Method(Function)
def main():
    # Creating an object of Tkinter
    root = Tk()
    try:
        # Creating an object of our application class and passing the Tkinter object to it.
        app = App(root)

    # Handling exceptions for empty string entry.
    except EmptyString as e:
        current_file = __file__  # gets the value of current running file
        subprocess.run(["python", current_file])
        sys.exit(0)

    # Handling exceptions for Inputs containing Integer value
    except ContainsInteger:
        current_file = __file__  # gets the value of current running file
        subprocess.run(["python", current_file])
        sys.exit(0)

    # Handling any other Exception.
    except Exception as e:
        import traceback
        messagebox.showerror("  Exception Occured", f"{traceback.format_exc()}\n\n{e}")

    root.mainloop()


if __name__ == "__main__":
    main()
