# from threading import Thread
import tkinter as tk
from tkinter import *
from tkinter import filedialog,messagebox
import tkinter.ttk as ttk
from PIL import ImageTk,Image
import sys
import subprocess
# import re
import pandas as pd


# class ThreadWithReturnValue(Thread):
#     def __init__(self):
#         #super().__init__(group=None,target=None,name=None,*args,**kwargs)
#         super().__init__(self)
#         self._returnvalue = None
    
#     def run(self):
#             if self._target is not None:
#                 # self.target = self._target
#                 # self.args = self._args
#                 self._returnvalue = self._target(*self._args,**self._kwargs)
#                 print("Thread is woking")
#                 return self._returnvalue
    
#     # def join(self,*args):
#     #     Thread.join(self,*args)
#     #     return self._returnvalue

class EmptyString (Exception):
    def __init__(self,msg):
        self.msg = msg
        super().__init__(self.msg)
        messagebox.showerror(" Empty String Not Allowed",self.msg)

class RegionHandlerException(Exception):
    def __init__(self,msg):
        self.msg = msg
        super().__init__(self.msg)
        messagebox.showerror(" Exception Occurred",self.msg)


class EveningTaskException(Exception):
    def __init__(self,msg):
        self.msg = msg
        super().__init__(self.msg)
        messagebox.showerror("  Exception Occurred",self.msg)

class ContainsInteger(Exception):
    def __init__(self,msg):
        self.msg = msg
        messagebox.showerror("  Integer Not Allowed",self.msg)

class FileNotSelected(Exception):
    def __init__(self,msg,title):
        self.msg = msg
        self.title = title
        super().__init__(self.msg,self.title)
        messagebox.showerror(self.title,self.msg)

class CustomException(Exception):
    def __init__(self,msg,title):
        self.msg = msg
        self.title = title
        super().__init__(self.msg,self.title)
        messagebox.showerror(self.title,self.msg)

class CustomWarning(Exception):
    def __init__(self,msg,title):
        self.msg = msg
        self.title = title
        super().__init__(self.msg,self.title)
        messagebox.showwarning(self.title,self.msg)
        

class App(tk.Tk):
    def __init__(self,main_win):
        self.empty_string_list = []
        self.integer_string_list = []
        self.main_win = main_win
        self.main_win_flag=-1
        self.main_win.geometry("1080x701")
        self.main_win.maxsize(1080,701)
        self.main_win.minsize(1080,701)
        self.main_win.iconbitmap("images/ericsson-blue-icon-logo.ico")
        self.main_win.title("   MPBN Planning Task")
        if self.main_win_flag  == 0:
            self.update(self.frame_idx.get())
        else:
            self.style = ttk.Style()
            

            self.style.theme_use("vista")
            self.style.theme_settings("vista",{
                "TButton" : {
                    "configure":{"padding":2,
                    "font": "Ericsson_Hilda 10 bold"},
        
                    }
                }
            )
            
            self.main_win.bind("<Return>",self.file_browser_func)
            self.main_win.bind("<Escape>",self.main_win_quit)
            
            self.main_win_background_img = ImageTk.PhotoImage(Image.open("images/MPBN PLANNING TASK_3_1.png"))
            self.main_win_canvas = Canvas(self.main_win,width = 1082,height = 701,bd = 0,highlightthickness = 0,relief = "ridge")
            self.main_win_canvas.grid(row = 0,column = 0,sticky = NW)
            
            
            self.frames = [PhotoImage(file = "images\AI-transparent-automation-III.gif",format = "gif -index %i" %(i)) for i in range(31)]
            
            self.main_win_canvas.create_image(0,0,image = self.main_win_background_img,anchor = "nw")
            self.file_browser_file = ""
            self.file_browser_btn = ttk.Button(self.main_win,text = "Browse",command = lambda:self.file_browser_func(1))
            self.file_browser_entry = ttk.Entry(self.main_win,width = 40,font = ("Ericsson Hilda",13))
            
            # self.circle_email_automation_task_btn_thread = Thread(target = self.circle_email_automation_task_func,args = ())
            # self.circle_email_automation_task_btn_thread.daemon = True
            # self.circle_email_automation_task_btn = ttk.Button(self.main_win,text = "Circle Email Automation Task",command = self.circle_email_automation_task_btn_thread.start())
            # self.circle_email_automation_task_btn_thread.join()
            # self.circle_email_automation_task_status = " Successful "
            self.circle_email_automation_task_btn = ttk.Button(self.main_win,text = "Circle Email Automation Task",command = self.circle_email_automation_task_func)

            self.interdomain_kpis_data_prep_btn = ttk.Button(self.main_win,text = "Interdomain KPIs Data Preparation",command = self.interdomain_kpis_data_prep_func)

            self.interdomain_kpis_mail_communication_btn = ttk.Button(self.main_win,text = "Interdomain KPIs Mail Communication",command = lambda: self.interdomain_kpis_mail_communication_func(1))

            self.evening_task_btn = ttk.Button(self.main_win,text = "      Evening Message Task     ",command = lambda: self.evening_task_func(1))

            # self.main_win_canvas.create_window(840,182,anchor = "nw",window = self.file_browser_btn)
            # self.main_win_canvas.create_window(420,182,anchor = "nw",window = self.file_browser_entry)
            # self.main_win_canvas.create_window(85,300,anchor = "nw",window = self.circle_email_automation_task_btn)
            # self.main_win_canvas.create_window(308,300,anchor = "nw",window = self.interdomain_kpis_data_prep_btn)
            # self.main_win_canvas.create_window(565,300,anchor = "nw",window = self.interdomain_kpis_mail_communication_btn)
            # self.main_win_canvas.create_text(120,192,text = "Choose The File",fill = "#FFFFFF",font = ("Ericsson Hilda",18,"bold"))

            ##################################################### Status setter variables #####################################################################################################################
            self.circle_email_automation_task_status = StringVar(self.main_win_canvas)
            self.circle_email_automation_task_status.set("")
            
            self.interdomain_kpis_data_prep_task_status = StringVar(self.main_win_canvas)
            self.interdomain_kpis_data_prep_task_status.set("")
            self.interdomain_kpis_data_prep_task_completed =0

            self.interdomain_kpis_mail_communication_status = StringVar(self.main_win_canvas)
            self.interdomain_kpis_mail_communication_status.set("")

            self.evening_task_status = StringVar(self.main_win_canvas)
            self.evening_task_status.set("")
            
            self.color = ["#FF0000","#00FF00","#FFFFFF"]

            self.circle_email_automation_task_color_get = StringVar()
            self.circle_email_automation_task_color_get.set("")

            self.interdomain_kpis_data_prep_color_get = StringVar()
            self.interdomain_kpis_data_prep_color_get.set("")

            self.interdomain_kpis_mail_communication_color_get = StringVar()
            self.interdomain_kpis_mail_communication_color_get.set("")

            self.evening_task_color_get = StringVar()
            self.evening_task_color_get.set("")

        ########################################################## Status checker flags #######################################################################################################################
            self.circle_email_automation_status_checker_flag = 0
            self.interdomain_kpis_data_prep_status_checker_flag = 0
            self.interdomain_kpis_mail_communication_status_checker_flag = 0
            self.evening_task_status_checker_flag = 0
            self.update(1)
        self.get_sender_name()

        
       
    def get_sender_name (self):
        self.sender_win = Toplevel(self.main_win)
        self.main_win.withdraw()
        self.sender_win.title(" Please Enter Your Name To Proceed")
        self.sender_win.iconbitmap("images/ericsson-blue-icon-logo.ico")
        self.sender_win.geometry("600x150")
        self.sender_win.minsize(600,150)
        self.sender_win.maxsize(600,150)
        self.sender_win.bind("<Escape>",self.sender_win_quit)
        self.sender_win.bind("<Return>",self.submit_sender_name)
        self.sender_win.protocol("WM_DELETE_WINDOW",lambda:self.sender_win_quit(1))
        self.sender_win.grab_set()
        
        self.sender_win_background = ImageTk.PhotoImage(Image.open("images/MPBN PLANNING TASK_3_2.png"))
        self.sender_win_canvas = Canvas(self.sender_win,width = 600,height = 150,bd = 0,highlightthickness = 0,relief = "ridge")
        self.sender_win_canvas.grid(row = 0,column = 0,sticky = NW)

        self.sender_win_entry = ttk.Entry(self.sender_win,width = 48,font = ("Ericsson Hilda",12))
        self.sender_win_entry.focus_force()
        self.sender_win_btn = ttk.Button(self.sender_win,text = "Submit",command = lambda:self.submit_sender_name(1))

        self.sender_win_canvas.create_image(0,0,image = self.sender_win_background,anchor = "nw")
        self.sender_win_canvas.create_text(295,30,text = "Please Enter Your Name To Proceed or N/n To Exit",fill = "#FFFFFF",font = ("Ericsson Hilda",19,"bold"))
        self.sender_win_canvas.create_window(20,70,anchor = "nw",window = self.sender_win_entry)
        self.sender_win_canvas.create_window(522,82,window = self.sender_win_btn)
        
        self.sender_win.mainloop()


    def update(self,ind):
        self.frame = self.frames[ind]
        ind  +=   1
        self.frame_idx = IntVar()
        self.frame_idx.set(ind)
        self.main_win_canvas.delete("all")
        self.main_win_canvas.create_image(0,0,image = self.main_win_background_img,anchor = "nw")
        self.main_win_canvas.create_image(870,2,image = self.frame,anchor = "nw")
        self.main_win_canvas.create_text(220,194,text = "Choose Daily Planning Sheet :-",fill = "#FFFFFF",font = ("Ericsson Hilda ExtraBold",21   ,"bold underline"))        
        self.main_win_canvas.create_window(420,182,anchor = "nw",window = self.file_browser_entry)
        self.main_win_canvas.create_window(840,180,anchor = "nw",window = self.file_browser_btn)
        
        self.main_win_canvas.create_window(100,268,anchor = "nw",window = self.circle_email_automation_task_btn)
        self.main_win_canvas.create_text(137,300,anchor = "nw",text = self.circle_email_automation_task_status.get(),fill = self.circle_email_automation_task_color_get.get(),font = ("Ericsson Hilda ExtraBold",15,"bold"))
        
        self.main_win_canvas.create_window(359,268,anchor = "nw",window = self.interdomain_kpis_data_prep_btn)
        self.main_win_canvas.create_text(406,300,anchor = "nw",text = self.interdomain_kpis_data_prep_task_status.get(),fill = self.interdomain_kpis_data_prep_color_get.get(),font = ("Ericsson Hilda ExtraBold",15,"bold"))

        self.main_win_canvas.create_window(655,268,anchor = "nw",window = self.interdomain_kpis_mail_communication_btn)
        self.main_win_canvas.create_text(710,300,anchor = "nw",text = self.interdomain_kpis_mail_communication_status.get(),fill =self.interdomain_kpis_mail_communication_color_get.get(),font = ("Ericsson Hilda ExtraBold",15,"bold"))

        self.main_win_canvas.create_window(100,388,anchor = "nw",window = self.evening_task_btn)
        self.main_win_canvas.create_text(137,420,anchor = "nw",text = self.evening_task_status.get(),fill =self.evening_task_color_get.get(),font = ("Ericsson Hilda ExtraBold",15,"bold"))

        self.main_win_canvas.update_idletasks()                     # Solves the flickering problem when the frame gets updated
        if ind  ==  31:
            ind = 1
        if self.main_win_flag != 0:
            self.main_win.after(31,self.update,ind)
        
        if self.main_win_flag == 0:
            self.main_win.after(ind,self.update,ind)

    def file_browser_func(self,event):
        self.file_browser_entry.delete(0,END)
        self.mystring = filedialog.askopenfilename(initialdir = "C:\\",title = "  Choose the worksheet",filetypes = (("Excel Files (.xlsx)","*.xlsx"),("Excel Files (.xls)","*.xls"),("All Files","*.*")))
        self.file_browser_entry.insert(0,self.mystring)
        self.file_browser_file = self.mystring


    def circle_email_automation_task_func (self):
        if (self.circle_email_automation_status_checker_flag == 0):
            try:
                self.circle_email_automation_task_color_get.set(self.color[2])
                self.circle_email_automation_task_status.set(" In Progress ")

                if (len(self.file_browser_file)==0):
                    raise FileNotSelected (" Please Select the MPBN Planning Workbbok first!","File Not Selected")

                else:
                    import circle_Email_Automation_Task
                    self.circle_email_automation_status_flag = circle_Email_Automation_Task.fetch_details(self.sender,self.file_browser_file)
                    
                    if (self.circle_email_automation_status_flag == "Successful"):
                        self.circle_email_automation_task_status.set(" Successful ")
                        self.circle_email_automation_task_color_get.set(self.color[1])
                        self.circle_email_automation_status_checker_flag = 1

                    if (self.circle_email_automation_status_flag == "Unsuccessful"):
                        self.circle_email_automation_task_status.set(" Unsuccessful ")
                        self.circle_email_automation_task_color_get.set(self.color[0])
                        self.circle_email_automation_status_checker_flag = 0

            except FileNotSelected:
                self.circle_email_automation_task_color_get.set(self.color[0])
                self.circle_email_automation_status_checker_flag = 0
                self.circle_email_automation_task_status.set(" Unsuccessful ")
            except Exception as error:
                #self.circle_email_automation_task_thread.join()
                messagebox.showerror(" Exception Occured",error)
                self.circle_email_automation_task_color_get.set(self.color[0])
                self.circle_email_automation_status_checker_flag = 0
                self.circle_email_automation_task_status.set(" Unsuccessful ")

        else :
            raise CustomWarning ("  Circle Automation Task Already Successfully Completed", " Task Already Done")    
        
    
    def interdomain_kpis_data_prep_func (self):
        if (self.interdomain_kpis_data_prep_status_checker_flag == 0):
            self.interdomain_kpis_data_prep_color_get.set(self.color[2])
            self.interdomain_kpis_data_prep_task_status.set(" In Progress ")
            
            try:
                #time.sleep(5)
                if (len(self.file_browser_file)==0):
                    raise FileNotSelected (" Please Select the MPBN Planning Excel Workbook first!","File Not Selected")
                
                else:
                    import interdomain_KPIs_Data_Prep_Task
                    # self.thread=ThreadWithReturnValue(interdomain_KPIs_Data_Prep_Task.paco_cscore(self.sender,self.file_browser_file))
                    # self.thread.daemon = True
                    # self.thread.start()
                    # print(self.thread._returnvalue)
                    self.interdomain_kpis_data_prep_status_flag = interdomain_KPIs_Data_Prep_Task.paco_cscore(self.sender,self.file_browser_file)
                    # self.thread.join()
                    
                    #interdomain_KPIs_Data_Prep_Task.paco_cscore(self.sender,self.file_browser_file)

                    if (self.interdomain_kpis_data_prep_status_flag == 'Successful'):
                        self.interdomain_kpis_data_prep_color_get.set(self.color[1])
                        self.interdomain_kpis_data_prep_task_completed = 1
                        self.interdomain_kpis_data_prep_status_checker_flag = 1
                        self.interdomain_kpis_data_prep_task_status.set(" Successful ")
                    
                    elif (self.interdomain_kpis_data_prep_status_flag == 'Unsuccessful'):
                        self.interdomain_kpis_data_prep_color_get.set(self.color[0])
                        self.interdomain_kpis_data_prep_task_completed = 0
                        self.interdomain_kpis_data_prep_status_checker_flag = 0
                        self.interdomain_kpis_data_prep_task_status.set(" Unsuccessful ")
                    
            
            except FileNotSelected:
                self.interdomain_kpis_data_prep_color_get.set(self.color[0])
                self.interdomain_kpis_data_prep_status_checker_flag = 0
                self.interdomain_kpis_data_prep_task_status.set(" Unsuccessful ")
            #  except Exception as error:
                #messagebox.showerror(" Exception Occured", error)
                # self.interdomain_kpis_data_prep_color_get.set(self.color[0])
                # self.interdomain_kpis_data_prep_status_checker_flag = 0
                # self.interdomain_kpis_data_prep_task_status.set(" Unsuccessful ")
        else:
            raise CustomWarning (" Interdomain KPIs Data Prep Task Already Successfully Completed"," Task Already Done")
    
    def interdomain_kpis_mail_communication_func (self,event):
        if (self.interdomain_kpis_mail_communication_status_checker_flag == 0):
            self.region_handler_names_win = Toplevel(self.main_win)
            
            if self.main_win.state() == "normal":
                self.main_win.withdraw()
            
            if (self.interdomain_kpis_data_prep_task_completed == 1):

                self.region_handler_names_win.geometry("450x300")
                self.region_handler_names_win.minsize(450,300)
                self.region_handler_names_win.maxsize(450,300)
                self.region_handler_names_win.iconbitmap("images/ericsson-blue-icon-logo.ico")
                self.region_handler_names_win.title("   Names for (PAN INDIA) MPBN Planning SPOC's ")
                self.region_handler_names_win.bind("<Escape>",self.region_handler_names_win_quit)
                #self.region_handler_names_win.protocol("WM_DELETE_WINDOW",lambda:self.region_handler_names_win_quit(1))

                

                self.region_handler_names_win_background = ImageTk.PhotoImage(Image.open("images/MPBN PLANNING TASK_3_3.png"))
                self.region_handler_names_win_canvas = Canvas(self.region_handler_names_win,width = 450,height = 300,bd = 0,highlightthickness = 0,relief = "ridge")
                self.region_handler_names_win_canvas.grid(row = 0,column = 0,sticky = NW)
                self.region_handler_names_win_canvas.create_image(0,0,image = self.region_handler_names_win_background,anchor = "nw")
                
                self.north_and_west_region_entry = ttk.Entry(self.region_handler_names_win_canvas,width = 40,font = ("Ericsson Hilda",13))
                self.region_handler_names_win_canvas.create_text(10,20,anchor = "nw",text = "Please Enter Name Of North and West Region Planner",fill = "#FFFFFF",font = ("Ericsson Hilda",13,"bold"))
                self.region_handler_names_win_canvas.create_window(10,65,anchor = "nw",window = self.north_and_west_region_entry)
                self.north_and_west_region_entry.focus_force()
                
                self.region_handler_names_win_canvas.create_text(10,120,anchor = "nw",text = "Please Enter Name Of South and East Region Planner",fill = "#FFFFFF",font = ("Ericsson Hilda",13,"bold"))
                self.east_region_and_south_region_entry  =  ttk.Entry(self.region_handler_names_win_canvas,width  =  40,font = ("Ericsson Hilda",13))
                self.region_handler_names_win_canvas.create_window(10,165,anchor = "nw",window = self.east_region_and_south_region_entry)
                
                self.north_and_west_region = self.north_and_west_region_entry.get()
                self.east_region_and_south_region = self.east_region_and_south_region_entry.get()

                self.region_handler_names_win_canvas_submit = ttk.Button(self.region_handler_names_win,text = "Submit",command = lambda:self.interdomain_kpis_mail_commmunication_starter_func(1))
                self.region_handler_names_win_canvas.create_window(380,270,anchor="se",window=self.region_handler_names_win_canvas_submit)
                self.region_handler_names_win.bind("<Return>",self.interdomain_kpis_mail_commmunication_starter_func)
                
                
                self.region_handler_names_win.protocol("WM_DELETE_WINDOW",lambda:self.region_handler_names_win_quit(1))
                
                if self.region_handler_names_win.state() != "normal":
                    if self.main_win.state() != "normal":
                        self.main_win_flag = 0
                        self.main_win.deiconify()
                    self.region_handler_names_win.destroy()
            
            else :
                self.interdomain_kpis_mail_communication_color_get.set(self.color[0])
                self.interdomain_kpis_mail_communication_status.set(' Unsuccessful ')
                self.region_handler_names_win.destroy()

                if self.main_win.state() != "normal":
                    self.main_win_flag = 0
                    self.main_win.deiconify()
                self.interdomain_kpis_mail_communication_status_checker_flag = 0
                raise CustomException ("Please! Run Interdomain KPIs Data Prep task First!","   Task Unsuccessful")
               
            self.region_handler_names_win.mainloop()
        
        else:
            self.interdomain_kpis_mail_communication_status_checker_flag = 0
            raise CustomWarning(" Interdomain KPIs mail Communication Task Already Successfully Completed"," Task Already Done")
        
    def interdomain_kpis_mail_commmunication_starter_func(self,event):
        self.east_region_and_south_region = self.east_region_and_south_region_entry.get()
        self.north_and_west_region = self.north_and_west_region_entry.get()
        self.interdomain_kpis_mail_communication_color_get.set(self.color[2])
        self.interdomain_kpis_mail_communication_status.set(" In Progress ")
        
        self.new_empty_string_list = []
        self.new_integer_string_list = []
        
        try:
            if (len(self.file_browser_file)==0):
                self.interdomain_kpis_mail_communication_status_checker_flag = 0
                raise FileNotSelected (" Please Select the MPBN Planning Excel Workbook first!","File Not Selected")
            
            if (len(self.north_and_west_region)>0) and (len(self.east_region_and_south_region)>0):
                if (not (any(c.isdigit() for c in self.north_and_west_region ))) and (not (any(c.isdigit() for c in self.east_region_and_south_region))):
                    self.main_win_flag = 0
                    self.main_win.deiconify()
                    self.region_handler_names_win.destroy()
                    import interdomain_KPIs_Mail_Comm_Task
                    interdomain_KPIs_Mail_Comm_Task.paco_cscore(self.sender,self.file_browser_file,self.north_and_west_region,self.east_region_and_south_region)
                    self.interdomain_kpis_mail_communication_color_get.set(self.color[1])
                    self.interdomain_kpis_mail_communication_status_checker_flag = 1
                    self.interdomain_kpis_mail_communication_status.set(" Successful ")
                    
                
            if (any(c.isdigit() for c in self.north_and_west_region)):
                self.new_integer_string_list.append("North & West Region Handler")
            
            if (any(c.isdigit() for c in self.east_region_and_south_region)):
                self.new_integer_string_list.append("South & East Region Handler")
            
            if (len(self.north_and_west_region) == 0):
                self.new_empty_string_list.append("North & West Region Handler")

            if (len(self.east_region_and_south_region) == 0):
                self.new_empty_string_list.append("South & East Region Handler")


            
            
            if (len(self.new_integer_string_list) > 0) and (len(self.new_empty_string_list) == 0):
                self.interdomain_kpis_mail_communication_color_get.set(self.color[0])
                self.interdomain_kpis_mail_communication_status_checker_flag = 0
                self.interdomain_kpis_mail_communication_status.set(" Unsucessful ")
                raise RegionHandlerException (f"Please Enter Valid Name/s, Fields with Numbers are not allowed \nField/Fields with Number: {','.join(self.new_integer_string_list)}")
            if (len(self.new_empty_string_list) > 0) and (len(self.new_integer_string_list) == 0):
                self.interdomain_kpis_mail_communication_color_get.set(self.color[0])
                self.interdomain_kpis_mail_communication_status_checker_flag = 0
                self.interdomain_kpis_mail_communication_status.set(" Unsucessful ")
                raise RegionHandlerException (f"Please Enter valid Name/s, Empty Strings are not allowed\nEmpty Field/Fields: {','.join(self.new_empty_string_list)}")
            if (len(self.new_empty_string_list) > 0) and (len(self.new_integer_string_list) > 0):
                self.interdomain_kpis_mail_communication_color_get.set(self.color[0])
                self.interdomain_kpis_mail_communication_status_checker_flag = 0
                self.interdomain_kpis_mail_communication_status.set(" Unsucessful ")
                raise RegionHandlerException (f"Please Enter Valid Names, Empty Strings and Numbers are not allowed \nEmpty Field: {','.join(self.new_empty_string_list)} \nField with Number: {','.join(self.new_integer_string_list)}")
            
        except FileNotSelected:
            self.interdomain_kpis_mail_communication_color_get.set(self.color[0])
            self.interdomain_kpis_mail_communication_status.set(' Unsuccessful ')

        except RegionHandlerException :
            self.new_empty_string_list = []
            self.new_integer_string_list = []
            self.north_and_west_region_entry.focus_force()
            self.interdomain_kpis_mail_communication_color_get.set(self.color[0])
            self.interdomain_kpis_mail_communication_status.set(' Unsuccessful ')
        
        except Exception as error:
            messagebox.showerror(" Exception Occured", error)
            self.interdomain_kpis_mail_communication_color_get.set(self.color[0])
            self.interdomain_kpis_mail_communication_status.set(' Unsuccessful ')
            

                

    def evening_task_func (self,event):
        if (self.evening_task_status_checker_flag == 0):
            
            if (len(self.file_browser_file) == 0):
                self.evening_task_color_get.set(self.color[0])
                self.evening_task_status_checker_flag = 0
                self.evening_task_status.set(' Unsuccessful ')
                raise FileNotSelected (" Please Select the MPBN Planning Excel Workbook first!","File Not Selected")
            
            else:
                self.interdomain_kpis_data_prep_creation_status_flag = 0
                self.workbook = pd.ExcelFile(self.file_browser_file)
                self.worksheet_names = self.workbook.sheet_names

                for sheet in self.worksheet_names:
                    if (sheet == 'Email-Package'):
                        self.worksheet = pd.read_excel(self.workbook,sheet)
                        if (len(self.worksheet) > 0):
                            self.interdomain_kpis_data_prep_creation_status_flag = 1

                if (self.interdomain_kpis_data_prep_creation_status_flag == 0):
                    raise CustomException('Kindly Click the Button for Interdomain Kpi Data Prep First!','Email-Package Worksheet Empty')
                
                else:
                    self.evening_task_win = Toplevel(self.main_win)
                    if self.main_win.state() == 'normal':
                        self.main_win.withdraw()
                    self.evening_task_win.iconbitmap('images/ericsson-blue-icon-logo.ico')
                    self.evening_task_win.title("   Please Enter The Names to Proceed")
                    self.evening_task_win.geometry("600x550")
                    self.evening_task_win.minsize(600,550)
                    self.evening_task_win.maxsize(600,550)
                    self.evening_task_win.bind("<Escape>",self.evening_task_func_quit)

                    self.evening_task_background = ImageTk.PhotoImage(Image.open("images/MPBN PLANNING TASK_3_4.png"))
                    self.evening_task_win_canvas=Canvas(self.evening_task_win,height = 550,width = 600,bd=0,highlightthickness=0, relief="ridge")
                    self.evening_task_win_canvas.grid(row = 0,column = 0,sticky = NW)
                    self.evening_task_win_canvas.create_image(0,0,image = self.evening_task_background, anchor = "nw")
                    

                    self.evening_task_win_canvas.create_text(10,20,anchor = "nw",text = "Please Enter Night Shift Lead Name", fill = "#FFFFFF",font = ("Ericsson Hilda",18,"bold"))
                    self.evening_task_win_canvas_night_shift_lead_entry = ttk.Entry(self.evening_task_win_canvas,width = 40, font = ("Ericsson Hilda",15))
                    self.evening_task_win_canvas.create_window(10,70,anchor = "nw",window=self.evening_task_win_canvas_night_shift_lead_entry)

                    self.evening_task_win_canvas.create_text(10,150,anchor = "nw",text = "Please Enter Buffer/Auditor/Trainer Name", fill = "#FFFFFF",font = ("Ericsson Hilda",18,"bold"))
                    self.evening_task_win_canvas_buffer_auditor_trainer_entry = ttk.Entry(self.evening_task_win_canvas,width = 40, font = ("Ericsson Hilda",15))
                    self.evening_task_win_canvas.create_window(10,200,anchor = "nw",window=self.evening_task_win_canvas_buffer_auditor_trainer_entry)

                    self.evening_task_win_canvas.create_text(10,280,anchor = "nw",text = "Please Enter Resource on Automation Name", fill = "#FFFFFF",font = ("Ericsson Hilda",18,"bold"))
                    self.evening_task_win_canvas_resource_on_automation_entry = ttk.Entry(self.evening_task_win_canvas,width = 40, font = ("Ericsson Hilda",15))
                    self.evening_task_win_canvas.create_window(10,310,anchor = "nw",window=self.evening_task_win_canvas_resource_on_automation_entry)

                    self.evening_task_submit_btn = ttk.Button(self.evening_task_win, text = "Submit", command = lambda: self.evening_task_func_starter(1))
                    self.evening_task_win_canvas.create_window(580,520,window = self.evening_task_submit_btn,anchor = "se")

                    self.evening_task_win_canvas_night_shift_lead_entry.focus_force()

                    self.evening_task_win.protocol("WM_DELETE_WINDOW",lambda:self.evening_task_func_quit(1))
                    self.evening_task_win.bind("<Return>",self.evening_task_func_starter)

                    if self.evening_task_win.state() != "normal" :
                        if self.main_win.state() != "normal":
                            self.main_win_flag = 0
                            self.main_win.deiconify()
                        self.evening_task_win.destroy()

                    self.evening_task_win.mainloop()
        
        else:
            raise CustomWarning ("Evening Task Already Successfully Completed"," Task Already Done")
        

    def evening_task_func_quit(self,event):
        self.evening_task_win.withdraw()
        self.evening_task_color_get.set(self.color[0])
        self.evening_task_status.set(' Unsuccessful ')
        self.main_win_flag = 0
        self.main_win.deiconify()
        self.evening_task_win.destroy()
    
    def evening_task_func_starter(self,event):
        self.night_shift_lead = self.evening_task_win_canvas_night_shift_lead_entry.get()
        self.buffer_auditor_trainer = self.evening_task_win_canvas_buffer_auditor_trainer_entry.get()
        self.resource_on_automation = self.evening_task_win_canvas_resource_on_automation_entry.get()
        self.evening_task_color_get.set(self.color[2])
        self.evening_task_status.set(' In Progress ')
        
        self.empty_string_list = []
        self.integer_string_list = []
        
        try:
            
            if (len(self.night_shift_lead)>0) and (len(self.buffer_auditor_trainer) > 0) and (len(self.resource_on_automation) > 0):
                if (not (any(c.isdigit() for c in self.night_shift_lead))) and (not (any(c.isdigit() for c in self.buffer_auditor_trainer))) and (not (any(c.isdigit() for c in self.resource_on_automation))):
                    self.main_win_flag = 0
                    self.main_win.deiconify()

                    self.evening_task_win.destroy()
                    import evening_mail_task
                    self.evening_mail_task_status_flag = evening_mail_task.evening_task(self.sender,self.night_shift_lead,self.buffer_auditor_trainer,self.resource_on_automation,self.file_browser_file)
                    
                    if (self.evening_mail_task_status_flag == 'Successful'):
                        self.evening_task_color_get.set(self.color[1])
                        self.evening_task_status_checker_flag = 1
                        self.evening_task_status.set(' Successful ')
                    
                    if (self.evening_mail_task_status_flag == 'Unsuccessful'):
                        self.evening_task_color_get.set(self.color[0])
                        self.evening_task_status_checker_flag = 0
                        self.evening_task_status.set(' Unsuccessful ')
                
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
                raise EveningTaskException (f"Please Enter Valid Names, Empty Strings are not allowed\nEmpty Field/Fields: {','.join(self.empty_string_list)}")
            if (len(self.empty_string_list) == 0) and (len(self.integer_string_list) > 0):
                self.evening_task_color_get.set(self.color[0])
                self.evening_task_status_checker_flag = 0
                self.evening_task_status.set(' Unsuccessful ')
                raise EveningTaskException (f"Please Enter Valid Names, Numbers are not allowed\nField/Fields with Numbers: {','.join(self.integer_string_list)}")
            if (len(self.empty_string_list) > 0) and (len(self.integer_string_list) > 0):
                self.evening_task_color_get.set(self.color[0])
                self.evening_task_status_checker_flag = 0
                self.evening_task_status.set(' Unsuccessful ')
                raise EveningTaskException (f"Please Enter Valid Names, Empty Strings and Numbers are not allowed\n Empty Field/Fields: {','.join(self.empty_string_list)}\nField/Fields with Numbers: {','.join(self.integer_string_list)}")

                
        except EveningTaskException:
            self.empty_string_list = []
            self.integer_string_list = []
            self.evening_task_color_get.set(self.color[0])
            self.evening_task_status_checker_flag = 0
            self.evening_task_win_canvas_night_shift_lead_entry.focus_force()
            self.evening_task_status.set(' Unsuccessful ')
        
        except Exception as error:
            messagebox.showerror(" Exception Occured",error)
            self.evening_task_color_get.set(self.color[0])
            self.evening_task_status_checker_flag = 0
            self.evening_task_status.set(' Unsuccessful ')

    
    def submit_sender_name (self,event):
        self.sender = str(self.sender_win_entry.get())
        if len(self.sender)  ==  0:
            raise EmptyString("Please enter your name not an Empty String.")
        
        elif (any(c.isdigit() for c in self.sender)):
            raise ContainsInteger("Invalid Name as it contains Integer")
        
        elif self.sender  ==  "n" or self.sender  ==  "N":
            sys.exit(0) # exiting the program
        
        else:
            self.main_win.deiconify()
            self.sender_win.destroy()

    def main_win_quit(self,event):
        sys.exit(0)

    def sender_win_quit(self,event):
        sys.exit(0)
    
    def region_handler_names_win_quit(self,event):
        self.region_handler_names_win.withdraw()
        self.interdomain_kpis_mail_communication_color_get.set(self.color[0])
        self.interdomain_kpis_mail_communication_status.set(' Unsuccessful ')
        self.main_win_flag = 0
        self.main_win.deiconify()
        self.region_handler_names_win.destroy()


def main():
    root = Tk()
    try:
        app = App(root)
    
    except EmptyString :
        current_file = __file__ # gets the value of current running file
        subprocess.run(["python", current_file])
        sys.exit(0)
    
    except ContainsInteger:
        current_file = __file__ # gets the value of current running file
        subprocess.run(["python", current_file])
        sys.exit(0)


    except Exception as e:
        messagebox.showerror("  Exception Occurred",e)
    root.mainloop()

if __name__  ==  "__main__":
    main()