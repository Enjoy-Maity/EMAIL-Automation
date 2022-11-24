import win32com.client as win32
import pandas as pd
import openpyxl
import numpy as np
from datetime import datetime, timedelta
from tkinter import messagebox

class CustomException(Exception):
    def __init__(self,title,msg):
        self.msg = msg
        self.title = title
        super().__init__(self.title, self.msg)
        messagebox.showerror(self.title,self.msg)

def dfizer(body):
    if (isinstance(body,list)):
        body = body[0]
        columns = body.columns
        
        if (np.dtype(columns) == "int64"):
            new_body = pd.DataFrame(body.values[1:],columns = body.iloc[0])
            del body
            del columns
            return new_body
        
        else:
            return body
    elif (isinstance(body,pd.DataFrame) is True):
        columns = body.columns
        
        if (np.dtype(columns) == "int64"):
            new_body = pd.DataFrame(body.values[1:],columns=body.iloc[0])
            del body
            del columns
            return new_body
        
        else:
            return body

def mpbn_signoff(workbook):
    try:
        if(len(workbook) == 0):
            raise CustomException("")
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")    # MAPI is an API for messaging to do functions like fetching, and manipulation of mails in outlook
        inbox = outlook.GetDefaultFolder(6)
        flag_for_search = 0
        today = datetime.now()
        today = today.strftime("%m/%d/%Y")
        messages = inbox.Items
        subject_we_are_looking_for = f"MPBN Activity Validation Sign Off : {today}"
        subject_we_are_looking_for = subject_we_are_looking_for.lower()
        mail_bodies = pd.DataFrame()
        
        for message in messages:
            mail_subject = message.Subject
            
            if mail_subject.lower().__contains__(subject_we_are_looking_for):
                flag_for_search = 1
                body = message.HTMLBody
                body = pd.read_html(body)
                body = dfizer(body)
                mail_bodies =  pd.concat([mail_bodies,body],ignore_index=True)

        if (flag_for_search == 0):
            messages = inbox.Folders["Team Mails"].Items

            for message in messages:
                mail_subject = message.Subject
                
                if (mail_subject.lower().__contains__(subject_we_are_looking_for)):
                    flag_for_search = 1
                    body = message.HTMLBody
                    body = pd.read_html(body)
                    body = dfizer(body)
                    mail_bodies =  pd.concat([mail_bodies,body],ignore_index=True)
        
        if (len(mail_bodies) == 0):
            raise CustomException(" Sign-Off Mails not Found"," User sign off mails not present in mailbox.")
            
        
        if (flag_for_search == 1):
            #workbook = r"C:\Daily\MPBN Daily Planning Sheet.xlsx"
            workbook = pd.ExcelFile(workbook)
            worksheets_in_workbook = workbook.sheet_names
            required_worksheet = ""
            for worksheet in worksheets_in_workbook:
                if (worksheet == "Email-Package"):
                    if (len(worksheet) > 0):
                        required_worksheet = worksheet
                        break
            daily_planning_email_package_sheet = pd.read_excel(workbook, required_worksheet)
            CR_list = daily_planning_email_package_sheet['CR NO'].to_list()
            CR_Set = set(CR_list)

            mail_bodies_CR_List = mail_bodies['CR No'].to_list()
            mail_bodies_CR_Set = set(mail_bodies_CR_List)

            remainder = list(CR_Set - mail_bodies_CR_Set)
            extra_remainder = list(mail_bodies_CR_Set - CR_Set)
            remainder = remainder.sort()

            change_responsible = dict()
            #change_responsible_changed = []
            for cr in remainder:
                df = daily_planning_email_package_sheet[daily_planning_email_package_sheet["CR NO"] == cr]
                if (df["Change Responsible"] not in change_responsible.keys()):
                    change_responsible[df['Change Responsible']] = []
                    change_responsible[df['Change Responsible']].append(cr)
                else:
                    change_responsible[df['Change Responsible']].append(cr)
            cr_index_dict = dict()
            for i in range(0,len(daily_planning_email_package_sheet)):
                cr_index_dict[daily_planning_email_package_sheet.iloc[i]['CR NO']] = i
            
            for i in range(0,len(mail_bodies)):
                idx = cr_index_dict[mail_bodies.iloc[i]["CR No"]]
                daily_planning_email_package_sheet.at[idx,"Change Responsible"] = mail_bodies.at[i,"Change Responsible"]
                daily_planning_email_package_sheet.at[idx,"Final Status"] = mail_bodies.at[i,"Final Status - Completed / Cancelled / Rollback"]
                daily_planning_email_package_sheet.at[idx,"Reason For Rollback / Cancel"] =mail_bodies.at[]
                daily_planning_email_package_sheet.at[idx,"KPI status"] = mail_bodies.at[]
                daily_planning_email_package_sheet.at[idx,"MOP View Status"] = mail_bodies.at[]
                daily_planning_email_package_sheet.at[idx,"Interdomain KPI status"] = mail_bodies.at[]
                daily_planning_email_package_sheet.at[idx,"Second Level Validation Status"] = mail_bodies.at[]

                # fields required are change responsible, final status, reason for rollback, Interdomain KPIs, MOP Status View,

            if (len(remainder) > 0):
                if (len(extra_remainder) > 0):
                    raise CustomException(" All CRs are not Reported and Extra CR found",f"The CR's For whom the Sign Off are not reported {[{key: value for key,value in change_responsible}]}\nThe CRs which are not present are {extra_remainder}\nKindly Check")
                else:
                    raise CustomException(" All the CRs Are not Reported",f"The CR's For whom the Sign Off are not reported {[{key: value for key,value in change_responsible}]}")
            
            if (len(extra_remainder) > 0):
                raise CustomException(" Extra CR reported",f"These CR are not originally present kindly check the Email-Package and the signoff mails\nCR: {extra_remainder}")
            
            else:
                return "Successful"


    except CustomException:
        return "Unsuccessful"
    
    except Exception as e:
        messagebox.showerror("  Exception Occured",e)
        return "Unsuccessful"

#mpbn_signoff(r"C:\Daily\MPBN Daily Planning Sheet.xlsx")