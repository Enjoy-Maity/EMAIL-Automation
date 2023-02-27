import pandas as pd                         # Importing Pandas for manipulation and reading of the excel sheet
import win32com.client as win32             # Importing win32com for opening and creation of outlook mail
from tkinter import messagebox              # Importing messagebox for rasing dialogues
from datetime import datetime, timedelta    # Importing datetime to manipulate time related variables and getting today's maintenance date
import numpy as np
import sys


# Creating Custom Exception inheriting base default Exception class for raising, handling and custom exceptions.
class CustomException(Exception):
    def __init__(self,title,message):
        self.title = title
        self.message = message
        super().__init__(title,message)
        messagebox.showerror(self.title, self.message)

def email_parser(body):
    new_body_list = body.splitlines()

    result = [[]]
    to     = []
    cc     = []
    
    for i in range(0,len(new_body_list)):
        new_body_list[i] = new_body_list[i].strip()

        if (new_body_list[i].startswith("To")):
            to = new_body_list[i].split(":")[1].split(">;")

        if (new_body_list[i].startswith("Cc")):
            cc = new_body_list[i].split(":")[1].split(">;")
        
        if(new_body_list[i].startswith("Subject")):
            break

    for i in range(0,len(to)):
        to[i] = to[i].split("<")[1]

    for i in range(0,len(cc)):
        cc[i] = cc[i].split("<")[1]

    result = [to,cc]
    
    
    del to
    del cc
    del i 
    del new_body_list

    return result
    

# Mail checker and send
def mail_checker_and_sender(today_maintenance_date,sender,required_worksheet,unique_circles,dictionary_for_change_responsible_to_mail_id):
    try:
        # Creating an COM object of Microsoft Office Client Suite (Outlook) through win32com.client module.
        outlook     = win32.Dispatch("Outlook.Application")
        mapi        = outlook.GetNamespace("MAPI")              # MAPI is an API for messaging to do functions like fetching, and manipulation of mails in outlook

        # Getting the inbox folder from the outlook.
        inbox       = mapi.GetDefaultFolder(6)
        
        mail_body = "<html>\
                        <body>\
                            <div>\
                                <p>Hi Team,<br><br>Please find the executor details for below mention CRs.</p>\
                                <p>Please share below mention details with an executor for smooth execution.</p>\
                                <p>1) Circle SPOC details for end to end coordination and confirmation.<br>\
                                2) Tester details for impacted node service testing pre & post activity.<br>\
                                3) 3PP resource detail (If required).<br></p>\
                            </div>\
                            <div>\
                                {}\
                                    <br><br>\
                            </div>\
                            <div>\
                                <p>Thanks & Regards,<br>{}<br>SRF MPBN | SDU Bharti<br>Ericsson India Global Services Pvt. Ltd.</p>\
                            </div>\
                        </body>\
                    </html>"
            

        # Getting all the mails present in the inbox folder.
        inbox_messages    = inbox.Items

        # Getting today's date
        today       = datetime.now()

        # Formatting today's date so that we can compare it with the received date and time of the messages in the inbox
        today       = today.replace(hour = 10, minute = 0, second = 0).strftime('%Y-%m-%d %H:%M %p')
        
        # Filtering messages from the messages.
        inbox_messages    = inbox_messages.Restrict("[ReceivedTime] >='"+today+"'")

        circle_mail_not_found = []
        new_unique_circles = unique_circles
        # Iterating through the unique circles for checking if the mails for the circle are found or not.
        for cir in unique_circles:
            # Making the subject for finding in the inbox
            subject_we_are_looking_for = f"RE: Connected End Nodes and their services on MPBN devices: {cir}_{today_maintenance_date.strftime('%d-%m-%Y')}"

            messages    = inbox_messages.Restrict(f"@SQL=urn:schemas:httpmail:subject like '%{subject_we_are_looking_for}%'")

            # Creating a flag variable for searching the mail.
            flag_variable = 0

            if(flag_variable == 0):
                if(messages):
                    flag_variable = 1
                
            if(flag_variable == 0):
                folders = inbox.Folders
                del messages

                # Checking if the there are subfolders in the inbox.
                if(len(folders) > 0):
                    for i in range(0,len(folders)):
                        messages = inbox.Folders[i].Items
                        
                        # Filtering messages from the messages.
                        messages    = messages.Restrict("[ReceivedTime] >='"+today+"'")
                        messages    = messages.Restrict(f"@SQL=urn:schemas:httpmail:subject like '%{subject_we_are_looking_for}%'")

                        if(messages):
                            flag_variable = 1
                            break
                
            if(flag_variable == 0):
                messages = inbox_messages.Restrict(f"[ReceivedTime] >='{today}'")
                messages.Sort("[ReceivedTime]",True)

                for message in messages:
                    if(message.Subject == subject_we_are_looking_for):
                        flag_variable = 1
                        break
                del messages

            if(flag_variable == 0):
                folders = inbox.Folders

                if(folders):
                    for i in range(0,len(folders)):
                        messages = inbox.Folders[i].Items
                        
                        # Filtering messages from the messages.
                        messages    = messages.Restrict("[ReceivedTime] >='"+today+"'")
                        messages.Sort(f"[ReceivedTime] >='{today}'")
                        
                        for message in messages:
                            if(message.Subject == subject_we_are_looking_for):
                                flag_variable = 1
                                break
            
            if (flag_variable == 0):
                new_unique_circles = np.delete(new_unique_circles,np.where(new_unique_circles == cir))
                circle_mail_not_found.append(cir)

        
        # Iterating through the unique circles for replying to circles.
        for cir in new_unique_circles:
            # Making the subject for finding in the inbox
            subject_we_are_looking_for = f"RE: Connected End Nodes and their services on MPBN devices: {cir}_{today_maintenance_date.strftime('%d-%m-%Y')}"

            messages    = inbox_messages.Restrict(f"@SQL=urn:schemas:httpmail:subject like '%{subject_we_are_looking_for}%'")

            messages.Sort("[ReceivedTime]",True)

            # Creating a flag variable for searching the mail.
            flag_variable = 0

        
            # Filtering out rows based on circle
            temp_df = required_worksheet[required_worksheet["Circle"] == cir]

            # Filtering out data for just required columns 
            temp_df = temp_df[["Execution Date","Maintenance Window","CR NO","Activity Title","Risk","Location","Circle","No. of Node Involved","Change Responsible"]]

            # Formatting the execution date of the temp_df dataframe
            temp_df['Execution Date'] = pd.to_datetime(temp_df['Execution Date'], format = "%m/%d/%Y")
            temp_df['Execution Date'] = temp_df['Execution Date'].dt.strftime("%d-%b-%Y")

            # Changing the format of the dataframe containing relevant data to be presented in a more presentable manner through the usage of inline CSS.
            temp_df = temp_df.style.set_table_styles([
                {'selector':'th','props':'border:1px solid black; border-collapse : collapse; color:white;padding: 10px; background-color:rgb(0, 51, 204);text-align:center;'},
                {'selector':'tr','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
                {'selector':'td','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
                {'selector':'tr:nth-child(even)','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'}])

            temp_df = temp_df.hide(axis='index') # hiding the index coloumn


            # Creating a to variable for sending mails to
            to = ''

            # Iterating through the temp_df to attach to the to variable.
            for i in range(0,len(temp_df)):
                to = f"{dictionary_for_change_responsible_to_mail_id[temp_df.iloc[i]['Change Responsible']]};{to}"
            
            if(flag_variable == 0):
                if(messages):
                    flag_variable = 1
                    mail        = messages.GetFirst().ReplyAll()
                    result          = email_parser(mail.Body)
                    Body            = mail_body.format(temp_df.to_html(index = False), sender)
                    mail.HTMLBody   = Body + mail.HTMLBody
                    mail.To         = f"{to};{';'.join(result[0])};"
                    mail.CC         = f"{';'.join(result[1])};"
                    mail.Save()
                    mail.Display()
                    #mail.Send()

            if(flag_variable == 0):
                folders = inbox.Folders
                del messages

                # Checking if the there are subfolders in the inbox.
                if(len(folders) > 0):
                    for i in range(0,len(folders)):
                        messages = inbox.Folders[i].Items
                        
                        # Filtering messages from the messages.
                        messages    = messages.Restrict("[ReceivedTime] >='"+today+"'")
                        messages    = messages.Restrict(f"@SQL=urn:schemas:httpmail:subject like '%{subject_we_are_looking_for}%'")
                        messages.Sort("[ReceivedTime]",True)

                        if(messages):
                            flag_variable = 1
                            mail        = messages.GetFirst().ReplyAll()
                            result          = email_parser(mail.Body)
                            Body            = mail_body.format(temp_df.to_html(index = False), sender)
                            mail.HTMLBody   = Body + mail.HTMLBody
                            mail.To         = f"{to};{';'.join(result[0])};"
                            mail.CC         = f"{';'.join(result[1])};"
                            mail.Save()
                            mail.Display()
                            #mail.Send()
                            break
            
            if(flag_variable == 0):
                if(flag_variable == 0):
                    messages = inbox_messages.Restrict(f"[ReceivedTime] >='{today}'")
                    messages.Sort("[ReceivedTime]",True)

                    for message in messages:
                        if(message.Subject == subject_we_are_looking_for):
                            flag_variable = 1
                            mail        = message.ReplyAll()
                            result          = email_parser(mail.Body)
                            Body            = mail_body.format(temp_df.to_html(index = False), sender)
                            mail.HTMLBody   = Body + mail.HTMLBody
                            mail.To         = f"{to};{';'.join(result[0])};"
                            mail.CC         = f"{';'.join(result[1])};"
                            mail.Save()
                            mail.Display()
                            #mail.Send()
                            break

                    del messages

                if(flag_variable == 0):
                    folders = inbox.Folders

                    if(folders):
                        for i in range(0,len(folders)):
                            messages = inbox.Folders[i].Items
                            
                            # Filtering messages from the messages.
                            messages    = messages.Restrict("[ReceivedTime] >='"+today+"'")
                            messages.Sort("[ReceivedTime]",True)
                            
                            for message in messages:
                                if(message.Subject == subject_we_are_looking_for):
                                    flag_variable = 1
                                    mail        = message.ReplyAll()
                                    result          = email_parser(mail.Body)
                                    Body            = mail_body.format(temp_df.to_html(index = False), sender)
                                    mail.HTMLBody   = Body + mail.HTMLBody
                                    mail.To         = f"{to};{';'.join(result[0])};"
                                    mail.CC         = f"{';'.join(result[1])};"
                                    mail.Save()
                                    mail.Display()
                                    #mail.Send()
                                    break
        
        if(len(circle_mail_not_found) == 0):
            messagebox.showinfo("   Mails Displayed","Mails for all the circles displayed!")
            
            # Removing all local variables in the current scope
            objects = dir()
            for object in objects:
                if not object.startswith("__"):
                    del object

            return "Successful"
        
        if(len(circle_mail_not_found)):
            messagebox.showwarning("    Mails Displayed",f"Mails for circles other than the given below circles displayed:\n{', '.join(circle_mail_not_found)}")

            # Removing all local variables in the current scope
            objects = dir()
            for object in objects:
                if not object.startswith("__"):
                    del object
            return "Unsuccessful"

        
    except Exception as error:
        messagebox.showerror("  Exception Occured!",error)
        
        # Removing all local variables in the current scope
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object

        return "Unsuccessful"

# Main Driver Method(Function)
def circle_reply_task(sender, workbook):
    try:
        # Reading workbook in pandas
        workbook_read_in_pandas = pd.ExcelFile(workbook)
        required_worksheet = pd.read_excel(workbook_read_in_pandas,"Email-Package")
        mail_id_sheet      = pd.read_excel(workbook_read_in_pandas,"Mail Id")

        # Checking the dataframe read in pandas is empty or not
        if (len(required_worksheet) == 0):
            raise CustomException(" Worksheet Empty or Not Present!","Email-Package worksheet is not present in the workbook, Kindly Check!")
        
        if(len(mail_id_sheet) == 0):
            raise CustomException(" Worksheet Empty or Not Present!","Mail ID worksheet is not present in the workbook, Kindly Check!")

        # Creating a variable to get today's maintenance date
        today_maintenance_date = datetime.now() + timedelta(1)
        today_maintenance_date = today_maintenance_date.strftime("%m/%d/%Y")

        # Changing the datatype of the Execution date column of the dataframe to pandas datetime datatype for comparision for today's maintenance date
        required_worksheet['Execution Date'] = pd.to_datetime(required_worksheet["Execution Date"], format = "%m/%d/%Y")
        required_worksheet["Execution Date"] = required_worksheet["Execution Date"].dt.strftime("%m/%d/%Y")

        # Creating an empty list for registering rows with wrong execution(maintenance) date
        sr_for_row_with_wrong_maintenance_date = []

        # Iterating through the dataframe for checking the execution date for each row to be exactly today's maintenance date
        for i in range(0,len(required_worksheet)):
            if (required_worksheet.iloc[i]['Execution Date'] != today_maintenance_date):
                sr_for_row_with_wrong_maintenance_date.append(required_worksheet.iloc[i]['S.NO'])

        
        # Checking whether there are rows with different maintenance date if there're such rows then raising exception.
        if(len(sr_for_row_with_wrong_maintenance_date) > 0):
            raise CustomException(" Data with Wrong Maintenance Date",f"Data with other Execution Date detected for the below S.NO, kindly check!:\n{', '.join(str(num) for num in sr_for_row_with_wrong_maintenance_date)}")

        # Filtering data for the today's maintenance date
        required_worksheet = required_worksheet[required_worksheet["Execution Date"] == today_maintenance_date]

        # Checking whether data present in the required sheet is of today's maintenance date
        if(len(required_worksheet) == 0):
            raise CustomException(" Data for Today's Maintenance Date Absent","Kindly Check! Data for today's maintenance data is not preset")
        
        else:
            # Filtering the data according to user
            required_worksheet = required_worksheet[required_worksheet["Technical Validator"] == sender]

            # Checking if the technical validator is present in the sheet or not.
            if(len(required_worksheet) == 0):
                raise CustomException(" Technical Validator Not Found!",f"'{sender}' is not found as Technical Validator in the Email Package Sheet of the selected workbook, Kindly check and try again!")
            
            else:
                # Getting Unique Circles from the required sheet in a list.
                unique_circles = required_worksheet['Circle'].unique()

                # Creating dictionary for change_responsible to mail ID
                dictionary_for_change_responsible_to_mail_id = dict(zip(mail_id_sheet['Change Responsible'],mail_id_sheet['Mail ID']))

                # Creating a variable to get today's maintenance date
                today_maintenance_date = datetime.now() + timedelta(1)

                # Calling the Method (function) for replying the mail.
                flag = mail_checker_and_sender(today_maintenance_date,sender,required_worksheet,unique_circles)
                
                # Deleting all the variables before returning the value for "Successful"
                # dir() gives the list of local variables.
                objects = dir()
                for object in objects:
                    if not object.startswith("__"):
                        del object

                return flag
    
    except CustomException:
        # Deleting all the variables before returning the value for "Unsuccessful"
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object

        return "Unsuccessful"

    except ValueError as error:
        messagebox.showerror("  Exception Occured!",error)
        
        # Deleting all the variables before returning the value for "Unsuccessful"
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
        return "Unsuccessful"
    
    except RecursionError:
        messagebox.showinfo("   Recursion Error","The Program is stuck inside an Infinite loop!")
        
        # Deleting all the variables before returning the value for "Unsuccessful"
        # dir() gives the list of local variables.
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object

        return "Unsuccessful"
    
    except RuntimeError as error:
        messagebox.showerror("  Exception Occured!",error)
        
        # Deleting all the variables before returning the value for "Unsuccessful"
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
        return "Unsuccessful"
    
    except Exception as error:
        messagebox.showerror("  Exception Occured!",error)
        
        # Deleting all the variables before returning the value for "Unsuccessful"
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
        return "Unsuccessful"

#circle_reply_task("Arka Maiti",r"C:\Users\emaienj\Downloads\MPBN Daily Planning Sheet - Copy.xlsx")