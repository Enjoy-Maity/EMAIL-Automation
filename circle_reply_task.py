import pandas as pd                         # Importing Pandas for manipulation and reading of the excel sheet
import win32com.client as win32             # Importing win32com for opening and creation of outlook mail
from tkinter import messagebox              # Importing messagebox for rasing dialogues
from datetime import datetime, timedelta    # Importing datetime to manipulate time related variables and getting today's maintenance date
import numpy as np                          # Importing Numpy for numpy array operations.
import sys                                  # Importing sys module for 


# Creating Custom Exception inheriting base default Exception class for raising, handling and custom exceptions.
class CustomException(Exception):
    def __init__(self,title,message):
        self.title = title
        self.message = message
        super().__init__(title,message)
        messagebox.showerror(self.title, self.message)

def email_parser(body):
    new_body_list = body.splitlines()

    result      = [[]]
    to          = []
    cc          = []
    from_mail   = ""
    for i in range(0,len(new_body_list)):
        new_body_list[i] = new_body_list[i].strip()

        if (new_body_list[i].startswith("From:")):
            from_mail = new_body_list[i].split(":")[1].split("<")[1].strip(">")

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

    to.append(from_mail)
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
        today       = today.replace(hour = 10)
        
        # Sorting the mails
        inbox_messages.Sort("[ReceivedTime]",True)

        circle_mail_not_found = []
        new_unique_circles = unique_circles
        
        print("Entering the test for mail")

        # Iterating through the unique circles for checking if the mails for the circle are found or not.
        for cir in unique_circles:
            # Making the subject for finding in the inbox
            subject_we_are_looking_for = f"RE: Connected End Nodes and their services on MPBN devices: {cir}_{today_maintenance_date.strftime('%d-%m-%Y')}"

            # Creating a flag variable for searching the mail.
            flag_variable = 0
                
            if(flag_variable == 0):
                messages = inbox_messages

                for message in messages:
                    print(message.Subject)
                    try:
                        dt = message.ReceivedTime
                        year,month,day,hour,minute = dt.year,dt.month,dt.day,dt.hour,dt.minute
                        if(datetime(year,month,day,hour,minute) >= today):
                            if(message.Subject.lower() == subject_we_are_looking_for.lower()):
                                print(f"test:{cir}")
                                flag_variable = 1
                                break
                        else:
                            break
                    except:
                        continue

                del messages

            if(flag_variable == 0):
                folders = inbox.Folders

                if(len(folders)):
                    for i in range(0,len(folders)):
                        try:
                            messages = inbox.Folders[i].Items
                            
                            # Filtering messages from the messages.
                            messages.Sort("[ReceivedTime]",True)
                            
                            for message in messages:
                                print(message.Subject)
                                try:
                                    dt = message.ReceivedTime
                                    year,month,day,hour,minute = dt.year,dt.month,dt.day,dt.hour,dt.minute
                                    if(datetime(year,month,day,hour,minute) >= today):
                                        if(message.Subject.lower() == subject_we_are_looking_for.lower()):
                                            print(f"test:{cir}")
                                            flag_variable = 1
                                            break
                                    else:
                                        break
                                except:
                                    continue
                            
                            if (flag_variable == 0):
                                sub_folders = inbox.Folders[i].Folders

                                if(len(sub_folders)):
                                    for sub_folder in range(len(sub_folders)):
                                        try:
                                            messages = inbox.Folders[i].Folders[sub_folder].Items
                                
                                            # Filtering messages from the messages.
                                            messages.Sort("[ReceivedTime]",True)
                                            
                                            for message in messages:
                                                print(message.Subject)
                                                try:
                                                    dt = message.ReceivedTime
                                                    year,month,day,hour,minute = dt.year,dt.month,dt.day,dt.hour,dt.minute
                                                    if(datetime(year,month,day,hour,minute) >= today):
                                                        if(message.Subject.lower() == subject_we_are_looking_for.lower()):
                                                            print(f"test:{cir}")
                                                            flag_variable = 1
                                                            break
                                                    else:
                                                        break
                                                
                                                except:
                                                    continue
                                        
                                            if(flag_variable == 1):
                                                break
                                        except:
                                            continue

                            if(flag_variable == 1):
                                break

                        except:
                            continue   
            
            if (flag_variable == 0):
                new_unique_circles = np.delete(new_unique_circles,np.where(new_unique_circles == cir))
                circle_mail_not_found.append(cir)

        
        # Iterating through the unique circles for replying to circles.
        for cir in new_unique_circles:
            # Making the subject for finding in the inbox
            subject_we_are_looking_for = f"RE: Connected End Nodes and their services on MPBN devices: {cir}_{today_maintenance_date.strftime('%d-%m-%Y')}"

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
            dataframe = temp_df.style.set_table_styles([
                {'selector':'th','props':'border:1px solid black; border-collapse : collapse; color:white;padding: 10px; background-color:rgb(0, 51, 204);text-align:center;'},
                {'selector':'tr','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
                {'selector':'td','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
                {'selector':'tr:nth-child(even)','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'}])

            dataframe = dataframe.hide(axis='index') # hiding the index coloumn


            # Creating a to variable for sending mails to
            to_list = []
            to      = []
            
            # Iterating through the temp_df to attach to the to variable.
            for i in range(0,len(temp_df)):
                to_list.append(temp_df.iloc[i]['Change Responsible'])
            
            # for i in to:
            #     if(i.upper() == 'NAN'):
            #         to.remove(i)
               
            for receipient in to_list:
                to.append(dictionary_for_change_responsible_to_mail_id[receipient])

            # Converting the list to set to have only unique values.
            to = set(to)
            
            if(flag_variable == 0):
                if(flag_variable == 0):
                    messages = inbox_messages
                    messages.Sort("[ReceivedTime]",True)

                    for message in messages:
                        try:
                            dt = message.ReceivedTime
                            year,month,day,hour,minute = dt.year,dt.month,dt.day,dt.hour,dt.minute
                            if(datetime(year,month,day,hour,minute) >= today):
                                if(message.Subject.lower() == subject_we_are_looking_for.lower()):
                                    flag_variable = 1
                                    mail        = message.ReplyAll()
                                    result          = email_parser(mail.Body)
                                    Body            = mail_body.format(dataframe.to_html(index = False), sender)
                                    mail.HTMLBody   = Body + mail.HTMLBody
                                    mail.To         = f"{';'.join(to)};{';'.join(result[0])};"
                                    mail.CC         = f"{';'.join(result[1])};"
                                    mail.Save()
                                    mail.Display()
                                    #mail.Send()
                                    break
                            
                            else:
                                break
                        
                        except:
                            continue
                    
                    del messages

                if(flag_variable == 0):
                    folders = inbox.Folders

                    if(folders):
                        for i in range(0,len(folders)):
                            try:
                                messages = inbox.Folders[i].Items
                                
                                # Filtering messages from the messages.
                                messages.Sort("[ReceivedTime]",True)
                                
                                for message in messages:
                                    try:
                                        dt = message.ReceivedTime
                                        year,month,day,hour,minute = dt.year,dt.month,dt.day,dt.hour,dt.minute
                                        if(datetime(year,month,day,hour,minute) >= today):
                                            if(message.Subject.lower() == subject_we_are_looking_for.lower()):
                                                flag_variable = 1
                                                mail        = message.ReplyAll()
                                                result          = email_parser(mail.Body)
                                                Body            = mail_body.format(dataframe.to_html(index = False), sender)
                                                mail.HTMLBody   = Body + mail.HTMLBody
                                                mail.To         = f"{';'.join(to)};{';'.join(result[0])};"
                                                mail.CC         = f"{';'.join(result[1])};"
                                                mail.Save()
                                                mail.Display()
                                                #mail.Send()
                                                break
                                        else:
                                            break
                                    
                                    except:
                                        continue
                                
                                if(flag_variable == 0):
                                    sub_folders = inbox.Folders[i].Folders
                                    
                                    if(len(sub_folders)):
                                        for sub_folder_index in range(0,len(sub_folders)):
                                            try:
                                                messages = inbox.Folders[i].Folder[sub_folder_index].Items
                                                # Filtering messages from the messages.
                                                messages.Sort("[ReceivedTime]",True)
                                                
                                                for message in messages:
                                                    try:
                                                        dt = message.ReceivedTime
                                                        year,month,day,hour,minute = dt.year,dt.month,dt.day,dt.hour,dt.minute
                                                        if(datetime(year,month,day,hour,minute) >= today):
                                                            if(message.Subject.lower() == subject_we_are_looking_for.lower()):
                                                                flag_variable = 1
                                                                mail        = message.ReplyAll()
                                                                result          = email_parser(mail.Body)
                                                                Body            = mail_body.format(dataframe.to_html(index = False), sender)
                                                                mail.HTMLBody   = Body + mail.HTMLBody
                                                                mail.To         = f"{';'.join(to)};{';'.join(result[0])};"
                                                                mail.CC         = f"{';'.join(result[1])};"
                                                                mail.Save()
                                                                mail.Display()
                                                                #mail.Send()
                                                                break
                                                        else:
                                                            break
                                                    except:
                                                        continue

                                                if(flag_variable == 1):
                                                    break
                                            
                                            except:
                                                continue
                                        
                                if(flag_variable == 1):
                                    break
                            
                            except:
                                continue

        
        if(len(circle_mail_not_found) == 0):
            messagebox.showinfo("   Mails Drafted",f"Change Responsible Mails for all mentioned {len(new_unique_circles)} circles have been drafted!")
            
            # Removing all local variables in the current scope
            objects = dir()
            for object in objects:
                if not object.startswith("__"):
                    del object

            return "Successful"
        
        if(len(circle_mail_not_found)):
            messagebox.showwarning("    Mails Drafted",f"Change Responsible Mails have been drafted except for below circles:\n{', '.join(circle_mail_not_found)}")

            # Removing all local variables in the current scope
            objects = dir()
            for object in objects:
                if not object.startswith("__"):
                    del object
            return "Unsuccessful"

        
    except Exception as error:
        import traceback
        messagebox.showerror("  Exception Occured!",f"{traceback.format_exc()}\n{error}")
        
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
                
                print("Function called")
                # Calling the Method (function) for replying the mail.
                flag = mail_checker_and_sender(today_maintenance_date,sender,required_worksheet,unique_circles,dictionary_for_change_responsible_to_mail_id)
                
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
        import traceback
        messagebox.showerror("  ValueError Occured!",f"{traceback.format_exc()}\n{error}")
        
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
        import traceback
        messagebox.showerror("  Exception Occured!",f"{traceback.format_exc()}\n{error}")
        
        # Deleting all the variables before returning the value for "Unsuccessful"
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
        return "Unsuccessful"
    
    except Exception as error:
        import traceback
        messagebox.showerror("  Exception Occured!",f"{traceback.format_exc()}\n{error}")
        
        # Deleting all the variables before returning the value for "Unsuccessful"
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
        return "Unsuccessful"

# circle_reply_task("Manoj Kumar",r"C:\Users\emaienj\Downloads\MPBN Daily Planning Sheet.xlsx")