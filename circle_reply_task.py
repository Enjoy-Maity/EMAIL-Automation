import pandas as pd                         # Importing Pandas for manipulation and reading of the excel sheet
import win32com.client as win32             # Importing win32com for opening and creation of outlook mail
from tkinter import messagebox              # Importing messagebox for rasing dialogues
from datetime import datetime, timedelta    # Importing datetime to manipulate time related variables and getting today's maintenance date


# Creating Custom Exception inheriting base default Exception class for raising, handling and custom exceptions.
class CustomException(Exception):
    def __init__(self,title,message):
        self.title = title
        self.message = message
        super().__init__(title,message)
        messagebox.showerror(self.title, self.message)

# Mail checker and send
def mail_checker_and_sender(subject_we_are_looking_for,body,dataframe,sender,to):
    # Creating an COM object of Microsoft Office Client Suite (Outlook) through win32com.client module.
    outlook     = win32.Dispatch("Outlook.Application")
    mapi        = outlook.GetNamespace("MAPI")              # MAPI is an API for messaging to do functions like fetching, and manipulation of mails in outlook

    # Getting the inbox folder from the outlook.
    inbox       = mapi.GetDefaultFolder(6)

    # Getting all the mails present in the inbox folder.
    messages    = inbox.Items

    # Getting today's date
    today       = datetime.now()

    # Formatting today's date so that we can compare it with the received date and time of the messages in the inbox
    today       = today.replace(hour = 10, minute = 0, second = 0).strftime('%Y-%m-%d %H:%M %p')
    
    # Filtering messages from the messages.
    messages    = messages.Restrict("[ReceivedTime] >='"+today+"'")
    messages.Sort("[ReceivedTime]",True)

    # Creating a flag variable for searching the mail.
    flag_variable = 0

    # Changing the format of the dataframe containing relevant data to be presented in a more presentable manner through the usage of inline CSS.
    dataframe=dataframe.style.set_table_styles([
        {'selector':'th','props':'border:1px solid black; border-collapse : collapse; color:white;padding: 10px; background-color:rgb(0, 51, 204);text-align:center;'},
        {'selector':'tr','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
        {'selector':'td','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
        {'selector':'tr:nth-child(even)','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'}])

    dataframe=dataframe.hide(axis='index') # hiding the index coloumn

    if(flag_variable == 0):
        for message in messages:
            if(message.Subject == subject_we_are_looking_for):
                flag_variable = 1
                mail        = message.ReplyAll()
                Body        = body.format(dataframe.to_html(index = False), sender)
                mail.HTMLBody   = Body + mail.HTMLBody
                mail.To         = f"{to};{mail.To};"
                mail.CC         = f"{mail.CC};"
                mail.Save()
                mail.Send()
                break

    if(flag_variable == 0):
        
        # Iterating through the messages for finding the mail with subject line.
        for i in range(0,len(inbox.Folders)):
            messages = inbox.Folders[i].Items
            
            # Filtering messages from the messages.
            messages    = messages.Restrict("[ReceivedTime] >='"+today+"'")
            messages.Sort("[ReceivedTime]",True)
            for message in messages:
                if(message.Subject == subject_we_are_looking_for):
                    flag_variable = 1
                    mail        = message.ReplyAll()
                    Body        = body.format(dataframe.to_html(index = False), sender)
                    mail.HTMLBody   = Body + mail.HTMLBody
                    mail.To         = f"{to};{mail.To};"
                    mail.CC         = f"{mail.CC};"
                    mail.Save()
                    mail.Send()
                    break
        
        if (flag_variable == 0):
            raise CustomException(" Mail For Reply Not Found!","Kindly check the mail box for the reply messages, as no reply thread found")

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
                unique_circles = list(required_worksheet['Circle'].unique())

                # Creating dictionary for change_responsible to mail ID
                dictionary_for_change_responsible_to_mail_id = dict(zip(mail_id_sheet['Change Responsible'],mail_id_sheet['Mail ID']))
                
                # Iterating through the unique circles for replying to circles.
                for cir in unique_circles:

                    # Filtering out rows based on circle
                    temp_df = required_worksheet[required_worksheet["Circle"] == cir]

                    # Filtering out data for just required columns 
                    temp_df = temp_df[["Execution Date","Maintenance Window","CR NO","Activity Title","Risk","Location","Circle","No. of Node Involved","CR Belongs to Same Activity of Previous CR- Yes/NO","Change Responsible"]]

                    # Creating a variable to get today's maintenance date
                    today_maintenance_date = datetime.now() + timedelta(1)

                    # Formatting the execution date of the temp_df dataframe
                    temp_df['Execution Date'] = pd.to_datetime(temp_df['Execution Date'], format = "%m/%d/%Y")
                    temp_df['Execution Date'] = temp_df['Execution Date'].dt.strftime("%d-%b-%Y")

                    # Creating a to variable for sending mails to
                    to = ''

                    # Iterating through the temp_df to attach to the to variable.
                    for i in range(0,len(temp_df)):
                        to = f"{dictionary_for_change_responsible_to_mail_id[temp_df.iloc[i]['Change Responsible']]};{to}"

                    # Making the subject for finding in the inbox
                    subject_we_are_looking_for = f"RE: Connected End Nodes and their services on MPBN devices: {cir}_{today_maintenance_date.strftime('%d-%m-%Y')}"
                    mail_body = "<html>\
                                    <body>\
                                        <div>\
                                            <p>Hi Team,<br><br>Kindly find executor detail for below mention CR.</p>\
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
                        
                    # Calling the Method (function) for replying the mail.
                    mail_checker_and_sender(subject_we_are_looking_for,mail_body,temp_df,sender,to)

                return "Successful"
    
    except CustomException:
        return "Unsuccessful"

    except ValueError as error:
        messagebox.showerror("  Exception Occured!",error)
        return "Unsuccessful"
    
    except Exception as error:
        messagebox.showerror("  Exception Occured!",error)
        return "Unsuccessful"

# circle_reply_task("Arka Maiti",r"C:\Users\emaienj\OneDrive - Ericsson\Documents\MPBN Daily Planning Sheet new copy.xlsx")