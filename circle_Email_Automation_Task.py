import os
import sys                                  # Importing the sys to run cmd commands from the script itself.
# import #logging                              # Importing the #logging module for creating logs.
import traceback
from pathlib import Path
from datetime import datetime,timedelta     # Importing datetime and timedelta to get today's maintenance date based on system's current date and time settings.
import pandas as pd                         # Importing Pandas to manipulate the data from the excel sheet.
import win32com.client as win32             # Importing the win32com.client module to create Microsoft Office Suite COM object for sending mails.
from tkinter import *                       # Importing all the classes from tkinter GUI Module of python.
from tkinter import messagebox              # Importing Messagebox from tkinter for displaying messages.
from Custom_Warning import CustomWarning
import pywintypes

# #logging_file_path = "C:/Ericsson_Application_Logs/MPBN Planning Task Logs/"
# Path(#logging_file_path).mkdir(parents=True, exist_ok = True)

# log_file = os.path.join(#logging_file_path,"MPBN_Planning_Task.logs")
today = datetime.now()

# #logging.basicConfig(filename = log_file,
#                     filemode = "a",
#                     format = f"[ {'%(asctime)s'} ]: <<{'%(levelname)s'}>>: ('MPBN_Planning_Task'/{'%(module)s'}): {'%(message)s'}",
#                     datefmt = '%d-%b-%Y %I:%M:%S %p',
#                     encoding = "UTF-8",
#                     level = #logging.DEBUG)

flag  = ""
workbook1 = ""

# Creating the Custom Exception class inheriting the base Exception Class defined in the defaul python libraries, for raising and handling 
# custom Exceptions
class CustomException(Exception):
    def __init__(self,msg):
        self.msg=msg

#####################################################################
#############################    Sendmail   #########################
#####################################################################

def sendmail(dataframe:pd.DataFrame, to:str, cc:str, body:str, subject:str, sender:str) -> None:
    """
        This method creates the mail item in outlook and sends them according to the data fetched from 'Mail ID' sheet of the uploaded 'MPBN Planning Task' workbook.

        Arguments: (dataframe, to, cc, body, subject, sender)
            dataframe : pd.DataFrame
                description =====> Dataframe containing the data to be sent as a table for mail communication to the circles
            
            to : str
                description =====> string containing Email-Ids fetched from 'Mail ID' sheet to be entered in 'To' Field of the Outlook Mail Item.

            cc : str
                description =====> string containing Email-Ids fetched from 'Mail ID' sheet to be entered in 'CC' Field of the Outlook Mail Item.

            body : str
                description =====> string containing the mail body in HTML format to be sent via Outlook Mail Item.
            
            sender:sender
    
    """
    try:
        # Creating the COM object of Microsoft Office Suite (Outlook) for sending mail.
        outlook_mailer=win32.Dispatch('Outlook.Application')
        msg=outlook_mailer.CreateItem(0)            # Creating Mail for sending.
        #logging.info("Created the outlook mail item")

        html_body=body                              # Assigning the Mail Body
        msg.Subject=subject                         # Assigning the Subject Line for mail through the COM Object
        msg.To=to                                   # Assigning the Mail Receipient for mail through COM Object
        msg.CC=cc                                   # Assigning the Mail CC Receipients for mail through COM Object
        
        # Filling the Nan of the dataframe with string 'NA'
        dataframe.fillna("NA",inplace = True)

        # Stylising the dataframe table to make it more presentable in the mail body through the usage of inline CSS.
        dataframe=dataframe.style.set_table_styles([
            {'selector':'th','props':'border:1px solid black; border-collapse : collapse; color:white;padding: 10px; background-color:rgb(0, 51, 204);text-align:center;'},
            {'selector':'tr','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
            {'selector':'td','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
            {'selector':'tr:nth-child(even)','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'}])
        
        dataframe=dataframe.hide(axis='index')                                                                      # Hiding the extra index given by the pandas, although we could also have dropped it.
        msg.HTMLBody=html_body.format(dataframe.to_html(classes = 'table table-stripped',index=False),sender)       # Formatting the Mail Body by entering the important data in the mail body through the usage of .format
        msg.Save()                                                                                                  # Saving the mail in drafts before sending it, as a failsafe mechanism in case any failure arises.
        msg.Send()                                                                                                  # Sending the mail.

        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
    
    except pywintypes.com_error as error:
        error_code = error.excepinfo[5]
        error_source = error.excepinfo[1]
        error_description = error.excepinfo[2]

        if(error_description.__contains__("Outlook does not recognize one or more names.")):
            #logging.debug(f"Found Error in Email-ID :{to}")
            messagebox.showerror("Exception Occurred",f"")
            pass

#####################################################################
############################# Fetch-details #########################
#####################################################################

# Method(Function) for quitting the program entirely.
def quit(event):
    sys.exit(0)

# Method(Function) that drives the task.
def fetch_details(sender:str, workbook:str)->str:
    """
        Arguments : (sender, workbook)
            sender ===> str
               description ====> Name of the user who is running the task.

            workbook ===> str
               description ====> Path of the workbook file selected for the input.
        
        returns : flag
            flag ===> str
               description ====> either 'Unsuccessful' or 'Successful' based upon the status of the completion of the method.
    """
    try:
        #user=subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 
        #workbook=r"C:\Daily\MPBN Daily Planning Sheet.xlsx" # system path from where the program will take the input
        #logging.debug()

        # Checking if any path for the workbook is selected or not.
        global workbook1;workbook1 = workbook
        
        if (len(workbook) == 0):
            # Raising exception when there is no file selected.
            raise CustomException ("Please Browse for the Excel File to continue")
        
        elif (len(workbook) > 0):
            # Reading the Excel Worksheet in the selected workbook in pandas.
            excel=pd.ExcelFile(workbook)
            excel_sheets = excel.sheet_names

            flag_for_daily_planning_sheet = 0
            flag_for_mail_id              = 0
            
            
            if ('Planning Sheet' in excel_sheets):
                flag_for_daily_planning_sheet = 1
            
            if ('Mail Id' in excel_sheets):
                flag_for_mail_id = 1

            if (flag_for_daily_planning_sheet == 0):
                excel.close()
                del excel
                raise CustomException("'Planning Sheet' worksheet is missing, Kindly Check!")
            
            if(flag_for_mail_id == 0):
                excel.close()
                del excel
                raise CustomException("'Mail Id' worksheet is missing, Kindly Check!")
            
            daily_plan_sheet=pd.read_excel(excel,'Planning Sheet')

            # Filling all the blank fields with NA, here inplace is used to reflect the changes made in the deep copy of dataframe.
            daily_plan_sheet.fillna("NA",inplace=True)
            input_error = []                                # Creating an empty list to insert the S.NO of the rows where input errors have been detected.

            # Here we are assingning today's maintenance date to the tomorrow variable with the help of datetime and timedelta
            tomorrow=datetime.now()+timedelta(1) # getting tomorrow date for data execution
            
            try:
                # Here we are coverting the data in the Execution Date column in datetime format of pandas for interoperability between Excel and pandas.
                daily_plan_sheet['Execution Date'] = pd.to_datetime(daily_plan_sheet['Execution Date'])
            
            except Exception as error:
                '''
                    If any exception is thrown during the conversion of Execution Date column data, it would be most probably due to wrong format in which
                    the date has been entered in the required sheet.
                '''
                messagebox.showerror(" Date Format Error!","Please check the Execution Date format in Planning Sheet!")
                
                del daily_plan_sheet
                excel.close()
                del excel
                
                objects = dir()
                for object in objects:
                    if not object.startswith("__"):
                        del object
                
                flag = "Unsuccessful"

            else:
                # Here we are finding out the rows with Execution date other than today's maintenance date.
                for i in range(0,len(daily_plan_sheet)):
                    if (daily_plan_sheet.iloc[i]['Execution Date'].strftime('%Y-%m-%d') != tomorrow.strftime('%Y-%m-%d')):
                        input_error.append(str(daily_plan_sheet.iloc[i]['S.NO']))

                # Filtering the data with today's maintenance data from the dataframe.
                daily_plan_sheet=daily_plan_sheet[daily_plan_sheet['Execution Date'] == tomorrow.strftime('%Y-%m-%d')]

                # Checking if there's no data for today's maintenance date in the sheet.
                if len(daily_plan_sheet)==0:
                    del daily_plan_sheet
                    excel.close()
                    del excel
                    raise CustomException(f"Today's Maintenance Data not Found in the {workbook}, kindly check!")
                
                # Checking if we do have data for Execution Date other than today's maintenance date.
                elif (len(input_error) > 0):
                    del daily_plan_sheet
                    excel.close()
                    del excel
                    raise CustomException(f"All the CR's present are not of Today's Maintenace Date for S.NO : {', '.join(input_error)}")
                
                else:
                    # Reading the Email ID worksheet.
                    Email_ID=pd.read_excel(excel,'Mail Id')

                    # Changing the case of the circle column data to upper case.
                    daily_plan_sheet['Circle'] = daily_plan_sheet['Circle'].str.upper()
                    
                    # Filtering data relevant only for the task at hand.
                    daily_plan_sheet = daily_plan_sheet[['S.NO','Execution Date','Maintenance Window','CR NO','Activity Title','Risk','Location','Circle','Planning Status']]
                    
                    # Filtering all the data from the daily_plan_sheet which are planned.
                    daily_plan_sheet['Planning Status'] = daily_plan_sheet['Planning Status'].str.strip()
                    daily_plan_sheet['Planning Status'] = daily_plan_sheet['Planning Status'].str.upper()
                    daily_plan_sheet = daily_plan_sheet[daily_plan_sheet['Planning Status'] == 'PLANNED']
                
                    
                    if (len(daily_plan_sheet) == 0):
                        del Email_ID
                        del daily_plan_sheet
                        excel.close()
                        del excel
                        
                        raise CustomException('Kindly Enter the Planning Status input in uploaded sheet!')
                    
                    # Creating an empty list to grab the S.NO. of all the rows with Input errors and an empty dataframe to filter out only relevant data from the dataframe.
                    input_error = []
                    result_df = pd.DataFrame()

                    # Finding all the unique circles present in the daily_plan_sheet.
                    circles=list(daily_plan_sheet['Circle'].unique())

                    # Removing any Blank Circle if present in the list of unique circles from the daily_plan_sheet.
                    for i in circles:
                        if (i == 'NA'):
                            circles.remove(i)
                    
                    # Finding all the genuine circles present in the daily_plan_sheet.
                    total_circles_in_planning_sheet = len(circles)
                    
                    # Finding all the unique mail circles defined in Email ID worksheet of the MPBN Planing workbook, so that the duplicated circles are ignored.
                    mail_circles = Email_ID['Circle'].unique()

                    # Getting the remainder circles which are mismatch with the Mail ID Sheet and removing it from the circles.
                    remainder=list(set(circles)-set(mail_circles))
                    circles=list(set(circles)-set(remainder))

                    '''
                        Checking if length of remainder list is bigger than zero, which indicates that there are circles in in the planning sheet
                        which are mismatch with the Mail ID sheet.
                    '''
                    if (len(remainder) > 0):
                        remainder.sort()
                        
                        del Email_ID
                        del daily_plan_sheet
                        excel.close()
                        del excel
                        
                        raise CustomException(f"The mails for these circles will not be sent as there's no mail IDs present in the Mail ID sheet\n{', '.join(remainder)}\nKindly Check! and then retry!")
                    
                    
                    # Finding the data where the data fields which should be non-blank or always correct are left blank or incorrect.
                    for i in range(0,len(daily_plan_sheet)):
                        if (daily_plan_sheet.iloc[i]['CR NO'] == "NA"):
                            input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                            continue
                        
                        if (daily_plan_sheet.iloc[i]['Activity Title'] == "NA"):
                            input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                            continue
                        
                        if (daily_plan_sheet.iloc[i]['Circle'] == "NA"):
                            input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                            continue

                        if (daily_plan_sheet.iloc[i]['Circle'] not in mail_circles):
                            input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                            continue

                        else:
                            result_df = pd.concat([result_df,daily_plan_sheet.iloc[i].to_frame().T],ignore_index = True)

                    input_error = list(set(input_error))
                    input_error.sort()

                    if len(input_error)>0:
                            # messagebox.showwarning("  Input Error Detected",f"Input Error! Kindly Check CR NO, Activity title and Circle fields in Planning Sheet for S.NO : {', '.join(str(num) for num in input_error)}")

                            if(len(result_df) > 0):
                                del result_df
                            
                            del Email_ID
                            del daily_plan_sheet
                            excel.close()
                            del excel

                            flag = "Unsuccessful"
                            raise CustomWarning("  Input Error Detected",f"Input Error! Kindly Check CR NO, Activity title and Circle fields in Planning Sheet for S.NO : {', '.join(str(num) for num in input_error)}")
                            
                    else:
                        daily_plan_sheet_unique_cr = result_df['CR NO'].value_counts().index.to_list()
                        
                        del daily_plan_sheet
                        new_result_df = pd.DataFrame()

                        for index, cr_no in enumerate(daily_plan_sheet_unique_cr):
                            counter = result_df['CR NO'].value_counts().to_list()[index]
                            temp_df = pd.DataFrame()
                            temp_df = result_df[result_df['CR NO'] == cr_no]
                            if (counter > 1):
                                if (len(temp_df['Circle'].unique()) > 1):
                                    for i in range(0,len(temp_df)):
                                        try:
                                            input_error.append(temp_df.iloc[i]['S.NO'])
                                        
                                        except TypeError as err:
                                            del Email_ID
                                            del result_df
                                            del new_result_df
                                            del temp_df
                                            excel.close()
                                            del excel
                                            if (err.__contains__("'<' not supported between instances of 'str' and 'int'")):
                                                messagebox.showerror("  S.NO error","Kindly Check all the S.NO. in the Planning Sheet, either ther's an empty S.NO. or a string is used as a S.NO.!")
                                                flag = "Unsuccessful"
                                            
                                            else:
                                                messagebox.showerror("  TypeError",err)
                                                flag = "Unsuccessful"
                                        
                                        except Exception as err:
                                            del Email_ID
                                            del result_df
                                            del new_result_df
                                            del temp_df
                                            excel.close()
                                            del excel
                                            messagebox.showerror("  Exception Occured!",err)
                                            flag = "Unsuccessful"
                                        
                                        else: 
                                            pass
                                        
                                        continue
                                    
                                if (len(temp_df['Circle'].unique()) == 1):
                                    new_result_df = pd.concat([new_result_df, temp_df.iloc[0].to_frame().T],ignore_index = True)
                            else:
                                new_result_df = pd.concat([new_result_df,temp_df.iloc[0].to_frame().T],ignore_index = True)
                            
                            del temp_df
                        
                        daily_plan_sheet = new_result_df.copy(deep = True)
                        
                        del new_result_df
                        del result_df
                        
                        for i in range(0,len(circles)):
                            execution_date=[]       #  list for collecting execution date of each Cr
                            circle=[]               #  list for collecting circle of each CR
                            maintenance_window=[]   #  list for collecting the maintenance window of each CR
                            cr_no=[]                #  list for collecting the CR No
                            activity_title=[]       #  list for collecting the activity title each CR
                            risk=[]                 #  list for collecting the risk level of each CR
                            location=[]             #  list for collecting the location of each CR

                            for j in range(0,len(daily_plan_sheet)):
                                if (daily_plan_sheet.iloc[j]['Circle']==circles[i]): # Adding constraint to check for CRs for next date only
                                    execution_date.append(daily_plan_sheet.iloc[j]['Execution Date'])
                                    maintenance_window.append(daily_plan_sheet.iloc[j]['Maintenance Window'])
                                    cr_no.append(daily_plan_sheet.iloc[j]['CR NO'])
                                    activity_title.append(daily_plan_sheet.iloc[j]['Activity Title'])
                                    risk.append(daily_plan_sheet.iloc[j]['Risk'])
                                    circle.append(daily_plan_sheet.iloc[j]['Circle'])
                                    location.append(daily_plan_sheet.iloc[j]['Location'])

                            dictionary_for_insertion={'Execution Date':execution_date, 'Maintenance Window':maintenance_window, 'CR NO':cr_no, 'Activity Title':activity_title, 'Risk':risk,'Location':location,'Circle':circle}
                            dataframe=pd.DataFrame(dictionary_for_insertion)
                            dataframe.reset_index(drop=True,inplace=True)
                            dataframe.fillna("NA",inplace=True) #adding inplace to replace nan or NaN with the string NA or else it won't replace the nan values
                            dataframe['Execution Date'] = pd.to_datetime(dataframe['Execution Date'],format = '%m/%d/%Y')
                            dataframe['Execution Date'] = dataframe['Execution Date'].dt.strftime('%m/%d/%Y')

                            # Taking the circle in the cir variable to avoid writing circles[i] again & again.
                            cir=circles[i]

                            # Creating the mail body and sending the circle mails.
                            for i in range(0,len(Email_ID)):
                                if (Email_ID.at[i,'Circle'] == cir):
                                    row_to_fetch = i
                                    to=Email_ID.iloc[row_to_fetch]['To Mail List']
                                    cc=Email_ID.iloc[row_to_fetch]['Copy Mail List']
                                    
                                    subject=f"Connected End Nodes and their services on MPBN devices: {cir}_{tomorrow.strftime('%d-%m-%Y')}"
                                    body="\
                                        <html>\
                                            <body>\
                                                <div><p>Hi team,<br></p>\
                                                    <p>Please confirm below points so that we will approve CR’s.<br></p>\
                                                    <p>1)  End nodes and service details are required which are running on respective MPBN device (In case of changes on Core/STP/DRA/PACO/HLR/IN connected MPBN nodes).</p>\
                                                    <p>2)  Design Maker & Checker confirmation mail need to be shared for all planned activity on Core/STP/DRA/PACO/HLR/IN connected MPBN nodes.</p>\
                                                    <p>3)  KPI & Tester details need to be shared for all impacted nodes in Level-1 CR’s (SA). Also same details need to be shared for all Level-2 CR’s (NSA) with respect to changes on Core/STP/DRA/PACO/HLR/IN connected MPBN nodes.</p>\
                                                    <p>4)  Is redundancy working or all parameters ensured in shared plan for SAPC integration/reachability related activity on MPBN nodes (Yes/No, Applicable only in case of SAPC integration)?</p>\
                                                    <p>5)  Is there any impact on 5G services/ KPI related to IN/Core/PACO/RAN?</p>\
                                                    <p>6)  Please confirm Integration/Expansion/Reachability activity will not impact running services (Yes/No)<br><br></p>\
                                                </div>\
                                                <div>\
                                                    <p>{}</p>\
                                                </div>\
                                                <div>\
                                                    <p>Regards,<br>{}<br>SRF MPBN | SDU Bharti<br>Ericsson India Global Services Pvt. Ltd.</p>\
                                                    </div>\
                                            </body>\
                                        </html>\
                                        "
                                    # Formatting the dataframe's 'Execution Date' column to "dd-mm-YYYY" format.
                                    dataframe['Execution Date'] = pd.to_datetime(dataframe['Execution Date'],format = "%m/%d/%Y")
                                    dataframe['Execution Date'] = dataframe['Execution Date'].dt.strftime("%d-%m-%Y")
                                    
                                    # Calling the Sendmail Method(Function) for sending circle emails.
                                    sendmail(dataframe,to,cc,body,subject,sender)

                        # Message Showing that the all the mails to all the present circles have been successfully sent.
                        messagebox.showinfo("  Mail Sent Successfully!",f"All Mails for mentioned planned {total_circles_in_planning_sheet} Circles in Daily Planning Sheet have been sent!")

                        del Email_ID
                        del daily_plan_sheet
                        del dataframe
                        excel.close()
                        del excel
                        objects = dir()
                        for object in objects:
                            if not object.startswith("__"):
                                del object

                        flag = "Successful"
    
    # Handling Custom Exceptions which are raised in the above try section.
    except CustomWarning:
        flag = "Unsuccessful"

    except CustomException as error:
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
        messagebox.showerror("  Data can't be found",error)
        flag = "Unsuccessful"
    
    # Handling the AttributeError Exception.
    except AttributeError as e:
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
        messagebox.showerror("  Heading missing!",f"Kindly Check the below Heading in Planning Sheet\n{e}")
        flag = "Unsuccessful"
    
    # Handling the Exception when the file that's required for the task is opened in background.
    except PermissionError as e:
        e = str(e).split(":")
        e.remove(e[0])
        e = ':'.join(e)
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
        messagebox.showerror("    Permission Error!",f"Kindly Close the selected {e} if opened in Excel!")
        flag = "Unsuccessful"

    # Handling other exceptions that are not handled.
    except Exception as e:
        messagebox.showerror("  Exception Occurred",f"{traceback.format_exc()}\n{e}")

        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
        flag = "Unsuccessful"
    
    finally:
        #logging.info(f"Returning \'{flag}\' as status to the main GUI")
        #logging.shutdown()
        return flag
    
#fetch_details("Enjoy Maity",r"C:\Users\emaienj\Downloads\MPBN Daily Planning Sheet - Copy.xlsx")