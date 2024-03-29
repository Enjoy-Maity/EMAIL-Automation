import sys                                  # Importing the sys to run cmd commands from the script itself.
from datetime import datetime,timedelta     # Importing datetime and timedelta to get today's maintenance date based on system's current date and time settings.
import pandas as pd                         # Importing Pandas to manipulate the data from the excel sheet.
import win32com.client as win32             # Importing the win32com.client module to create Microsoft Office Suite COM object for sending mails.
from tkinter import messagebox              # Importing Messagebox from tkinter for displaying messages.

flag = ""
# workbook1 = ""

# Creating the Custom Exception Class for raising and handling custom Exceptions that are not defined by-default in the system.
class TomorrowDataNotFound(Exception):
    def __init__(self,msg):
        self.msg=msg

# Function for sending the mail.
def sendmail(dataframe,to,cc,body,subject,north_and_west_region,east_and_south_region,sender,i):
    # Creating an COM object of Microsoft Office Client Suite (Outlook) through win32com.client module.
    outlook_mailer=win32.Dispatch('Outlook.Application')
    msg=outlook_mailer.CreateItem(0)                            # Creating Mail for sending.
    html_body=body                                              # Setting the body of the mail.
    msg.Subject=subject                                         # Setting the Subject Line of the mail.
    msg.To=to                                                   # Setting the mail receipient IDs
    msg.CC=cc                                                   # Setting the mail CC receipient IDs

    # Changing the format of the dataframe containing relevant data to be presented in a more presentable manner through the usage of inline CSS.
    dataframe=dataframe.style.set_table_styles([
        {'selector':'th','props':'border:1px solid black; border-collapse : collapse; color:white;padding: 10px; background-color:rgb(0, 51, 204);text-align:center;'},
        {'selector':'tr','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
        {'selector':'td','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
        {'selector':'tr:nth-child(even)','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'}])

    dataframe=dataframe.hide(axis='index') # hiding the index coloumn

    '''
        Adding all the relevant data to the mail body like sender details, data table, etc. before sending the mail.
    '''
    msg.HTMLBody=html_body.format(i,north_and_west_region,east_and_south_region,dataframe.to_html(index=False),sender) 
    
    '''
        Saving the mail first to mail drafts, incase there's any failure before sending the mail like power failure, or other such failures, 
        so that the user can send it manually.
    '''
    msg.Save()

    # Sending the mail.
    msg.Send()

    # Deleting objects and variables in local sccope
    objects = dir()
    for object in objects:
        if not object.startswith("__"):
            del object


# Function/Method for quitting the program.
def quit(event):
    sys.exit(0)

# Main Working Function/Method
def paco_cscore(sender,workbook,north_and_west_region,east_region_and_south_region):
    try: 
        #user=subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 
        #workbook=r"C:\Daily\MPBN Daily Planning Sheet.xlsx" # system path from where the program will take the input

        # global workbook1;workbook1 = workbook
        # Reading the Worksheet in pandas.
        wb               = pd.ExcelFile(workbook)
        wb_sheets        = wb.sheet_names

        flag_for_planning_sheet = 0
        flag_for_mail_id        = 0

        if('Planning Sheet' in wb_sheets):
            flag_for_planning_sheet = 1
            
        if('Mail Id' in wb_sheets):
            flag_for_mail_id        = 1
                
        if(flag_for_mail_id == 0):
            del wb
            raise TomorrowDataNotFound("Mail Id sheet not Found in the Workbook, Kindly Check!")
        
        if(flag_for_planning_sheet == 0):
            del wb
            raise TomorrowDataNotFound("Planning sheet not Found in the Workbook, Kindly Check!")
        
        daily_plan_sheet = pd.read_excel(workbook,'Planning Sheet')
        daily_plan_sheet.fillna("NA",inplace=True)

        # Reading the Mail Id sheet to a dataframe so that we can access the mail IDs for sending mails for the respective domain kpis
        Email_Id=pd.read_excel(workbook,'Mail Id')
        
        # Creating an empty list to grab all the S.NO of the rows where there's any input error made by the user.
        input_error = []
        tomorrow = datetime.now() + timedelta(1) # getting tomorrow date for data execution
            
        daily_plan_sheet['Execution Date'] = pd.to_datetime(daily_plan_sheet['Execution Date'])
        for i in range(0,len(daily_plan_sheet)):
            # Finding out rows with S.NO where the execution date is not of today's maintenance date
            if (daily_plan_sheet.iloc[i]['Execution Date'].strftime('%Y-%m-%d') != tomorrow.strftime('%Y-%m-%d')):
                input_error.append(str(daily_plan_sheet.iloc[i]['S.NO']))
        
        # Filtering all the relevant data (rows) from the daily_plan_sheet dataframe pertaining to today's maintenance date.
        daily_plan_sheet=daily_plan_sheet[daily_plan_sheet['Execution Date']==tomorrow.strftime('%Y-%m-%d')]
        

        if (len(daily_plan_sheet) == 0):
                del daily_plan_sheet
                del Email_Id
                del wb
                # Raising Custom Exception defined through a class inheriting the base Exception class (defined in the default Python lib modules),
                # for the case when there's no data pertaining today's maintenance date.
                raise TomorrowDataNotFound("Data for tomorrow's date is not present in the MPBN Daily Planning Sheet, kindly check!")
        
        elif (len(input_error) > 0):
                del daily_plan_sheet
                del Email_Id
                del wb
                # Raising Custom Exception defined through a class inheriting the base Exception class (defined in the default Python lib modules) for 
                # illegal date, other than today's maintenance date present in our dataframe.
                raise TomorrowDataNotFound(f"All the CR's present are not of Today's Maintenace Date for S.NO : {', '.join(input_error)}")    
            
        else:
            # Creating a list of available list of interdomains to send mails.
            list_of_interdomains=[]

            # Reading each Interdomain sheet and replacing the blank fields that have been replaced by the default value of the pandas that is "Nan"
            # Note---> Pandas replaces all the blank fields with Nan which stands for 'Not A Number' to denote blank fields and these are reflected in the dataframe wherever the dataframe is used.
            if("CS Core-Inter Domain" in wb_sheets):
                list_of_interdomains.append("CS-Core")
                df2=pd.read_excel(workbook,sheet_name="CS Core-Inter Domain")
                df2.fillna(" ",inplace=True)
            
            if("PS Core-Inter Domain" in wb_sheets):
                list_of_interdomains.append("PS-Core")
                df=pd.read_excel(workbook,sheet_name="PS Core-Inter Domain")
                df.fillna(" ",inplace=True)
            
            if("RAN-Inter Domain" in wb_sheets):
                list_of_interdomains.append("RAN")
                df3=pd.read_excel(workbook,sheet_name="RAN-Inter Domain")
                df3.fillna(" ",inplace=True)
            
            if("VAS-Inter Domain" in wb_sheets):
                list_of_interdomains.append("VAS")
                df4=pd.read_excel(workbook,sheet_name="VAS-Inter Domain")
                df4.fillna(" ",inplace=True)

            # Formatting the date for the Subject line like dates ending with 1 to have 'st' suffix, ending with 2 to have 'nd' suffix and so on.
            suffix=["st","nd","rd","th"]
            date_end_digit = int(tomorrow.strftime("%d"))%10        # here we're finding the end digit of the date i.e. whether it's 1,2,3 or any other digit.
            date_digits = int(tomorrow.strftime("%d"))%100          # here we're finding the date digits in two-digit format
            if date_digits<10 or date_digits>20:                    # here we are segregating the dates so that the proper suffix for dates 0-10 and 20-31 are given proper suffix
                if date_end_digit == 1:                             # but dates from 11-19 have common suffix of 'th'.
                    suffix_for_date = suffix[0]
                elif date_end_digit == 2:
                    suffix_for_date = suffix[1]
                elif date_end_digit == 3:
                    suffix_for_date = suffix[2]
                else:
                    suffix_for_date = suffix[3]
            else:
                suffix_for_date = suffix[3]
            for_date = tomorrow.strftime("%d{}_%b'%y").format(suffix_for_date)  # here we formatted the date with the relevant suffix so that we can add it to our subject.

            for i in list_of_interdomains:
                subject=f"KPI Monitoring | {i} for MPBN CRs | {for_date}"   # Formatting our subject line with date and respective to domain kpi 
                if (i == "CS-Core"):
                    to=Email_Id.iloc[24]['To Mail List']
                    cc=Email_Id.iloc[24]['Copy Mail List']
                    # to = 'enjoy.maity@ericsson.com'
                    # cc = 'enjoy.maity@ericsson.com'
                    dataframe=df2
                    
                    # Checking if the dataframe for the respective interdomain is empty or not, i.e. whether there's any data to sent to the respective 
                    # Interdomain mail receipients.
                    if (len(dataframe) == 0):
                        continue

                elif (i == "PS-Core"):
                    to=Email_Id.iloc[23]['To Mail List']
                    cc=Email_Id.iloc[23]['Copy Mail List']
                    # to = 'enjoy.maity@ericsson.com'
                    # cc = 'enjoy.maity@ericsson.com'
                    dataframe=df
                    
                    # Checking if the dataframe for the respective interdomain is empty or not, i.e. whether there's any data to sent to the respective 
                    # Interdomain mail receipients.
                    if (len(dataframe) == 0):
                        continue

                elif (i == "RAN"):
                    to=Email_Id.iloc[25]['To Mail List']
                    cc=Email_Id.iloc[25]['Copy Mail List']
                    # to = 'enjoy.maity@ericsson.com'
                    # cc = 'enjoy.maity@ericsson.com'
                    dataframe=df3
                    
                    # Checking if the dataframe for the respective interdomain is empty or not, i.e. whether there's any data to sent to the respective 
                    # Interdomain mail receipients.
                    if (len(dataframe) == 0):
                        continue

                elif (i == "VAS"):
                    to=Email_Id.iloc[26]['To Mail List']
                    cc=Email_Id.iloc[26]['Copy Mail List']
                    # to = 'enjoy.maity@ericsson.com'
                    # cc = 'enjoy.maity@ericsson.com'
                    dataframe=df4
                    
                    # Checking if the dataframe for the respective interdomain is empty or not, i.e. whether there's any data to sent to the respective 
                    # Interdomain mail receipients.
                    if (len(dataframe) == 0):
                        continue
                
                # Creating the mail body for the mail in HTML with spaces left by {} for the data to be formatted in during sending the mail.
                mpbn_html_body="\
                    <html>\
                        <body>\
                            <div>\
                                    <p>Hi Team,</p>\
                                    <p>Please find below the list of MPBN activity which includes Core nodes, so KPI monitoring required. Impacted nodes with KPI details given below. Please share KPI monitoring resource from your end.<br><br></p>\
                                    <p>@{} Team: Please contact below spoc region wise if any issue with KPI input.<br><br></p>\
                                    <p>{}: North region and west region <br>\
                                       {}: East region and South region <br></p>\
                                    <p>Note:-If there is any deviation in KPI please call to Executer before 6 AM. After that please call to technical validator/Team Lead.<br><br></p>\
                            </div>\
                            <div>\
                                {}\
                            </div>\
                            <div>\
                                    <p>With Regards<br>{}<br>SRF MPBN | SDU Bharti<br>Ericsson India Global Services Pvt. Ltd.</p>\
                            </div>\
                        </body>\
                    </html>"
                # Calling the Sendmail function with relevant arguments for sending the mails to relevant interdomain mail recepients according to the data served.
                sendmail(dataframe,to,cc,mpbn_html_body,subject,north_and_west_region,east_region_and_south_region,sender,i)
                
                # Message showing that the respective selected interdomain mail has been sent.
                messagebox.showinfo("     Mail Sent Info",f"Mail sent for {i} Interdomain KPIs!")
                if(i == "PS-Core"):
                    del df
                
                if(i == "CS-Core"):
                    del df2
                
                if(i == "RAN"):
                    del df3
                
                if(i == "VAS"):
                    del df4

                del dataframe
            
        # Message showing that all the interdomain mails have been successfuly sent.
        messagebox.showinfo(" Mails Sent Successfully","Mails for all Interdomain Kpis have been sent!")

        #Deleting all the user-defined variables before calling returning back to the main gui program
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
        
        del daily_plan_sheet
        del Email_Id
        del wb

        flag = "Successful"

   
    # Exception Handling in case File not found, in our case the workbook path is illegal.
    except FileNotFoundError:
        working_directory = r"C:\Daily"
        messagebox.showerror(" File not Found","Check {} for MPBN Daily Planning Sheet.xlsx".format(working_directory))
        
        # Deleting all the variables before returning the value for "Unsuccessful"
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object

        flag = "Unsuccessful"
    
    # Handling the Exception when the file that's required for the task is opened in background.
    except PermissionError as e:
        e = str(e).split(":")
        e.remove(e[0])
        e = ':'.join(e)
        messagebox.showerror("    Permission Error!",f"Kindly Close the selected {e} if opened in Excel!")
        
        # Deleting all the variables before returning the value for "Unsuccessful"
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object

        flag = "Unsuccessful"

    # Exception for handling Value error, in our case when the relevant Sheet is missing the workbook.
    except ValueError:
        working_directory = r"C:\Daily"
        messagebox.showwarning("    Value Error","Check {} for MPBN Daily Planning Sheet.xlsx for all the requirement sheet".format(working_directory))
         
        # Deleting all the variables before returning the value for "Unsuccessful"
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object

        flag =  "Unsuccessful"

    # Custom Exception for handling Date Error, in our case Wrong dates other than today's maintenance date present in our data.
    except TomorrowDataNotFound as error:
        messagebox.showerror(" Data for tomorrow's date not found",error)
        
        # Deleting all the variables before returning the value for "Unsuccessful"
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
                
        flag = "Unsuccessful"
    
    except Exception as error:
        import traceback
        messagebox.showerror(" Exception Occurred!",f"{traceback.format_exc()}\n{error}")
        
        # Deleting all the variables before returning the value for "Unsuccessful"
        objects = dir()
        for object in objects:
            if not object.startswith("__"):
                del object
                
        flag = "Unsuccessful"
    
    finally:
        import gc

        # excel = win32.Dispatch("Excel.Application")

        # if(len(workbook1) > 0):
        #     wb = excel.Workbooks.Open(workbook1)
        #     wb.Close()

        # excel.Application.Quit()
        gc.collect()
        
        return flag
    
# paco_cscore("Manoj Kumar",r"C:/Users/emaienj/Downloads/MPBN Daily Planning Sheet.xlsx","","")