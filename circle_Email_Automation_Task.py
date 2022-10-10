import sys
from datetime import datetime,timedelta
import pandas as pd
import win32com.client as win32
from tkinter import *
from tkinter import messagebox

class TomorrowDataNotFound(Exception):
    def __init__(self,msg):
        self.msg=msg

#####################################################################
#############################    Sendmail   #########################
#####################################################################

def sendmail(dataframe,to,cc,body,subject,sender):
    outlook_mailer=win32.Dispatch('Outlook.Application')
    msg=outlook_mailer.CreateItem(0)
    html_body=body
    msg.Subject=subject
    msg.To=to
    msg.CC=cc
    dataframe=dataframe.style.set_table_styles([
        {'selector':'th','props':'border:1px solid black; color:white; background-color:rgb(0, 51, 204);text-align:center;'},
        {'selector':'tr','props':'border:1px solid black;text-align:center;'},
        {'selector':'td','props':'border:1px solid black;text-align:center;'},
        {'selector':'tr:nth-child(even)','props':'border:1px solid black;text-align:center;'}])
    dataframe=dataframe.hide(axis='index')
    msg.HTMLBody=html_body.format(dataframe.to_html(index=False),sender)
    msg.Save()
    msg.Send()

    # messagebox.showinfo("   Successful Completion","Circle Email Automation task completed")

#####################################################################
############################# Fetch-details #########################
#####################################################################

def quit(event):
    sys.exit(0)

def fetch_details(sender,workbook):
    try:

        #user=subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 

        #workbook=r"C:\Daily\MPBN Daily Planning Sheet.xlsx" # system path from where the program will take the input
        if (len(workbook) == 0):
            raise TomorrowDataNotFound ("Please Browse for the Excel File to continue")
        elif (len(workbook) > 0):
            excel=pd.ExcelFile(workbook)
            daily_plan_sheet=pd.read_excel(excel,'Planning Sheet')
            daily_plan_sheet.fillna("NA",inplace=True)

            tomorrow=datetime.now()+timedelta(1) # getting tomorrow date for data execution
            daily_plan_sheet=daily_plan_sheet[daily_plan_sheet['Execution Date']==tomorrow.strftime('%Y-%m-%d')]

            if len(daily_plan_sheet)==0:
                raise TomorrowDataNotFound("Today's Maintenance Data not Found in the , kindly check!")
            
            else:
                
                Email_ID=pd.read_excel(excel,'Mail Id')

                for j in range(0,len(daily_plan_sheet)):
                    temp=daily_plan_sheet.iloc[j]['Circle']
                    daily_plan_sheet.iloc[j]['Circle']=temp.upper()

                circles=daily_plan_sheet['Circle'].unique()
                total_circles_in_planning_sheet = len(circles)
                email_id_list=Email_ID['Circle'].unique()
                # print(circles) # checking for all the unique values of circles in the MPBN Planning Sheets
                remainder=list(set(circles)-set(email_id_list))
                remainder.sort()
                remainder_list=",".join(remainder)
                
                circles=list(set(circles)-set(remainder))
                #daily_plan_sheet['Execution Date']=daily_plan_sheet['Execution Date'].dt.to_pydatetime()
                    
                for i in range(0,len(circles)):

                    execution_date=[]       #  list for collecting execution date of each Cr
                    circle=[]               #  list for collecting circle of each CR
                    maintenance_window=[]   #  list for collecting the maintenance window of each CR
                    cr_no=[]                #  list for collecting the CR No
                    activity_title=[]       #  list for collecting the activity title each CR
                    risk=[]                 #  list for collecting the risk level of each CR
                    location=[]             #  list for collecting the location of each CR

                    for j in range(0,len(daily_plan_sheet)):
                        #print(str(tomorrow.strftime("%d-%m-%Y")))
                        
                        if daily_plan_sheet.iloc[j]['Circle']==circles[i]: # Adding constraint to check for CRs for next date only

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
                    # dataframe['Execution Date']=pd.to_datetime(dataframe['Execution Date'])
                    dataframe['Execution Date']=dataframe['Execution Date'].dt.strftime('%d-%m-%Y')

                    #print(dataframe.head())

                    cir=circles[i]

                    if cir=='DL':
                        row_to_fetch=0

                    elif cir=='UPE':
                        row_to_fetch=1

                    elif cir=='UPW':
                        row_to_fetch=2
                    
                    elif cir=='PB':
                        row_to_fetch=3

                    elif cir=='HRY':
                        row_to_fetch=4

                    elif cir=='HP':
                        row_to_fetch=5

                    elif cir=='JK':
                        row_to_fetch=6

                    elif cir=='BH':
                        row_to_fetch=7

                    elif cir=='OR':
                        row_to_fetch=8

                    elif cir=='KOL':
                        row_to_fetch=9

                    elif cir=='WB':
                        row_to_fetch=10

                    elif cir=='AS':
                        row_to_fetch=11

                    elif cir=='NE':
                        row_to_fetch=12

                    elif cir=='GUJ':
                        row_to_fetch=13

                    elif cir=='RAJ':
                        row_to_fetch=14
                    
                    elif cir=='MH':
                        row_to_fetch=15

                    elif cir=='MP':
                        row_to_fetch=16

                    elif cir=='MU':
                        row_to_fetch=17

                    elif cir=='AP':
                        row_to_fetch=18

                    elif cir=='KK':
                        row_to_fetch=19

                    elif cir=='TN':
                        row_to_fetch=20
                        
                    elif cir=='KL':
                        row_to_fetch=21

                    elif cir=='CHN':
                        row_to_fetch=22

                    else :
                        pass


                    to=Email_ID.iloc[row_to_fetch]['To Mail List']
                    cc=Email_ID.iloc[row_to_fetch]['Copy Mail List']
                    
                    subject=f"Connected End Nodes and their services on MPBN devices: {cir}_{tomorrow.strftime('%Y-%m-%d')}"
                    body="""
                        <html>        
                            <body>
                                <div><p>Hi team,<br></p>
                                    <p>Please confirm below points so that we will approve CR’s.<br></p>
                                    <p>1)  End nodes and service details are required which are running on respective MPBN device (in case of changes on Core/PACO/HLR devices ).</p>
                                    <p>2)  Design Maker & Checker confirmation mail need to be shared for all planned activity on Core/PACO/HLR devices.</p>
                                    <p>3)  KPI & Tester details need to be shared for all impacted nodes in Level-1 CR’s (SA).Also same details need to be shared for all Level-2 CR’s (NSA) with respect to changes on Core/PACO/HLR devices.<br><br></p>
                                </div>
                                <div>
                                    <p>{}</p>
                                </div>
                                <div>
                                    <p>Regards<br>{}<br>Ericsson India Global Services Pvt. Ltd.</p>
                                    </div>
                            </body>
                        </html>
                        """
                    sendmail(dataframe,to,cc,body,subject,sender)
                    #messagebox.showinfo("   Mail Successfully Sent",f"Mail Sent for the Circle {cir}\n\nPlease! Press The Enter Key or Click The OK Button To Proceed")
                if len(remainder)>0:
                    flag=2
                    #messagebox.showwarning("  Mail Not Sent",f"Mail could not be sent for {remainder_list} as there's no email id present for the {remainder_list} in the Email ID sheet in {workbook}")
                    messagebox.showinfo("  Mail Sent",f"Mail Sent For { len(total_circles_in_planning_sheet) - len(remainder) }/{ len(total_circles_in_planning_sheet) } Circles\n remainder: {remainder_list}")
                else:
                    flag=1
                    messagebox.showinfo("  Mail Sent Successfully","All The Mails for All the Circles in Daily Planning Sheet Have Been Sent")
                return flag

    except FileNotFoundError:
        working_directory=workbook
        messagebox.showwarning("    File not Found","Check {} for the Planning Sheet".format(working_directory)).bind("<Return>",quit)
        sys.exit(0)
    
    except ValueError:
         working_directory=workbook
         messagebox.showwarning(" Value Error"," Check {} for all the requirement sheets".format(working_directory)).bind("<Return>",quit)
         sys.exit(0)
    
    except TomorrowDataNotFound as error:
        messagebox.showerror("  Data can't be found",error).bind("<Return>",quit)
        sys.exit(0)
    
    