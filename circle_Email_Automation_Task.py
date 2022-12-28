import sys
from datetime import datetime,timedelta
import pandas as pd
import win32com.client as win32
from tkinter import *
from tkinter import messagebox

class CustomException(Exception):
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
        {'selector':'th','props':'border:1px solid black; border-collapse : collapse; color:white;padding: 10px; background-color:rgb(0, 51, 204);text-align:center;'},
        {'selector':'tr','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
        {'selector':'td','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'},
        {'selector':'tr:nth-child(even)','props':'border:1px solid black; border-collapse : collapse;padding: 10px;text-align:center;'}])
    dataframe=dataframe.hide(axis='index')
    msg.HTMLBody=html_body.format(dataframe.to_html(classes = 'table table-stripped',index=False),sender)
    msg.Save()
    msg.Send()

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
            raise CustomException ("Please Browse for the Excel File to continue")
        
        elif (len(workbook) > 0):
            excel=pd.ExcelFile(workbook)
            daily_plan_sheet=pd.read_excel(excel,'Planning Sheet')
            daily_plan_sheet.fillna("NA",inplace=True)
            input_error = []

            tomorrow=datetime.now()+timedelta(1) # getting tomorrow date for data execution
            daily_plan_sheet['Execution Date'] = pd.to_datetime(daily_plan_sheet['Execution Date'])
            for i in range(0,len(daily_plan_sheet)):
                if (daily_plan_sheet.iloc[i]['Execution Date'].strftime('%Y-%m-%d') != tomorrow.strftime('%Y-%m-%d')):
                    input_error.append(str(daily_plan_sheet.iloc[i]['S.NO']))

            daily_plan_sheet=daily_plan_sheet[daily_plan_sheet['Execution Date'] == tomorrow.strftime('%Y-%m-%d')]

            if len(daily_plan_sheet)==0:
                raise CustomException(f"Today's Maintenance Data not Found in the {workbook}, kindly check!")
            
            elif (len(input_error) > 0):
                raise CustomException(f"All the CR's present are not of Today's Maintenace Date for S.NO : {', '.join(input_error)}")
            
            else:
                flag = "Unsuccessful"
                
                Email_ID=pd.read_excel(excel,'Mail Id')

                # daily_plan_sheet['Circle'] = daily_plan_sheet['Circle'].str.upper()
                # for i in range(0,len(daily_plan_sheet)):
                #     daily_plan_sheet.at[i,'Circle'] = daily_plan_sheet.at[i,'Circle'].str.upper()
                daily_plan_sheet['Circle'].str.upper()
                daily_plan_sheet = daily_plan_sheet[['S.NO','Execution Date','Maintenance Window','CR NO','Activity Title','Risk','Location','Circle','Planning Status']]
                daily_plan_sheet = daily_plan_sheet[daily_plan_sheet['Planning Status'].str.upper() == 'PLANNED']
                input_error = []
                result_df = pd.DataFrame()
                circles=list(daily_plan_sheet['Circle'].unique())

                for i in circles:
                    if (i == 'NA'):
                        circles.remove(i)
                total_circles_in_planning_sheet = len(circles)
                
                mail_circles = Email_ID['Circle'].unique()
                for i in range(0,len(daily_plan_sheet)):
                    if (daily_plan_sheet.at[i,'CR NO'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    
                    if (daily_plan_sheet.at[i,'Activity Title'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    
                    if (daily_plan_sheet.at[i,'Circle'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Circle'] not in mail_circles):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                    else:
                        result_df = pd.concat([result_df,daily_plan_sheet.iloc[i].to_frame().T],ignore_index = True)
                
                
                #print(len(circles))
                #total_circles_in_planning_sheet = len(circles)

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
                                input_error.append(temp_df.at[i,'S.NO'])
                            continue
                        if (len(temp_df['Circle'].unique()) == 1):
                            new_result_df = pd.concat([new_result_df, temp_df.iloc[0].to_frame().T],ignore_index = True)
                    else:
                        new_result_df = pd.concat([new_result_df,temp_df.iloc[0].to_frame().T],ignore_index = True)
                
                daily_plan_sheet = new_result_df.copy(deep = True)
                
                del new_result_df
                del result_df
                #print(f"\n{daily_plan_sheet}\n")

                input_error = list(set(input_error))
                input_error.sort()

                if len(input_error)>0:
                        flag = "Unsuccessful"
                        messagebox.showwarning("  Input Error Detected",f"Input Error in Planning Sheet for S.NO : {', '.join(str(num) for num in input_error)}")
                        return flag
                
                else:
                    remainder=list(set(circles)-set(mail_circles))
                    circles=list(set(circles)-set(remainder))
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
                        # dataframe['Execution Date']=pd.to_datetime(dataframe['Execution Date'])
                        dataframe['Execution Date'] = pd.to_datetime(dataframe['Execution Date'],format = '%m/%d/%Y')
                        dataframe['Execution Date'] = dataframe['Execution Date'].dt.strftime('%m/%d/%Y')

                        dataframe.replace(to_replace = 'NA',value = '')
                    
                    
                        cir=circles[i]

                        for i in range(0,len(Email_ID)):
                            if (Email_ID.at[i,'Circle'] == cir):
                                row_to_fetch = i

                        to=Email_ID.iloc[row_to_fetch]['To Mail List']
                        cc=Email_ID.iloc[row_to_fetch]['Copy Mail List']
                        
                        subject=f"Connected End Nodes and their services on MPBN devices: {cir}_{tomorrow.strftime('%d-%m-%Y')}"
                        body="""
                            <html>        
                                <body>
                                    <div><p>Hi team,<br></p>
                                        <p>Please confirm below points so that we will approve CR’s.<br></p>
                                        <p>1)  End nodes and service details are required which are running on respective MPBN device (In case of changes on Core/STP/DRA/PACO/HLR connected MPBN nodes).</p>
                                        <p>2)  Design Maker & Checker confirmation mail need to be shared for all planned activity on Core/STP/DRA/PACO/HLR connected MPBN nodes.</p>
                                        <p>3)  KPI & Tester details need to be shared for all impacted nodes in Level-1 CR’s (SA).Also same details need to be shared for all Level-2 CR’s (NSA) with respect to changes on Core/STP/DRA/PACO/HLR conneted MPBN nodes.<br><br></p>
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
                    
                    flag = "Successful"
                    messagebox.showinfo("  Mail Sent Successfully",f"All Mails for mentioned planned {total_circles_in_planning_sheet} Circles in Daily Planning Sheet have been sent!")
                
                return flag

    except FileNotFoundError:
        working_directory=workbook
        messagebox.showwarning("    File not Found","Check {} for the Planning Sheet".format(working_directory))
        return "Unsuccessful"
    
    except ValueError:
         working_directory=workbook
         messagebox.showwarning(" Value Error"," Check {} for all the requirement sheets".format(working_directory))
         return "Unsuccessful"
    
    except CustomException as error:
        messagebox.showerror("  Data can't be found",error)
        return "Unsuccessful"
    
    # except Exception as e:
    #     messagebox.showerror("  Exception Occurred",e)
    #     return "Unsuccessful"
    
#fetch_details("Enjoy Maity",r"C:\Users\emaienj\Downloads\MPBN Daily Planning Sheet new copy.xlsx")