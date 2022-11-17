import sys
from datetime import datetime,timedelta
import subprocess
import pandas as pd
import win32com.client as win32
from openpyxl import load_workbook
from tkinter import messagebox

class TomorrowDataNotFound(Exception):
    def __init__(self,msg):
        self.msg=msg

def sendmail(dataframe,to,cc,body,subject,north_and_west_region,east_and_south_region,sender):
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

    dataframe=dataframe.hide(axis='index') # hiding the index coloumn
    msg.HTMLBody=html_body.format(north_and_west_region,east_and_south_region,dataframe.to_html(index=False),sender)
    msg.Save()
    msg.Send()

    #messagebox.showinfo("   Successful Completion","Interdomain KPIs Mail Communication Task completed")

def quit(event):
    sys.exit(0)
    
def paco_cscore(sender,workbook,north_and_west_region,east_region_and_south_region):
    try: 
        #   user=subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 

        #workbook=r"C:\Daily\MPBN Daily Planning Sheet.xlsx" # system path from where the program will take the input
        daily_plan_sheet=pd.read_excel(workbook,'Planning Sheet')
        daily_plan_sheet.fillna("NA",inplace=True)
        input_error = []
        tomorrow=datetime.now()+timedelta(1) # getting tomorrow date for data execution
            
        for i in range(0,len(daily_plan_sheet)):
            if (daily_plan_sheet.iloc[i]['Execution Date'].strftime('%Y-%m-%d') != tomorrow.strftime('%Y-%m-%d')):
                input_error.append(str(daily_plan_sheet.iloc[i]['S.NO']))
        #print(daily_plan_sheet.iloc[2]['Execution Date'])
        daily_plan_sheet=daily_plan_sheet[daily_plan_sheet['Execution Date']==tomorrow.strftime('%Y-%m-%d')]
        

        if len(daily_plan_sheet)==0:
                raise TomorrowDataNotFound("Data for tomorrow's date is not present in the MPBN Daily Planning Sheet, kindly check!")
        
        elif (len(input_error) > 0):
                raise TomorrowDataNotFound(f"All the CR's present are not of Today's Maintenace Date for S.NO : {', '.join(input_error)}")    
            
        else:
            #daily_plan_sheet=daily_plan_sheet.drop_duplicates(subset=['CR NO'])
            Email_Id=pd.read_excel(workbook,'Mail Id')
            list_of_interdomains=["CS Core","PS Core","RAN"]
            df2=pd.read_excel(workbook,sheet_name="CS Core-Inter Domain")
            df2.fillna("NA",inplace=True)
            df=pd.read_excel(workbook,sheet_name="PS Core-Inter Domain")
            df.fillna("NA",inplace=True)
            df3=pd.read_excel(workbook,sheet_name="RAN-Inter Domain")
            df3.fillna("NA",inplace=True)

            suffix=["st","nd","rd","th"]
            date_end_digit=int(tomorrow.strftime("%d"))%10
            date_digits=int(tomorrow.strftime("%d"))%100
            if date_digits<10 or date_digits>20:
                if date_end_digit==1:
                    suffix_for_date=suffix[0]
                elif date_end_digit==2:
                    suffix_for_date=suffix[1]
                elif date_end_digit==3:
                    suffix_for_date=suffix[2]
                else:
                    suffix_for_date=suffix[3]
            else:
                suffix_for_date=suffix[3]
            for_date=tomorrow.strftime("%d{}_%b'%y").format(suffix_for_date)

            
            list_of_dfs=[df2,df,df3]

            for i in list_of_interdomains:
                subject=f"ONLY FOR TESTING: KPI Monitoring | {i} for MPBN CRs | {for_date}"
                if i=="CS Core":
                    to=Email_Id.iloc[24]['To Mail List']
                    cc=Email_Id.iloc[24]['Copy Mail List']
                    dataframe=df2
                elif i=="PS Core":
                    to=Email_Id.iloc[23]['To Mail List']
                    cc=Email_Id.iloc[23]['Copy Mail List']
                    dataframe=df
                elif i=="RAN":
                    to=Email_Id.iloc[25]['To Mail List']
                    cc=Email_Id.iloc[25]['Copy Mail List']
                    dataframe=df3

                mpbn_html_body="""
                    <html>
                        <body>
                            <div>
                                    <p>Hi Team,</p>
                                    <p>Please find below the list of MPBN activity which includes Core nodes, so KPI monitoring required. Impacted nodes with KPI details given below. Please share KPI monitoring resource from your end.<br><br></p>
                                    <p>@Core Team: Please contact below spoc region wise if any issue with KPI input.<br><br></p>
                                    <p>{}: North region and west region </p>
                                    <p>{}: East region and South region <br></p>
                                    <p>Note:-If there is any deviation in KPI please call to Executer before 6 AM. After that please call to technical validator/Team Lead.<br><br></p>
                            
                            </div>
                            <div>
                                {}
                            </div>
                            <div>
                                    <p>With Regards<br>{}<br>Ericsson India Global Services Pvt. Ltd.</p>
                            </div>
                        </body>
                    </html>
                """
                sendmail(dataframe,to,cc,mpbn_html_body,subject,north_and_west_region,east_region_and_south_region,sender)
                messagebox.showinfo("     Mail Sent Info",f"Mail sent for {i} Interdomain KPIs")
            
        messagebox.showinfo(" Mail Sent Successfully","Mail For the All the Interdomain Kpis Have Been Sent")


   

    except FileNotFoundError:
        working_directory=r"C:\Daily"
        messagebox.showerror(" File not Found","Check {} for MPBN Daily Planning Sheet.xlsx".format(working_directory)).bind("<Return>",quit)
        sys.exit(0)
    
    except ValueError:
         working_directory=r"C:\Daily"
         messagebox.showwarning("    Value Error","Check {} for MPBN Daily Planning Sheet.xlsx for all the requirement sheet".format(working_directory)).bind("<Return>",quit)
         sys.exit(0)
    

    except TomorrowDataNotFound as error:
        messagebox.showerror(" Data for tomorrow's date not found",error).bind("<Return>",quit)
        sys.exit(0)
    
   