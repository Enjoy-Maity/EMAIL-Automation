import sys                                                          # Importing the sys to run cmd commands from the script itself.
from openpyxl import load_workbook                                  # Importing load_workbook class from the openpyxl to load existing excel workbook.
from openpyxl.styles import Font,Border,Side,PatternFill,Alignment  # Importing classes from openpyxl to style the excel workbooks.
from openpyxl import Workbook                                       # Importing Workbook to Create Workbook using openpyxl.
from openpyxl.utils import get_column_letter                        # Importing the get_column_letter from openpyxl to convert the column numbers to alphabet letter used in the excel sheet.
import pandas as pd                                                 # Importing Pandas to manipulate the data from the excel sheet.
from datetime import datetime,timedelta                             # Importing datetime and timedelta to get today's maintenance date based on system's current date and time settings.
from tkinter import *                                               # Importing all the classes from tkinter GUI Module of python.
from tkinter import messagebox                                      # Importing Messagebox to invoke messages where required.
from pathlib import Path                                            # Importing Path from pathlib to check the existence of a file.

# Creating Custom classes to handle custom defined Exceptions (Interuptions) to handle the flow of program.
class TomorrowDataNotFound(Exception):
    def __init__(self,msg):
        self.msg = msg

class CustomException(Exception):
    def __init__(self,title,msg):
        self.title = title
        self.msg = msg
        messagebox.showerror(self.title,self.msg)


#####################################################################
#########################  P1 P3 appender  ##########################
#####################################################################

def p_one_p_three_appender(sender,workbook):
    # Getting the email-package-sheet
    wb              = pd.ExcelFile(workbook)
    email_package   = pd.read_excel(wb,"Email-Package")
    planning_sheet  = pd.read_excel(wb,"Planning Sheet")

    planning_sheet = planning_sheet[planning_sheet['Planning Status'] == 'Discussed']
    planning_sheet['Execution Date'] = pd.to_datetime(planning_sheet['Execution Date'], format="%d-%b-%Y")
    planning_sheet['Execution Date'] = planning_sheet['Execution Date'].dt.strftime("%m/%d/%Y")

    email_package = pd.concat([email_package,planning_sheet],ignore_index=True)
    
    # Getting the unique technical validator.
    unique_technical_validator_set = set(email_package['Technical Validator'].unique())
    p1 = ''
    p3 = ''
    
    ''' 
        If the User is not a technical validator then we are throwing an Exception so that only the respective Technical Validator file 
        gets written out which are present in the Planning Sheet.
    '''
    
    if (sender not in unique_technical_validator_set):
        raise CustomException(' Technical Validator not Found!','Technical Validator is not found in the Planning Sheet, Kindly Check!')
    
    if ('Arka Maiti' in unique_technical_validator_set):
        p3 = 'Arka Maiti'
        unique_technical_validator_set.remove(p3)

    if ('Manoj Kumar' in unique_technical_validator_set):
        p1 = 'Manoj Kumar'
        unique_technical_validator_set.remove(p1)

    if ((len(p1) > 0) and (len(p3) == 0)):
        p3 = list(unique_technical_validator_set - set(p1))[0]
    
    if ((len(p3) > 0) and (len(p1) == 0)):
        p1 = list(unique_technical_validator_set - set(p3))[0]
    
    if (len(p1)>0 and len(p3)>0):
        if (sender == p1):
            # Here we are trying to get the parent folder path of the Workbook containing the Email_package sheet.
            file_path = workbook.split("/")
            file_path.remove(file_path[-1])
            file_path = '/'.join(file_path)
            p1_workbook_file = f'{file_path}/MPBN Planning Automation Tracker P1.xlsx'
            p1_sheet_name = 'MPBN Activity List'

            # Here we are filtering rows with the particular Technical Validator to write into the excel sheet.
            p1_dataframe = email_package[email_package['Technical Validator'] == p1]
            p1_dataframe.drop("S.NO",axis = "columns",inplace = True)

            p1_dataframe.replace('NA'," ",inplace = True)
            p1_columns = p1_dataframe.columns.to_list()
            
            # Finding out whether the file for MPBN Planning Automation Tracker P1.xlsx exists or not
            # If the File does not exists then in that case the file is created

            if (Path(p1_workbook_file).exists() == False):
                wb = Workbook()
                wb.create_sheet(index=0,title = p1_sheet_name)
                ws = wb[p1_sheet_name]
                ws['A1'] = 'S.NO'
                for i in range(0,len(p1_columns)):
                    col = get_column_letter(i+2)
                    ws[col+'1'] = p1_columns[i]
                wb.save(p1_workbook_file)

            # Loading the workbook to find the number of rows occupied in the worksheet to continue the S.NO series in that worksheet.
            p1_workbook = load_workbook(p1_workbook_file)

            # Changing the Index of the dataframe to start from 1
            p1_dataframe.reset_index(drop = True, inplace = True)
            p1_dataframe.index += (p1_workbook[p1_sheet_name].max_row)
            p1_dataframe.insert(0,'S.NO',p1_dataframe.index)
            
            # Reading the Excel file in pandas.
            p1_file_read = pd.ExcelFile(p1_workbook_file)
            p1_file_read = pd.read_excel(p1_file_read,p1_sheet_name)

            #Converting the execution date column values in the email_package to datetime datatype to execute the further operations
            p1_file_read['Execution Date'] = pd.to_datetime(p1_file_read['Execution Date'],format= "%M/%d/%Y")
            p1_file_read['Execution Date'] = p1_file_read['Execution Date'].dt.strftime("%M/%d/%Y")

            # Getting the unique Execution Date from the Execution Date Column of the MPBN Planning Automation Tracker
            p1_file_read_unique_execution_date = list(p1_file_read['Execution Date'].unique())
            
            # Assigning a Variable to get the today's maintenance date to check whether today's maintenance date's data is present in the MPBN Planning Automation Tracker
            todays_maintenance_date = email_package.iloc[1]['Execution Date']
            
            ''' 
                In this condition we are trying to check whether today's maintenance date is present in the MPBN Planning Automation Tracker Workbook's 
                MPBN Activity List 
            '''
            if (todays_maintenance_date not in p1_file_read_unique_execution_date):
                writer1 = pd.ExcelWriter(p1_workbook_file, engine = 'openpyxl', mode = 'a', if_sheet_exists = 'overlay')
                p1_dataframe.to_excel(writer1,p1_sheet_name,startrow = p1_workbook[p1_sheet_name].max_row, index = False,index_label = 'S.NO',header = False)
                writer1.close()
                
                # Styling the worksheet.
                styling(p1_workbook_file,p1_sheet_name)
                
                # message showing MPBN Planning Automation Tracker Status is successfully edited.
                messagebox.showinfo("   MPBN Planning Automation Tracker Status",f"All planned CRs for Validator {sender} has been updated in MPBN Planning Automation Tracker!")
        
            else:
                # Message showing that the data for today's maintenance date is already present in the MPBN Planning Automation Tracker Status Excel worksheet.
                messagebox.showinfo("   Data already present","Data for today's maintenance date is already present in the MPBN Planning Automation Tracker")
        
            
                
            
        if (sender == p3):
            # Here we are trying to get the parent folder path of the Workbook containing the Email_package sheet.
            file_path = workbook.split("/")
            file_path.remove(file_path[-1])
            file_path = "/".join(file_path)
            p3_workbook_file = f'{file_path}/MPBN Planning Automation Tracker P3.xlsx'
            p3_sheet_name = 'MPBN Activity List'

            # Here we are filtering rows with the particular Technical Validator to write into the excel sheet.
            p3_dataframe = email_package[email_package['Technical Validator'] == p3]
            p3_dataframe.reset_index(drop = True, inplace = True)
            p3_dataframe.drop("S.NO",axis = "columns", inplace = True)
            p3_dataframe.replace('NA'," ",inplace = True)
            p3_columns = p3_dataframe.columns.to_list()

            # Finding out whether the file for MPBN Planning Automation Tracker P3.xlsx exists or not
            # If the File does not exists then in that case the file is created
            if (Path(p3_workbook_file).exists() == False):
                wb = Workbook()
                wb.create_sheet(index=0,title = p3_sheet_name)
                ws = wb[p3_sheet_name]
                ws['A1'] = 'S.NO'
                for i in range(0,len(p3_columns)):
                    col = get_column_letter(i+2)
                    ws[col+'1'] = p3_columns[i]
                wb.save(p3_workbook_file)
            
            # Loading the workbook to find the number of rows occupied in the worksheet to continue the S.NO series in that worksheet.
            p3_workbook = load_workbook(p3_workbook_file)

            # Changing the Index of the dataframe to start from 1
            p3_dataframe.index += (p3_workbook[p3_sheet_name].max_row)
            p3_dataframe.insert(0,'S.NO',p3_dataframe.index)

            # Reading the Excel sheet in pandas.
            p3_file_read = pd.ExcelFile(p3_workbook_file)
            p3_file_read = pd.read_excel(p3_file_read, p3_sheet_name)

            # Formatting the 'Execution Date' to pandas datetime datatype for further usage in the program.
            p3_file_read['Execution Date'] = pd.to_datetime(p3_file_read['Execution Date'],format="%M/%d/%Y")
            p3_file_read['Execution Date'] = p3_file_read['Execution Date'].dt.strftime("%M/%d/%Y")

            # Getting the unique 'Execution Date' from the Execution Date Column of the MPBN Planning Automation Tracker.
            p3_file_read_unique_execution_date = list(p3_file_read['Execution Date'].unique())
            
            # Assigning a Variable to get the today's maintenance date to check whether today's maintenance date's data is present in the MPBN Planning Automation Tracker
            todays_maintenance_date = email_package.iloc[1]['Execution Date']
            
            ''' 
                In this condition we are trying to check whether today's maintenance date is present in the MPBN Planning Automation Tracker Workbook's 
                MPBN Activity List 
            '''
            if (todays_maintenance_date not in p3_file_read_unique_execution_date):
                writer3 = pd.ExcelWriter(p3_workbook_file, engine = 'openpyxl', mode = 'a', if_sheet_exists = 'overlay')
                p3_dataframe.to_excel(writer3,p3_sheet_name,startrow = p3_workbook[p3_sheet_name].max_row, index = False,index_label = 'S.NO', header = False)
                writer3.close()

                # Styling the worksheet.
                styling(p3_workbook_file,p3_sheet_name)

                # message showing MPBN Planning Automation Tracker Status is successfully edited.
                messagebox.showinfo("   MPBN Planning Automation Tracker Status",f"All planned CRs for Validator '{sender}' has been updated in MPBN Planning Automation Tracker!")
            
            else:
                # Message showing that the data for today's maintenance date is already present in the MPBN Planning Automation Tracker Status Excel worksheet.
                messagebox.showinfo("   Data already present","Data for today's mainteance date is already present in the MPBN Planning Automation Tracker")

    else:
        # Message showing that the user who is running the application is not a technical validator.
        messagebox.showinfo("   Technical Validator Name Mismatch!",f"{sender}'s name is not matching with Technical Validator")
        return "Unsuccessful"


#####################################################################
#############################    Styling   ##########################
#####################################################################

# Method(Function) for styling the worksheet.
def styling(workbook,sheetname):
    wb  =  load_workbook(workbook)                          # loading the workbook.
    ws  =  wb[sheetname]                                    # loading the worksheet.
    font_style  =  Font(color = "FFFFFF",bold = True)       # Setting the font style with color.
    col_widths = []                                         # Empty list for entering the max length of string in each column.

    # Iterating through the row values to find the max length of string in each column in the row and appending it to the col_widths list

    for row_values in ws.iter_rows(values_only = True):
        for j,value in enumerate(row_values):
            if len(col_widths)>j:
                if col_widths[j] < len(str(value)):
                    col_widths[j] = len(str(value))
            else:
                col_widths.insert(j,len(str(value)))

    # Standardising the length of each column in the sheet.

    for i,column_width in enumerate(col_widths,1):
        if column_width <= 47:
            ws.column_dimensions[get_column_letter(i)].width = column_width+3
        else:
            ws.column_dimensions[get_column_letter(i)].width = 50


    # Coloring the headers and alingning the headers text to center both horizontally and vertically.
    for column in range(1,ws.max_column+1):   # ws.max_column returns the total number of columns present
        col = get_column_letter(column)
        color_fill = PatternFill(start_color = '0033CC',end_color = '0033CC',fill_type = 'solid')
        ws[col+'1'].font = font_style
        ws[col+'1'].fill = color_fill
        ws[col+'1'].alignment = Alignment(horizontal = 'center',vertical = 'center')

    # Styling all the occupied cells in the sheet, by adding border to the cells, aligning the text in the center
    
    for row in ws:
        for cell in row:
            cell.alignment = Alignment(horizontal = 'center',vertical = 'center',wrap_text=False)
            cell.border = Border(top = Side(border_style = 'medium',color = '000000'),bottom = Side(border_style = 'medium',color = '000000'),left = Side(border_style = 'medium',color = '000000'),right = Side(border_style = 'medium',color = '000000'))

    # Saving the workbook with worksheet with all the changes.
    wb.save(workbook)

# Method(Function) for quitting the application.
def quit(event):
    sys.exit(0)

#####################################################################
#############################  Paco_cscore  #########################
#####################################################################

def paco_cscore(sender,workbook):
    try:
        #user = subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 

        #workbook = r"C:\Daily\MPBN Daily Planning Sheet.xlsx" # system path from where the program will take the input
        
        daily_plan_sheet = pd.read_excel(workbook,'Planning Sheet')
        tomorrow = datetime.today()+timedelta(1) # getting tomorrow date for data execution
        difference = []
        daily_plan_sheet['Execution Date'] = pd.to_datetime(daily_plan_sheet['Execution Date'])
        
        if (len(daily_plan_sheet) == 0):
            raise TomorrowDataNotFound("The Planning Sheet is empty! Kindly Check!")

        for i in range(0,len(daily_plan_sheet)):
            if (daily_plan_sheet.iloc[i]['Execution Date'].strftime('%Y-%m-%d') != tomorrow.strftime('%Y-%m-%d')):
                difference.append(str(daily_plan_sheet.iloc[i]['S.NO']))
        
        daily_plan_sheet['Execution Date'] = daily_plan_sheet['Execution Date'].dt.strftime("%m/%d/%Y")
        daily_plan_sheet = daily_plan_sheet[daily_plan_sheet['Execution Date'] == tomorrow.strftime("%m/%d/%Y")]
        
        if (len(daily_plan_sheet) == 0):
            raise TomorrowDataNotFound("Data for today's maintenance date is not present in the MPBN Daily Planning Sheet, kindly check!")
        
        if (len(difference) > 0):
            raise TomorrowDataNotFound(f"All the CR's present are not of Today's Maintenace Date for S.NO : {', '.join([str(num) for num in difference])}")
        
        else:
            
            daily_plan_sheet = daily_plan_sheet[daily_plan_sheet['Planning Status'].str.upper() == 'PLANNED']
            Email_ID = pd.read_excel(workbook,"Mail Id")
            
            # Finding the Circles and Change Responsible available in the Mail ID worksheet of the MPBN Daily Planning workbook.
            circle = Email_ID['Circle'].tolist()
            original_change_responsible = list(Email_ID['Change Responsible'].unique())

            # Changing the case of each original change responsible to upper.
            for i in range(0,len(original_change_responsible)):
                original_change_responsible[i] = str(original_change_responsible[i]).strip().upper()
                
            # Removing the NAN value from the list.
            original_change_responsible.remove('NAN')

            # Creating an empty list and empty dataframe to append the S.NO. of rows with input errors and creating a new dataframe from the daily_plan_sheet dataframe with only required data(rows).
            input_error = []
            result_df = pd.DataFrame()
            
            # Replacing all the blank fields(excel cells) in the dataframe with 'NA'
            daily_plan_sheet.fillna("NA",inplace = True)

            # Creating empty list to find out the serial numbers of the rows where the Circle input and the Change responsible is not properly entered by the user.
            circle_not_proper = []
            change_responsible_not_proper = []

            # Iterating (Looping) through the daily_plan_sheet dataframe index wise (index given by pandas to each row with data), to find out the serial 
            # numbers of the rows where the Circle input and the Change responsible is not properly entered by the user and any other fields that should be left unblank
            # by the user.
            for i in range(0,len(daily_plan_sheet)):
                if (daily_plan_sheet.iloc[i]['CR NO'] == "NA") or (daily_plan_sheet.iloc[i]['CR NO'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Circle'] not in circle):
                    circle_not_proper.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Change Responsible'].strip().upper() not in original_change_responsible):
                    change_responsible_not_proper.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Activity Title'] == 'NA') or (daily_plan_sheet.iloc[i]['Activity Title'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Circle'] == 'NA') or (daily_plan_sheet.iloc[i]['Circle'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Risk'] == 'NA') or (daily_plan_sheet.iloc[i]['Risk'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Location'] == 'NA') or (daily_plan_sheet.iloc[i]['Location'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Change Responsible'] == 'NA') or (daily_plan_sheet.iloc[i]['Change Responsible'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Impact'] == 'NA') or (daily_plan_sheet.iloc[i]['Impact'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Technical Validator'] == 'NA') or (daily_plan_sheet.iloc[i]['Technical Validator'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Activity-Type'] == 'NA') or (daily_plan_sheet.iloc[i]['Activity-Type'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Vendor'] == 'NA') or (daily_plan_sheet.iloc[i]['Vendor'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Protocol'] == 'NA') or (daily_plan_sheet.iloc[i]['Protocol'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                if (daily_plan_sheet.iloc[i]['Execution Projection'] == 'NA') or (daily_plan_sheet.iloc[i]['Execution Projection'] == None):
                    input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    continue
                else:
                    result_df = pd.concat([result_df,daily_plan_sheet.iloc[i].to_frame().T], ignore_index= True)
            
            result_df.drop_duplicates(keep = 'first', inplace= True)

            # Deleting the old daily_plan_sheet dataframe which we won't be using, to free up memory space.
            del daily_plan_sheet
            daily_plan_sheet = result_df.copy(deep = True)
            # Deleting the result_df after creating a deep copy of it and assigning variable daily_plan_sheet to that deep copy.
            del result_df
            
            # Sorting the input error indices to get the list of input error in ascending order.
            input_error.sort()

            if (len(input_error) > 0):
                messagebox.showerror("  Input Errors",f"Input Error in Planning Sheet! Check 'Location', 'Circle', 'Change Responsible', 'Impact', 'Vendor', 'Protocol' & 'Execution Projection' for S.NO.: {','.join([str(num) for num in input_error])}")
                return 'Unsuccessful'
            if (len(circle_not_proper) > 0):
                messagebox.showerror("  Circles Errors",f"Input Circles are wrong in Planning Sheet for S.NO. : {','.join([str(num) for num in circle_not_proper])}")
                return 'Unsuccessful'
            if (len(change_responsible_not_proper) > 0):
                messagebox.showerror("  Change Responsible Errors",f"Input Change Responsible are wrong in Planning Sheet for S.NO.: {','.join([str(num) for num in change_responsible_not_proper])}")
                return 'Unsuccessful'
            else:
            
                sheetname = "PS Core-Inter Domain"
                sheetname2 = "CS Core-Inter Domain"
                sheetname3 = "RAN-Inter Domain"
                sheetname4 = "VAS-Inter Domain"

                category = "MPBN-MS"
                owner_domain = "SRF MPBN"
                team_leader = "Karan Loomba"

                ####################################################### Entering details for ps core or paco circle ###########################################################
                execution_date = []
                maintenance_window = []
                mpbn_cr_no = []
                location = []
                mpbn_change_responsible_executor = []
                validator = []
                impact = []
                circle = []
                mpbn_activity_title = []
                cr_owner_domain = []
                inter_domain = []
                cr_category = []
                impacted_node_details = []
                Kpis_to_be_monitored = []
                # Execution Date	Maintenance Window	MPBN CR NO	CR Category	Impact	Location	Circle	MPBN Activity Title	CR Owner Domain	MPBN Change Responsible	Technical Validator/Team Lead	InterDomain	Impacted Node Details	KPI's to be monitored
                for i in range(0,len(daily_plan_sheet)):
                    if ((daily_plan_sheet.iloc[i]['Domain kpi'].upper() == 'PS-CORE') or (daily_plan_sheet.iloc[i]['Domain kpi'].upper() == 'PS') or (daily_plan_sheet.iloc[i]['Domain kpi'].upper() == 'PS_CORE') or (daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith('PACO')) or (daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith("PS")) and (daily_plan_sheet.iloc[i]['Planning Status'].upper() == 'PLANNED')):
                        execution_date.append(daily_plan_sheet.iloc[i]['Execution Date'])
                        maintenance_window.append(daily_plan_sheet.iloc[i]['Maintenance Window'])
                        mpbn_cr_no.append(daily_plan_sheet.iloc[i]['CR NO'])
                        cr_category.append(category)
                        impact.append(daily_plan_sheet.iloc[i]['Impact'])
                        location.append(daily_plan_sheet.iloc[i]['Location'])
                        txt = str(daily_plan_sheet.iloc[i]['Circle'])
                        circle.append(txt.upper())
                        mpbn_activity_title.append(daily_plan_sheet.iloc[i]['Activity Title'])
                        cr_owner_domain.append(owner_domain)
                        mpbn_change_responsible_executor.append(daily_plan_sheet.iloc[i]['Change Responsible'])
                        technical_validator = daily_plan_sheet.iloc[i]['Technical Validator']
                        if technical_validator == team_leader:
                            validator.append(team_leader)
                        else:
                            tech_validator_team_leader = technical_validator+"/"+team_leader
                            validator.append(tech_validator_team_leader)
                        inter_domain.append(daily_plan_sheet.iloc[i]['Domain kpi'].upper())
                        impacted_node_details.append(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
                        Kpis_to_be_monitored.append(daily_plan_sheet.iloc[i]['KPI DETAILS'])

                dictionary1 = {'CR':mpbn_cr_no,'Maintenance Window':maintenance_window,'CR Category':cr_category,'Impact':impact,'Location':location,'Circle':circle,'MPBN Activity Title':mpbn_activity_title,'CR Owner Domain':cr_owner_domain,'Change Responsible':mpbn_change_responsible_executor,'Technical Validator/Team Lead':validator,'InterDomain':inter_domain,'Impacted Node Details':impacted_node_details,'KPIs to be monitored':Kpis_to_be_monitored}
                df = pd.DataFrame(dictionary1)
                df.drop_duplicates(subset = 'CR',keep = "first", inplace = True)

                ######################################################### Entering details for Cs core #######################################################################
                execution_date = []
                maintenance_window = []
                mpbn_cr_no = []
                location = []
                mpbn_change_responsible_executor = []
                validator = []
                impact = []
                circle = []
                mpbn_activity_title = []
                cr_owner_domain = []
                inter_domain = []
                cr_category = []
                impacted_node_details = []
                Kpis_to_be_monitored = []
                for i in range(0,len(daily_plan_sheet)):
                    if ((daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith("CS")) or (daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith("STP")) or (daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith("CORE")) and (daily_plan_sheet.iloc[i]['Planning Status'].upper() == 'PLANNED')) :
                        execution_date.append((daily_plan_sheet.iloc[i]['Execution Date']))
                        maintenance_window.append(daily_plan_sheet.iloc[i]['Maintenance Window'])
                        mpbn_cr_no.append(daily_plan_sheet.iloc[i]['CR NO'])
                        cr_category.append(category)
                        impact.append(daily_plan_sheet.iloc[i]['Impact'])
                        location.append(daily_plan_sheet.iloc[i]['Location'])
                        txt = str(daily_plan_sheet.iloc[i]['Circle'])
                        circle.append(txt.upper())
                        mpbn_activity_title.append(daily_plan_sheet.iloc[i]['Activity Title'])
                        cr_owner_domain.append(owner_domain)
                        mpbn_change_responsible_executor.append(daily_plan_sheet.iloc[i]['Change Responsible'])
                        technical_validator = daily_plan_sheet.iloc[i]['Technical Validator']
                        if technical_validator == team_leader:
                            validator.append(team_leader)
                        else:
                            tech_validator_team_leader = technical_validator+"/"+team_leader
                            validator.append(tech_validator_team_leader)
                        inter_domain.append(daily_plan_sheet.iloc[i]['Domain kpi'].upper())
                        impacted_node_details.append(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
                        Kpis_to_be_monitored.append(daily_plan_sheet.iloc[i]['KPI DETAILS'])
                dictionary2 = {'CR':mpbn_cr_no,'Maintenance Window':maintenance_window,'CR Category':cr_category,'Impact':impact,'Location':location,'Circle':circle,'MPBN Activity Title':mpbn_activity_title,'CR Owner Domain':cr_owner_domain,'Change Responsible':mpbn_change_responsible_executor,'Technical Validator/Team Lead':validator,'InterDomain':inter_domain,'Impacted Node Details':impacted_node_details,'KPIs to be monitored':Kpis_to_be_monitored}
                df2 = pd.DataFrame(dictionary2)
                df2.drop_duplicates(subset = 'CR',keep = "first", inplace = True)

                ##########################################################  Entering details for RAN  ########################################################################
                execution_date = []
                maintenance_window = []
                mpbn_cr_no = []
                location = []
                mpbn_change_responsible_executor = []
                validator = []
                impact = []
                circle = []
                mpbn_activity_title = []
                cr_owner_domain = []
                inter_domain = []
                cr_category = []
                impacted_node_details = []
                Kpis_to_be_monitored = []
                oss_name = []
                oss_IP = []
                # Execution Date	Maintenance Window	MPBN CR NO	CR Category	Impact	Location	Circle	MPBN Activity Title	CR Owner Domain	MPBN Change Responsible	Technical Validator/Team Lead	InterDomain	Impacted Node Details	KPI's to be monitored
                for i in range(0,len(daily_plan_sheet)):
                    if ((daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith("RAN") and (daily_plan_sheet.iloc[i]['Planning Status'].upper() == 'PLANNED'))):
                        execution_date.append(daily_plan_sheet.iloc[i]['Execution Date'])
                        maintenance_window.append(daily_plan_sheet.iloc[i]['Maintenance Window'])
                        mpbn_cr_no.append(daily_plan_sheet.iloc[i]['CR NO'])
                        cr_category.append(category)
                        impact.append(daily_plan_sheet.iloc[i]['Impact'])
                        location.append(daily_plan_sheet.iloc[i]['Location'])
                        txt = str(daily_plan_sheet.iloc[i]['Circle'])
                        circle.append(txt.upper())
                        mpbn_activity_title.append(daily_plan_sheet.iloc[i]['Activity Title'])
                        cr_owner_domain.append(owner_domain)
                        mpbn_change_responsible_executor.append(daily_plan_sheet.iloc[i]['Change Responsible'])
                        technical_validator = daily_plan_sheet.iloc[i]['Technical Validator']
                        if technical_validator == team_leader:
                            validator.append(team_leader)
                        else:
                            tech_validator_team_leader = technical_validator+"/"+team_leader
                            validator.append(tech_validator_team_leader)
                        inter_domain.append(daily_plan_sheet.iloc[i]['Domain kpi'])
                        impacted_node_details.append(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
                        Kpis_to_be_monitored.append(daily_plan_sheet.iloc[i]['KPI DETAILS'])
                        oss_name.append(daily_plan_sheet.iloc[i]['oss name'])
                        oss_IP.append(daily_plan_sheet.iloc[i]['oss ip'])

                dictionary3 = {'CR':mpbn_cr_no,'Maintenance Window':maintenance_window,'CR Category':cr_category,'Impact':impact,'Location':location,'Circle':circle,'MPBN Activity Title':mpbn_activity_title,'CR Owner Domain':cr_owner_domain,'Change Responsible':mpbn_change_responsible_executor,'Technical Validator/Team Lead':validator,'InterDomain':inter_domain,'Impacted Node Details':impacted_node_details,'KPIs to be monitored':Kpis_to_be_monitored,'OSS Name':oss_name,'OSS IP':oss_IP}
                df3 = pd.DataFrame(dictionary3)
                df3.drop_duplicates(subset = 'CR',keep = "first", inplace = True)

                ##########################################################  Entering details for VAS  ########################################################################
                
                execution_date = []
                maintenance_window = []
                mpbn_cr_no = []
                location = []
                mpbn_change_responsible_executor = []
                validator = []
                impact = []
                circle = []
                mpbn_activity_title = []
                cr_owner_domain = []
                inter_domain = []
                cr_category = []
                impacted_node_details = []
                Kpis_to_be_monitored = []
                oss_name = []
                oss_IP = []
                for i in range(0,len(daily_plan_sheet)):
                    if ((daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith('VAS')) and (daily_plan_sheet.iloc[i]['Planning Status'].upper() == 'PLANNED')):
                        execution_date.append(daily_plan_sheet.iloc[i]['Execution Date'])
                        maintenance_window.append(daily_plan_sheet.iloc[i]['Maintenance Window'])
                        mpbn_cr_no.append(daily_plan_sheet.iloc[i]['CR NO'])
                        cr_category.append(category)
                        impact.append(daily_plan_sheet.iloc[i]['Impact'])
                        location.append(daily_plan_sheet.iloc[i]['Location'])
                        txt = str(daily_plan_sheet.iloc[i]['Circle'])
                        circle.append(txt.upper())
                        mpbn_activity_title.append(daily_plan_sheet.iloc[i]['Activity Title'])
                        cr_owner_domain.append(owner_domain)
                        mpbn_change_responsible_executor.append(daily_plan_sheet.iloc[i]['Change Responsible'])
                        technical_validator = daily_plan_sheet.iloc[i]['Technical Validator']
                        if technical_validator == team_leader:
                            validator.append(team_leader)
                        else:
                            tech_validator_team_leader = technical_validator+"/"+team_leader
                            validator.append(tech_validator_team_leader)
                        inter_domain.append(daily_plan_sheet.iloc[i]['Domain kpi'].upper())
                        impacted_node_details.append(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
                        Kpis_to_be_monitored.append(daily_plan_sheet.iloc[i]['KPI DETAILS'])

                dictionary4 = {'CR':mpbn_cr_no,'Maintenance Window':maintenance_window,'CR Category':cr_category,'Impact':impact,'Location':location,'Circle':circle,'MPBN Activity Title':mpbn_activity_title,'CR Owner Domain':cr_owner_domain,'Change Responsible':mpbn_change_responsible_executor,'Technical Validator/Team Lead':validator,'InterDomain':inter_domain,'Impacted Node Details':impacted_node_details,'KPIs to be monitored':Kpis_to_be_monitored}
                df4 = pd.DataFrame(dictionary4)
                df4.drop_duplicates(subset = 'CR', keep = "first", inplace = True)

                # Dropping the Index of each Dataframe so that they're not written into the excel sheets.
                df.reset_index(drop = True,inplace = True)
                df2.reset_index(drop = True,inplace = True)
                df3.reset_index(drop = True,inplace = True)
                df4.reset_index(drop = True,inplace = True)

               
                # writer = pd.ExcelWriter(workbook,engine = 'xlsxwriter')

                # daily_plan_sheet.to_excel(writer,sheet_name = 'Planning Sheet',index = False)
                # df.to_excel(writer,sheet_name = sheetname,index = False)
                # df2.to_excel(writer,sheet_name = sheetname2,index = False)
                # df3.to_excel(writer,sheet_name = sheetname3,index = False)
                # Email_Id.to_excel(writer,sheet_name = 'Mail Id',index = False)

                # Writing the dataframes into the worksheets.
                # Checking whether the interdomain data are present in the sheet or not.
                if((len(df) > 0) or (len(df2) > 0) or (len(df3) > 0) or (len(df4) > 0)):
                    writer = pd.ExcelWriter(workbook,engine = "openpyxl",mode = "a",if_sheet_exists = "replace")
                    if (len(df) > 0):
                        df.to_excel(writer,sheet_name = sheetname,index = False)
                    
                    if (len(df2) > 0):
                        df2.to_excel(writer,sheet_name = sheetname2,index = False)
                    
                    if (len(df3) > 0):
                        df3.to_excel(writer,sheet_name = sheetname3,index = False)
                    
                    if (len(df4) > 0):
                        df4.to_excel(writer,sheet_name = sheetname4,index = False)

                    writer.close()

                    
                    # Styling the worksheets.
                    styling(workbook,sheetname)
                    styling(workbook,sheetname2)
                    styling(workbook,sheetname3)
                    styling(workbook,sheetname4)
                    
                    # Message showing that all the Interdomain Sheets have been written.
                    messagebox.showinfo("   Interdomain Data Preparation Status","Interdomain KPIs Mail Data Preparation Task Completed!")
                    
                    # Calling the Method(Function) that can write into the Automation tracker sheet.
                    p_one_p_three_appender(sender,workbook)

                    return 'Successful'
                
                else:

                    # Message for the case when there's no Interdomain data present(pre-defined interdomains i.e. CS-Core, PS-Core, RAN, VAS, aren't present in the 
                    # 'Domain Kpi' column of the sheet)
                    messagebox.showerror("  No Interdomain Data Present!","Kindly Check the 'Domain kpi' column of the Planning Status Sheet for the required Domains, i.e, CS-Core(CS,Core), PS-Core(Paco), RAN, VAS! and then retry!")


    # Exception for condition when Today's maintenance date is not present.
    except TomorrowDataNotFound as error:
        messagebox.showerror("  Data for today's maintenance not found",error)
        return "Unsuccessful"
    
    # Handling Custom Exception
    except CustomException as error:
        return "Unsuccessful"
    
    # Handling Key Error 
    except KeyError as e:
        messagebox.showerror("  Check for the below Header ",e)
        return "Unsuccessful"
    
    # Handling Attribute Error 
    except AttributeError as e:
        messagebox.showerror("  Exception Occured",e)
        return "Unsuccessful"
    
    # Handling Exception for permission error for opening/editing Workbook.
    except PermissionError as e:
        e = str(e).split(":")
        e.remove(e[0])
        e = ':'.join(e)
        messagebox.showerror("  Permission Error!",f"Kindly close {e} as it's open in Excel!")
        return "Unsuccessful"

    # Handling any other Exception that has not been handled.
    except Exception as e:
        messagebox.showerror("  Exception Occured",e)
        return "Unsuccessful"

#paco_cscore("Karan Loomba",r"C:/Users/emaienj/Downloads/MPBN Daily Planning Sheet.xlsx")