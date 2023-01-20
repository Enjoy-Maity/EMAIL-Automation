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

def p_one_p_three_appender(email_package,sender,workbook):
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
                messagebox.showinfo("   MPBN Planning Automation Tracker Status",f"All planned CRs for Validator {sender} has been updated in MPBN Planning Automation Tracker!")
            
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
##################   Email package sheet creater   ##################
#####################################################################

def email_package__sheet_creater(daily_plan_sheet,workbook,sender):
            # The required columns to write into the Email-Package Worksheet in 
            #S.NO	Execution Date	Maintenance Window	CR NO	Activity Title	Risk	Location	Circle	"No. of Node Involved"
            #"CR Belongs to Same Activity of Previous CR - Yes/NO"	Change Responsible	Activity Checker	Activity Initiator	Impact	Planning Status	Domain	
            # Final Status	Reason For Rollback / Cancel	Design Availability	Technical Validator	Complexity	Activity-Type	Domain kpi	IMPACTED NODE	KPI DETAILS	oss name	oss ip	Total Time spent on Planned CRs (Mins)	Vendor	Protocol	Execution Projection	
            # Interdomin Inter-domain KPI status	Second Level Validation Status	Inter-domain KPI status	MOP View Status
            # Creating the empty lists for the column enteries of the email package sheet.
            execution_date = []
            maintenance_window = []
            cr_no = []
            activity_title = []
            risk = []
            location = []
            circle = []
            no_of_node_involved = []
            cr_belongs_to_same_activity_of_previous_cr_yes_no = []
            change_responsible = []
            activity_checker = []
            activity_initiator = []
            impact = []
            planning_status = []
            domain = []
            final_status = []
            reason_for_rollback_cancel = []
            design_availability = []
            technical_validator = []
            complexity = []
            activity_type = []
            domain_kpi = []
            impacted_node = []
            kpi_details = []
            oss_name = []
            oss_IP = []
            total_time_spent_on_planned_crs_mins = []
            vendor = []
            protocol = []
            execution_projection = []
            interdomain_kpi_status = []
            second_level_validation_status = []
            kpi_status = []
            mop_view_status = []
            
            # Getting the unique values of the planning status column of the excel sheet.
            planned_status_unique_values = list(daily_plan_sheet['Planning Status'].unique())
            
            for i in range(0,len(planned_status_unique_values)):
                # Changing the state of the unique inputs in the planning status column of the excel sheet.
                planned_status_unique_values[i] = planned_status_unique_values[i].strip().upper()
            
            if ((len(planned_status_unique_values) == 1) and (planned_status_unique_values[0] == "NA")):
                # Raising custom exception for condition when there's no input in the planning status column of the Excel Sheet.
                raise CustomException(" Input Missing!","Kindly Enter the Planning Status input in uploaded sheet!")
            
            if ((len(planned_status_unique_values) > 1)):
                if ("NA" in planned_status_unique_values):
                    # Empty list for adding S.NO of rows with wrong input.
                    input_error = []
                    for i in range(0,len(daily_plan_sheet)):
                        if (daily_plan_sheet.iloc[i]['Planning Status'] == "NA"):
                            # Appending the S.NO of row with wrong input.
                            input_error.append(daily_plan_sheet.iloc[i]['S.NO'])
                    
                    # Raising the Exception for rows with no planning status.
                    raise CustomException(" Input Missing!",f"Planning Status Input is Missing for the below S.NO:\n{', '.join(str(num) for num in input_error)}")
                
                if ("PLANNED" in planned_status_unique_values):
                    # Filtering the rows with planned crs
                    daily_plan_sheet['Planning Status'] = daily_plan_sheet['Planning Status'].str.strip()
                    daily_plan_sheet = daily_plan_sheet[daily_plan_sheet['Planning Status'].str.upper() == 'PLANNED']

                else:
                    # Raising Custom Exception for not finding any dataframe row with Planning Status.
                    raise CustomException(" Incorrect Input","Kindly Check the Planning Status input in uploaded sheet!")
            
            # Writing into the Email-Package Sheet
            daily_plan_sheet_unique_cr = daily_plan_sheet['CR NO'].value_counts().index.to_list()
            for idx,cr in enumerate(daily_plan_sheet_unique_cr):
                # Creating the count variable to assign the number of occurences of the CR in the 
                count = daily_plan_sheet['CR NO'].value_counts()[idx]
                
                # Setting the counter to 0 to run a loop until the count value of the CR, to assess all the rows in the dataframe with the same CR Number.
                counter = 0

                # Creating temp variables to hold the temporary data until that temporary data is written onto the result dataframe.
                execution_date_temp = daily_plan_sheet.iloc[0]['Execution Date']
                cr_no_temp = cr
                maintenance_window_temp  =  ''
                activity_title_temp  =  ''
                risk_temp  =  ''
                location_temp  =  ''
                circle_temp  =  ''
                no_of_node_involved_temp  =  ''
                cr_belongs_to_same_activity_of_previous_cr_yes_no_temp  =  ''
                change_responsible_temp  =  ''
                activity_checker_temp  =  ''
                activity_initiator_temp  =  ''
                impact_temp  =  ''
                planning_status_temp  =  ''
                domain_temp  =  ''
                final_status_temp  =  ''
                reason_for_rollback_cancel_temp  =  ''
                design_availability_temp  =  ''
                technical_validator_temp  =  ''
                complexity_temp  =  ''
                activity_type_temp  =  ''
                domain_kpi_temp  =  ''
                impacted_node_temp  =  ''
                kpi_details_temp  =  ''
                oss_name_temp  =  ''
                oss_IP_temp  =  ''
                total_time_spent_on_planned_crs_mins_temp  =  ''
                vendor_temp  =  ''
                protocol_temp  =  ''
                execution_projection_temp  =  ''
                interdomain_kpi_status_temp  =  ''
                second_level_validation_status_temp  =  ''
                kpi_status_temp  =  ''
                mop_view_status_temp  =  ''

                # Starting another loop to collect all data of particular CR no. from the daily_plan_sheet dataframe to assess the data and manipulate it according to our needs.
                for i in range(0,len(daily_plan_sheet)):
                    if (daily_plan_sheet.iloc[i]['CR NO'] == cr):
                        if (counter<count):
                            if (count>1):
                                # Data for the RAN should be written first for the CR Data, if there's any row with RAN domain KPI for the CR number. 
                                if (daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith('RAN')):

                                    if (len(daily_plan_sheet.iloc[i]['IMPACTED NODE'].strip()) == 0) or (str(daily_plan_sheet.iloc[i]['IMPACTED NODE']).__contains__('NA')) or (str(daily_plan_sheet.iloc[i]['IMPACTED NODE']).__contains__('na')):
                                        impacted_node_temp = impacted_node_temp
                                    else:
                                        if (len(impacted_node_temp) == 0):
                                            impacted_node_temp = f"({str(daily_plan_sheet.iloc[i]['Domain kpi'])} ):- {str(daily_plan_sheet.iloc[i]['IMPACTED NODE'])}"
                                        else:
                                            impacted_node_temp = f"({str(daily_plan_sheet.iloc[i]['Domain kpi'])} ):- {str(daily_plan_sheet.iloc[i]['IMPACTED NODE'])} || {impacted_node_temp}"
                                    
                                    if (len(domain_kpi_temp) == 0):
                                        domain_kpi_temp = f"{daily_plan_sheet.iloc[i]['Domain kpi']}"
                                    elif (len(domain_kpi_temp) > 0):
                                        domain_kpi_temp = f"{daily_plan_sheet.iloc[i]['Domain kpi']} || {domain_kpi_temp}"
                                    
                                    if (len(daily_plan_sheet.iloc[i]['KPI DETAILS'].strip()) == 0):
                                        kpi_details_temp = kpi_details_temp
                                    else:
                                        if (len(kpi_details_temp) == 0):
                                            kpi_details_temp = f"({str(daily_plan_sheet.iloc[i]['Domain kpi'])} ):- {(daily_plan_sheet.iloc[i]['KPI DETAILS'])}"
                                        if (len(kpi_details_temp)>0):
                                            kpi_details_temp = f"({str(daily_plan_sheet.iloc[i]['Domain kpi'])} ):- {(daily_plan_sheet.iloc[i]['KPI DETAILS'])} || {kpi_details_temp}"
                                    
                                    if (len(str(daily_plan_sheet.iloc[i]['oss name']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['oss name']).__contains__('NA')) :
                                        oss_name_temp = oss_name_temp
                                    else: 
                                        oss_name_temp  =  daily_plan_sheet.iloc[i]['oss name']
                                    
                                    if (len(str(daily_plan_sheet.iloc[i]['oss ip']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['oss ip']).__contains__('NA')) :
                                        oss_IP_temp = oss_IP_temp
                                    else:
                                        oss_IP_temp  =  daily_plan_sheet.iloc[i]['oss ip']

                                    if (len(maintenance_window_temp)) == 0:
                                        if (len(str(daily_plan_sheet.iloc[i]['Maintenance Window']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Maintenance Window']).__contains__('NA')):
                                            maintenance_window_temp = maintenance_window_temp
                                        else:
                                            maintenance_window_temp  =  daily_plan_sheet.iloc[i]['Maintenance Window']

                                    if(len(activity_title_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity Title']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Activity Title']).__contains__('NA')):
                                            activity_title_temp = activity_title_temp
                                        else:
                                            activity_title_temp  =  daily_plan_sheet.iloc[i]['Activity Title']

                                    if(len(risk_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Risk']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Risk']).__contains__('NA')):
                                            risk_temp = risk_temp
                                        else:
                                            risk_temp  =  daily_plan_sheet.iloc[i]['Risk']

                                    if (len(location_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Location']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Location']).__contains__('NA')):
                                            location_temp = location_temp
                                        else:    
                                            location_temp  =  daily_plan_sheet.iloc[i]['Location']

                                    if (len(circle_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Circle']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Circle']).__contains__('NA')):
                                            circle_temp = circle_temp
                                        else:
                                            circle_temp  =  daily_plan_sheet.iloc[i]['Circle']
                                    
                                    if (len(str(no_of_node_involved_temp)) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['No. of Node Involved']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['No. of Node Involved']).__contains__('NA')):
                                            no_of_node_involved_temp = no_of_node_involved_temp
                                        else:
                                            no_of_node_involved_temp  =  daily_plan_sheet.iloc[i]['No. of Node Involved']
                                    
                                    if (len(cr_belongs_to_same_activity_of_previous_cr_yes_no_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['CR Belongs to Same Activity of Previous CR- Yes/NO']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['CR Belongs to Same Activity of Previous CR- Yes/NO']).__contains__('NA')):
                                            cr_belongs_to_same_activity_of_previous_cr_yes_no_temp = cr_belongs_to_same_activity_of_previous_cr_yes_no_temp
                                        else:
                                            cr_belongs_to_same_activity_of_previous_cr_yes_no_temp  =  daily_plan_sheet.iloc[i]['CR Belongs to Same Activity of Previous CR- Yes/NO']
                        
                                    if (len(change_responsible_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Change Responsible']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Change Responsible']).__contains__('NA')):
                                            change_responsible_temp = change_responsible_temp
                                        else:
                                            change_responsible_temp =  daily_plan_sheet.iloc[i]['Change Responsible']

                                    if (len(activity_checker_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity Checker']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Activity Checker']).__contains__('NA')):
                                            activity_checker_temp = activity_checker_temp
                                        else:
                                            activity_checker_temp  =  daily_plan_sheet.iloc[i]['Activity Checker']

                                    if (len(activity_initiator_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity Initiator']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Activity Initiator']).__contains__('NA')):
                                            activity_initiator_temp = activity_initiator_temp
                                        else:
                                            activity_initiator_temp  =  daily_plan_sheet.iloc[i]['Activity Initiator']

                                    if (len(impact_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Impact']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Impact']).__contains__('NA')):
                                            impact_temp = impact_temp
                                        else:
                                            impact_temp  =  daily_plan_sheet.iloc[i]['Impact']

                                    if (len(planning_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Planning Status']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Planning Status']).__contains__('NA')):
                                            planning_status_temp = planning_status_temp
                                        else:
                                            planning_status_temp  =  daily_plan_sheet.iloc[i]['Planning Status']

                                    if (len(domain_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Domain']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Domain']).__contains__('NA')):
                                            domain_temp = domain_temp
                                        else:
                                            domain_temp  =  daily_plan_sheet.iloc[i]['Domain']

                                    if (len(final_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Final Status']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Final Status']).__contains__('NA')):
                                            final_status_temp = final_status_temp
                                        else:
                                            final_status_temp  =  daily_plan_sheet.iloc[i]['Final Status']

                                    if (len(reason_for_rollback_cancel_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Reason For Rollback / Cancel']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Reason For Rollback / Cancel']).__contains__('NA')):
                                            reason_for_rollback_cancel_temp  =  reason_for_rollback_cancel_temp
                                        else:
                                            reason_for_rollback_cancel_temp  =  daily_plan_sheet.iloc[i]['Reason For Rollback / Cancel']

                                    if (len(design_availability_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Design Availability']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Design Availability']).__contains__('NA')):
                                            design_availability_temp  =  design_availability_temp
                                        else:
                                            design_availability_temp  =  daily_plan_sheet.iloc[i]['Design Availability']

                                    if (len(technical_validator_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Technical Validator']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Technical Validator']).__contains__('NA')):
                                            technical_validator_temp  =  technical_validator_temp
                                        else:
                                            technical_validator_temp  =  daily_plan_sheet.iloc[i]['Technical Validator']

                                    if (len(complexity_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Complexity']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Complexity']).__contains__('NA')):
                                            complexity_temp  =  complexity_temp
                                        else:
                                            complexity_temp  =  daily_plan_sheet.iloc[i]['Complexity']

                                    if (len(activity_type_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity-Type']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Activity-Type']).__contains__('NA')):
                                            activity_type_temp  =  activity_type_temp
                                        else:
                                            activity_type_temp  =  daily_plan_sheet.iloc[i]['Activity-Type']

                                    if (len(total_time_spent_on_planned_crs_mins_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Total Time spent on Planned CRs (Mins)']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Total Time spent on Planned CRs (Mins)']).__contains__('NA')):
                                            total_time_spent_on_planned_crs_mins_temp  =  total_time_spent_on_planned_crs_mins_temp
                                        else:
                                            total_time_spent_on_planned_crs_mins_temp  =  daily_plan_sheet.iloc[i]['Total Time spent on Planned CRs (Mins)']

                                    if (len(vendor_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Vendor']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Vendor']).__contains__('NA')):
                                            vendor_temp  =  vendor_temp
                                        else:
                                            vendor_temp  =  daily_plan_sheet.iloc[i]['Vendor']

                                    if (len(protocol_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Protocol']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Protocol']).__contains__('NA')):
                                            protocol_temp  =  protocol_temp
                                        else:
                                            protocol_temp  =  daily_plan_sheet.iloc[i]['Protocol']

                                    if (len(execution_projection_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Execution Projection']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Execution Projection']).__contains__('NA')):
                                            execution_projection_temp  =  execution_projection_temp
                                        else:
                                            execution_projection_temp  =  daily_plan_sheet.iloc[i]['Execution Projection']

                                    if (len(interdomain_kpi_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Inter-domain Name']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Inter-domain Name']).__contains__('NA')):
                                            interdomain_kpi_status_temp  =  interdomain_kpi_status_temp
                                        else:
                                            interdomain_kpi_status_temp  =  daily_plan_sheet.iloc[i]['Inter-domain Name']

                                    if (len(second_level_validation_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Second Level Validation Status']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Second Level Validation Status']).__contains__('NA')):
                                            second_level_validation_status_temp  =  second_level_validation_status_temp
                                        else:
                                            second_level_validation_status_temp  =  daily_plan_sheet.iloc[i]['Second Level Validation Status']

                                    if (len(kpi_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Inter-domain KPI status']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['Inter-domain KPI status']).__contains__('NA')):
                                            kpi_status_temp  =  kpi_status_temp
                                        else:
                                            kpi_status_temp  =  daily_plan_sheet.iloc[i]['Inter-domain KPI status']

                                    if (len(mop_view_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['MOP View Status']).strip()) == 0) or (str(daily_plan_sheet.iloc[i]['MOP View Status']).__contains__('NA')):
                                            mop_view_status_temp  =  mop_view_status_temp
                                        else:
                                            mop_view_status_temp  =  daily_plan_sheet.iloc[i]['MOP View Status']

                                    
                                else:
                                    if (len(daily_plan_sheet.iloc[i]['IMPACTED NODE'].strip()) == 0) or (str(daily_plan_sheet.iloc[i]['IMPACTED NODE']).__contains__('NA')):
                                        impacted_node_temp = impacted_node_temp
                                    else:
                                        if (len(impacted_node_temp) == 0):
                                            impacted_node_temp = '('+str(daily_plan_sheet.iloc[i]['Domain kpi'])+' ):- '+str(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
                                        else:
                                            impacted_node_temp +=  ' || '+'('+str(daily_plan_sheet.iloc[i]['Domain kpi'])+' ):- '+str(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
                                    
                                    if (len(domain_kpi_temp) == 0):
                                        domain_kpi_temp = daily_plan_sheet.iloc[i]['Domain kpi']
                                    
                                    elif (len(domain_kpi_temp)>0):
                                        domain_kpi_temp +=  ' || '+daily_plan_sheet.iloc[i]['Domain kpi']
                                    
                                    if (len(daily_plan_sheet.iloc[i]['KPI DETAILS'].strip()) == 0):
                                        kpi_details_temp = kpi_details_temp
                                    else:
                                        if (len(kpi_details_temp) == 0):
                                            kpi_details_temp = f"({str(daily_plan_sheet.iloc[i]['Domain kpi'])} ):- {str(daily_plan_sheet.iloc[i]['KPI DETAILS'])}"
                                        elif (len(kpi_details_temp)>0):
                                            kpi_details_temp +=  f" || ({str(daily_plan_sheet.iloc[i]['Domain kpi'])} ):- {str(daily_plan_sheet.iloc[i]['KPI DETAILS'])}"
                                    
                                    if (len(str(daily_plan_sheet.iloc[i]['oss name']).strip()) == 0) :
                                        oss_name_temp = oss_name_temp
                                    else: 
                                        oss_name_temp = oss_name_temp
                                    
                                    if (len(str(daily_plan_sheet.iloc[i]['oss ip']).strip()) == 0) :
                                        oss_IP_temp = oss_IP_temp
                                    else:
                                        oss_IP_temp  =  daily_plan_sheet.iloc[i]['oss ip']

                                    if (len(maintenance_window_temp)) == 0:
                                        if (len(str(daily_plan_sheet.iloc[i]['Maintenance Window']).strip()) == 0):
                                            maintenance_window_temp = maintenance_window_temp
                                        else:
                                            maintenance_window_temp  =  daily_plan_sheet.iloc[i]['Maintenance Window']

                                    if(len(activity_title_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity Title']).strip()) == 0) :
                                            activity_title_temp = activity_title_temp
                                        else:
                                            activity_title_temp  =  daily_plan_sheet.iloc[i]['Activity Title']

                                    if(len(risk_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Risk']).strip()) == 0) :
                                            risk_temp = risk_temp
                                        else:
                                            risk_temp  =  daily_plan_sheet.iloc[i]['Risk']

                                    if (len(location_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Location']).strip()) == 0) :
                                            location_temp = location_temp
                                        else:    
                                            location_temp  =  daily_plan_sheet.iloc[i]['Location']

                                    if (len(circle_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Circle']).strip()) == 0) :
                                            circle_temp = circle_temp
                                        else:
                                            circle_temp  =  daily_plan_sheet.iloc[i]['Circle']
                                    
                                    if (len(str(no_of_node_involved_temp)) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['No. of Node Involved']).strip()) == 0) :
                                            no_of_node_involved_temp = no_of_node_involved_temp
                                        else:
                                            no_of_node_involved_temp  =  daily_plan_sheet.iloc[i]['No. of Node Involved']
                                    
                                    if (len(cr_belongs_to_same_activity_of_previous_cr_yes_no_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['CR Belongs to Same Activity of Previous CR- Yes/NO']).strip()) == 0) :
                                            cr_belongs_to_same_activity_of_previous_cr_yes_no_temp = cr_belongs_to_same_activity_of_previous_cr_yes_no_temp
                                        else:
                                            cr_belongs_to_same_activity_of_previous_cr_yes_no_temp  =  daily_plan_sheet.iloc[i]['CR Belongs to Same Activity of Previous CR- Yes/NO']
                        
                                    if (len(change_responsible_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Change Responsible']).strip()) == 0):
                                            change_responsible_temp = change_responsible_temp
                                        else:
                                            change_responsible_temp =  daily_plan_sheet.iloc[i]['Change Responsible']

                                    if (len(activity_checker_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity Checker']).strip()) == 0) :
                                            activity_checker_temp = activity_checker_temp
                                        else:
                                            activity_checker_temp  =  daily_plan_sheet.iloc[i]['Activity Checker']

                                    if (len(activity_initiator_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity Initiator']).strip()) == 0) :
                                            activity_initiator_temp = activity_initiator_temp
                                        else:
                                            activity_initiator_temp  =  daily_plan_sheet.iloc[i]['Activity Initiator']

                                    if (len(impact_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Impact']).strip()) == 0):
                                            impact_temp = impact_temp
                                        else:
                                            impact_temp  =  daily_plan_sheet.iloc[i]['Impact']

                                    if (len(planning_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Planning Status']).strip()) == 0):
                                            planning_status_temp = planning_status_temp
                                        else:
                                            planning_status_temp  =  daily_plan_sheet.iloc[i]['Planning Status']

                                    if (len(domain_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Domain']).strip()) == 0):
                                            domain_temp = domain_temp
                                        else:
                                            domain_temp  =  daily_plan_sheet.iloc[i]['Domain']

                                    if (len(final_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Final Status']).strip()) == 0):
                                            final_status_temp = final_status_temp
                                        else:
                                            final_status_temp  =  daily_plan_sheet.iloc[i]['Final Status']

                                    if (len(reason_for_rollback_cancel_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Reason For Rollback / Cancel']).strip()) == 0) :
                                            reason_for_rollback_cancel_temp  =  reason_for_rollback_cancel_temp
                                        else:
                                            reason_for_rollback_cancel_temp  =  daily_plan_sheet.iloc[i]['Reason For Rollback / Cancel']

                                    if (len(design_availability_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Design Availability']).strip()) == 0):
                                            design_availability_temp  =  design_availability_temp
                                        else:
                                            design_availability_temp  =  daily_plan_sheet.iloc[i]['Design Availability']

                                    if (len(technical_validator_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Technical Validator']).strip()) == 0):
                                            technical_validator_temp  =  technical_validator_temp
                                        else:
                                            technical_validator_temp  =  daily_plan_sheet.iloc[i]['Technical Validator']

                                    if (len(complexity_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Complexity']).strip()) == 0):
                                            complexity_temp  =  complexity_temp
                                        else:
                                            complexity_temp  =  daily_plan_sheet.iloc[i]['Complexity']

                                    if (len(activity_type_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity-Type']).strip()) == 0):
                                            activity_type_temp  =  activity_type_temp
                                        else:
                                            activity_type_temp  =  daily_plan_sheet.iloc[i]['Activity-Type']

                                    if (len(total_time_spent_on_planned_crs_mins_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Total Time spent on Planned CRs (Mins)']).strip()) == 0):
                                            total_time_spent_on_planned_crs_mins_temp  =  total_time_spent_on_planned_crs_mins_temp
                                        else:
                                            total_time_spent_on_planned_crs_mins_temp  =  daily_plan_sheet.iloc[i]['Total Time spent on Planned CRs (Mins)']

                                    if (len(vendor_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Vendor']).strip()) == 0):
                                            vendor_temp  =  vendor_temp
                                        else:
                                            vendor_temp  =  daily_plan_sheet.iloc[i]['Vendor']

                                    if (len(protocol_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Protocol']).strip()) == 0):
                                            protocol_temp  =  protocol_temp
                                        else:
                                            protocol_temp  =  daily_plan_sheet.iloc[i]['Protocol']

                                    if (len(execution_projection_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Execution Projection']).strip()) == 0):
                                            execution_projection_temp  =  execution_projection_temp
                                        else:
                                            execution_projection_temp  =  daily_plan_sheet.iloc[i]['Execution Projection']

                                    if (len(interdomain_kpi_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Inter-domain Name']).strip()) == 0):
                                            interdomain_kpi_status_temp  =  interdomain_kpi_status_temp
                                        else:
                                            interdomain_kpi_status_temp  =  daily_plan_sheet.iloc[i]['Inter-domain Name']

                                    if (len(second_level_validation_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Second Level Validation Status']).strip()) == 0):
                                            second_level_validation_status_temp  =  second_level_validation_status_temp
                                        else:
                                            second_level_validation_status_temp  =  daily_plan_sheet.iloc[i]['Second Level Validation Status']

                                    if (len(kpi_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Inter-domain KPI status']).strip()) == 0):
                                            kpi_status_temp  =  kpi_status_temp
                                        else:
                                            kpi_status_temp  =  daily_plan_sheet.iloc[i]['Inter-domain KPI status']

                                    if (len(mop_view_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['MOP View Status']).strip()) == 0):
                                            mop_view_status_temp  =  mop_view_status_temp
                                        else:
                                            mop_view_status_temp  =  daily_plan_sheet.iloc[i]['MOP View Status']
                        
                            elif (count == 1):
                                if (daily_plan_sheet.iloc[i]['CR NO'] == cr):
                                    
                                    if (len(daily_plan_sheet.iloc[i]['IMPACTED NODE'].strip()) == 0):
                                                impacted_node_temp = impacted_node_temp
                                    else:
                                        if (len(impacted_node_temp) == 0):
                                            impacted_node_temp = str(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
                                    
                                    if (len(domain_kpi_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Domain kpi']).strip()) == 0):
                                            domain_kpi_temp = domain_kpi_temp
                                        else:
                                            domain_kpi_temp = daily_plan_sheet.iloc[i]['Domain kpi']
                                    
                                    if (len(daily_plan_sheet.iloc[i]['KPI DETAILS'].strip()) == 0):
                                        kpi_details_temp = kpi_details_temp
                                    else:
                                        if (len(kpi_details_temp) == 0):
                                            kpi_details_temp = str(daily_plan_sheet.iloc[i]['KPI DETAILS'])
                                    
                                    if (len(str(daily_plan_sheet.iloc[i]['oss name']).strip()) == 0):
                                        oss_name_temp = oss_name_temp   
                                    else: 
                                        oss_name_temp  =  daily_plan_sheet.iloc[i]['oss name']
                                    
                                    if (len(str(daily_plan_sheet.iloc[i]['oss ip']).strip()) == 0):
                                        oss_IP_temp = oss_IP_temp
                                    else:
                                        oss_IP_temp  =  daily_plan_sheet.iloc[i]['oss ip']

                                    if (len(maintenance_window_temp)) == 0:
                                        if (len(str(daily_plan_sheet.iloc[i]['Maintenance Window']).strip()) == 0):
                                            maintenance_window_temp = maintenance_window_temp
                                        else:
                                            maintenance_window_temp  =  daily_plan_sheet.iloc[i]['Maintenance Window']

                                    if(len(activity_title_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity Title']).strip()) == 0):
                                            activity_title_temp = activity_title_temp
                                        else:
                                            activity_title_temp  =  daily_plan_sheet.iloc[i]['Activity Title']

                                    if(len(risk_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Risk']).strip()) == 0):
                                            risk_temp = risk_temp
                                        else:
                                            risk_temp  =  daily_plan_sheet.iloc[i]['Risk']

                                    if (len(location_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Location']).strip()) == 0):
                                            location_temp = location_temp
                                        else:    
                                            location_temp  =  daily_plan_sheet.iloc[i]['Location']

                                    if (len(circle_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Circle']).strip()) == 0):
                                            circle_temp = circle_temp
                                        else:
                                            circle_temp  =  daily_plan_sheet.iloc[i]['Circle']
                                    
                                    if (len(str(no_of_node_involved_temp)) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['No. of Node Involved']).strip()) == 0):
                                            no_of_node_involved_temp = no_of_node_involved_temp
                                        else:
                                            no_of_node_involved_temp  =  daily_plan_sheet.iloc[i]['No. of Node Involved']
                                    
                                    if (len(cr_belongs_to_same_activity_of_previous_cr_yes_no_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['CR Belongs to Same Activity of Previous CR- Yes/NO']).strip()) == 0):
                                            cr_belongs_to_same_activity_of_previous_cr_yes_no_temp = cr_belongs_to_same_activity_of_previous_cr_yes_no_temp
                                        else:
                                            cr_belongs_to_same_activity_of_previous_cr_yes_no_temp  =  daily_plan_sheet.iloc[i]['CR Belongs to Same Activity of Previous CR- Yes/NO']
                        
                                    if (len(change_responsible_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Change Responsible']).strip()) == 0):
                                            change_responsible_temp = change_responsible_temp
                                        else:
                                            change_responsible_temp =  daily_plan_sheet.iloc[i]['Change Responsible']

                                    if (len(activity_checker_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity Checker']).strip()) == 0):
                                            activity_checker_temp = activity_checker_temp
                                        else:
                                            activity_checker_temp  =  daily_plan_sheet.iloc[i]['Activity Checker']

                                    if (len(activity_initiator_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity Initiator']).strip()) == 0):
                                            activity_initiator_temp = activity_initiator_temp
                                        else:
                                            activity_initiator_temp  =  daily_plan_sheet.iloc[i]['Activity Initiator']

                                    if (len(impact_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Impact']).strip()) == 0):
                                            impact_temp = impact_temp
                                        else:
                                            impact_temp  =  daily_plan_sheet.iloc[i]['Impact']

                                    if (len(planning_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Planning Status']).strip()) == 0):
                                            planning_status_temp = planning_status_temp
                                        else:
                                            planning_status_temp  =  daily_plan_sheet.iloc[i]['Planning Status']

                                    if (len(domain_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Domain']).strip()) == 0):
                                            domain_temp = domain_temp
                                        else:
                                            domain_temp  =  daily_plan_sheet.iloc[i]['Domain']

                                    if (len(final_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Final Status']).strip()) == 0):
                                            final_status_temp = final_status_temp
                                        else:
                                            final_status_temp  =  daily_plan_sheet.iloc[i]['Final Status']

                                    if (len(reason_for_rollback_cancel_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Reason For Rollback / Cancel']).strip()) == 0):
                                            reason_for_rollback_cancel_temp  =  reason_for_rollback_cancel_temp
                                        else:
                                            reason_for_rollback_cancel_temp  =  daily_plan_sheet.iloc[i]['Reason For Rollback / Cancel']

                                    if (len(design_availability_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Design Availability']).strip()) == 0):
                                            design_availability_temp  =  design_availability_temp
                                        else:
                                            design_availability_temp  =  daily_plan_sheet.iloc[i]['Design Availability']

                                    if (len(technical_validator_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Technical Validator']).strip()) == 0):
                                            technical_validator_temp  =  technical_validator_temp
                                        else:
                                            technical_validator_temp  =  daily_plan_sheet.iloc[i]['Technical Validator']

                                    if (len(complexity_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Complexity']).strip()) == 0):
                                            complexity_temp  =  complexity_temp
                                        else:
                                            complexity_temp  =  daily_plan_sheet.iloc[i]['Complexity']

                                    if (len(activity_type_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Activity-Type']).strip()) == 0):
                                            activity_type_temp  =  activity_type_temp
                                        else:
                                            activity_type_temp  =  daily_plan_sheet.iloc[i]['Activity-Type']

                                    if (len(total_time_spent_on_planned_crs_mins_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Total Time spent on Planned CRs (Mins)']).strip()) == 0):
                                            total_time_spent_on_planned_crs_mins_temp  =  total_time_spent_on_planned_crs_mins_temp
                                        else:
                                            total_time_spent_on_planned_crs_mins_temp  =  daily_plan_sheet.iloc[i]['Total Time spent on Planned CRs (Mins)']

                                    if (len(vendor_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Vendor']).strip()) == 0):
                                            vendor_temp  =  vendor_temp
                                        else:
                                            vendor_temp  =  daily_plan_sheet.iloc[i]['Vendor']

                                    if (len(protocol_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Protocol']).strip()) == 0):
                                            protocol_temp  =  protocol_temp
                                        else:
                                            protocol_temp  =  daily_plan_sheet.iloc[i]['Protocol']

                                    if (len(execution_projection_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Execution Projection']).strip()) == 0):
                                            execution_projection_temp  =  execution_projection_temp
                                        else:
                                            execution_projection_temp  =  daily_plan_sheet.iloc[i]['Execution Projection']

                                    if (len(interdomain_kpi_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Inter-domain Name']).strip()) == 0):
                                            interdomain_kpi_status_temp  =  interdomain_kpi_status_temp
                                        else:
                                            interdomain_kpi_status_temp  =  daily_plan_sheet.iloc[i]['Inter-domain Name']

                                    if (len(second_level_validation_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Second Level Validation Status']).strip()) == 0):
                                            second_level_validation_status_temp  =  second_level_validation_status_temp
                                        else:
                                            second_level_validation_status_temp  =  daily_plan_sheet.iloc[i]['Second Level Validation Status']

                                    if (len(kpi_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['Inter-domain KPI status']).strip()) == 0):
                                            kpi_status_temp  =  kpi_status_temp
                                        else:
                                            kpi_status_temp  =  daily_plan_sheet.iloc[i]['Inter-domain KPI status']

                                    if (len(mop_view_status_temp) == 0):
                                        if (len(str(daily_plan_sheet.iloc[i]['MOP View Status']).strip()) == 0):
                                            mop_view_status_temp  =  mop_view_status_temp
                                        else:
                                            mop_view_status_temp  =  daily_plan_sheet.iloc[i]['MOP View Status']
                            
                        # Incrementing the value of counter
                        counter +=  1
                
                # Creating the List for each column data by appending the temp variable data to respective column list of the particular selected CR Number.
                execution_date.append(execution_date_temp)
                maintenance_window.append(maintenance_window_temp)
                cr_no.append(cr_no_temp)
                activity_title.append(activity_title_temp)
                risk.append(risk_temp)
                location.append(location_temp)
                circle.append(circle_temp)
                no_of_node_involved.append(no_of_node_involved_temp)
                cr_belongs_to_same_activity_of_previous_cr_yes_no.append(cr_belongs_to_same_activity_of_previous_cr_yes_no_temp)
                change_responsible.append(change_responsible_temp)
                activity_checker.append(activity_checker_temp)
                activity_initiator.append(activity_initiator_temp)
                impact.append(impact_temp)
                planning_status.append(planning_status_temp)
                domain.append(domain_temp)
                final_status.append(final_status_temp)
                reason_for_rollback_cancel.append(reason_for_rollback_cancel_temp)
                design_availability.append(design_availability_temp)
                technical_validator.append(technical_validator_temp)
                complexity.append(complexity_temp)
                activity_type.append(activity_type_temp)
                domain_kpi.append(domain_kpi_temp)
                impacted_node.append(impacted_node_temp)
                kpi_details.append(kpi_details_temp)
                oss_name.append(oss_name_temp)
                oss_IP.append(oss_IP_temp)
                total_time_spent_on_planned_crs_mins.append(total_time_spent_on_planned_crs_mins_temp)
                vendor.append(vendor_temp)
                protocol.append(protocol_temp)
                execution_projection.append(execution_projection_temp)
                interdomain_kpi_status.append(interdomain_kpi_status_temp)
                second_level_validation_status.append(second_level_validation_status_temp)
                kpi_status.append(kpi_status_temp)
                mop_view_status.append(mop_view_status_temp)
            
            # Creating the Dictionary for the columns to make the dictionary a pandas dataframe to be written into the excel sheet.
            dictionary1 = {
                'Execution Date':execution_date,
                'Maintenance Window':maintenance_window,
                'CR NO':cr_no,
                'Activity Title':activity_title,
                'Risk':risk,
                'Location':location,
                'Circle':circle,
                'No. of Node Involved':no_of_node_involved,
                'CR Belongs to Same Activity of Previous CR- Yes/NO':cr_belongs_to_same_activity_of_previous_cr_yes_no,
                'Change Responsible':change_responsible,
                'Activity Checker':activity_checker,
                'Activity Initiator':activity_initiator,
                'Impact':impact,
                'Planning Status':planning_status,
                'Domain':domain,
                'Final Status':final_status,
                'Reason For Rollback / Cancel':reason_for_rollback_cancel,
                'Design Availability':design_availability,
                'Technical Validator':technical_validator,
                'Complexity':complexity,
                'Activity-Type':activity_type,
                'Domain kpi':domain_kpi,
                'IMPACTED NODE':impacted_node,
                'KPI DETAILS':kpi_details,
                'oss name':oss_name,
                'oss ip':oss_IP,
                'Total Time spent on Planned CRs (Mins)':total_time_spent_on_planned_crs_mins,
                'Vendor':vendor,
                'Protocol':protocol,
                'Execution Projection':execution_projection,
                'Inter-domain Name':interdomain_kpi_status,
                'Second Level Validation Status':second_level_validation_status,
                'Inter-domain KPI status':kpi_status,
                'MOP View Status':mop_view_status
                }
            
            df = pd.DataFrame(dictionary1)
            df['Execution Date'] = df['Execution Date'].dt.strftime('%m/%d/%Y')
            
            writer = pd.ExcelWriter(workbook,engine = 'openpyxl',mode = 'a',if_sheet_exists = 'replace')
            new_sheetname = 'Email-Package'
            df.index +=  1
            df.replace("NA"," ", inplace=True)
            df.to_excel(writer,sheet_name = new_sheetname,index_label = 'S.NO')
            writer.close()
            
            # Calling the styling function to stylise the worksheet.
            styling(workbook,new_sheetname)

            messagebox.showinfo("   Email Package Data Preparation Status",'Email-Package Sheet also prepared!')
            p_one_p_three_appender(df,sender,workbook)

            # Deleting the dataframe, once it's use is finished.
            del df

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
        for i in range(0,len(daily_plan_sheet)):
            if (daily_plan_sheet.iloc[i]['Execution Date'].strftime('%Y-%m-%d') != tomorrow.strftime('%Y-%m-%d')):
                difference.append(str(daily_plan_sheet.iloc[i]['S.NO']))
        

        

        if len(daily_plan_sheet) == 0:
            raise TomorrowDataNotFound("Data for tomorrow's date is not present in the MPBN Daily Planning Sheet, kindly check!")
        
        if (len(difference) > 0):
            raise TomorrowDataNotFound(f"All the CR's present are not of Today's Maintenace Date for S.NO : {', '.join([str(num) for num in difference])}")
        
        else:
            
            daily_plan_sheet = daily_plan_sheet[daily_plan_sheet['Execution Date'] == tomorrow.strftime('%Y-%m-%d')]
            Email_ID = pd.read_excel(workbook,"Mail Id")
            
            # Finding the Circles and Change Responsible available in the Mail ID worksheet of the MPBN Daily Planning workbook.
            circle = Email_ID['Circle'].tolist()
            original_change_responsible = Email_ID['Change Responsible'].tolist()

            # Changing the case of each original change responsible to upper.
            for i in range(0,len(original_change_responsible)):
                original_change_responsible[i] = original_change_responsible[i].strip().upper()

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
                messagebox.showerror("  Input Errors",f"Input Error in Planning Sheet for S.NO.: {','.join([str(num) for num in input_error])}")
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
                writer = pd.ExcelWriter(workbook,engine = "openpyxl",mode = "a",if_sheet_exists = "replace")
                df.to_excel(writer,sheet_name = sheetname,index = False)
                df2.to_excel(writer,sheet_name = sheetname2,index = False)
                df3.to_excel(writer,sheet_name = sheetname3,index = False)
                df4.to_excel(writer,sheet_name = sheetname4,index = False)

                writer.close()

                # Styling the worksheets.
                styling(workbook,sheetname)
                styling(workbook,sheetname2)
                styling(workbook,sheetname3)
                styling(workbook,sheetname4)
                
                # Message showing that all the Interdomain Sheets have been written.
                messagebox.showinfo("   Interdomain Data Preparation Status","Interdomain KPIs Mail Data Preparation Task Completed!")
                
                # Calling the Method(Function) that can write into the Email-package sheet.
                email_package__sheet_creater(daily_plan_sheet,workbook,sender)

                return 'Successful'


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
        messagebox.showerror("  Permission Error1",f"Kindly close {e} as it's open in Excel!")
        return "Unsuccessful"

    # Handling any other Exception that has not been handled.
    except Exception as e:
        messagebox.showerror("  Exception Occured",e)
        return "Unsuccessful"

#paco_cscore("Manoj Kumar",r"C:/Users/emaienj/Downloads/MPBN Daily Planning Sheet.xlsx")