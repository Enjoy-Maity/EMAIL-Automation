import pandas as pd
import numpy as np
from tkinter import messagebox
from Custom_Exception import CustomException
# from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl import load_workbook,Workbook
from openpyxl.utils import get_column_letter
import os

flag = ''

def styler(workbook_path, sheetname):
    workbook = load_workbook(workbook_path)
    if(sheetname == 'Summary'):
        pass
    if(sheetname == 'Team Availability'):
        pass
    workbook.save(workbook_path)
    workbook.close()
    del workbook

def not_available_writer(**kwargs):
    team_availability_sheet = kwargs['team_availability_sheet']
    remaining_change_responsible_that_are_absent = kwargs['remaining_change_resposible']

    i =0
    while(i < remaining_change_responsible_that_are_absent.size):
        team_availability_sheet.loc[len(team_availability_sheet),remaining_change_responsible_that_are_absent[i]] = 'Not Available'
        i += 1
    
    return team_availability_sheet

def available_writer(array_of_resources_involved, team_availability_sheet):
    i = 0
    while(i < array_of_resources_involved.size):
        team_availability_sheet.loc[len(team_availability_sheet),array_of_resources_involved[i]] = 'Available'
        i += 1

    return team_availability_sheet

def empty_space_writer(remaining_change_responsible, team_availability_sheet):
    i = 0
    while(i < remaining_change_responsible.size):
        team_availability_sheet.loc[len(team_availability_sheet),remaining_change_responsible[i]] = ''
        i+=1
    return team_availability_sheet

def data_filler(**kwargs):
    month = kwargs['current_month']
    day = kwargs['day']
    day_type = kwargs['day_type']
    array_of_resources_involved = kwargs['array_of_resources_involved']
    array_of_resources_involved = np.unique(array_of_resources_involved)
    acceptable_change_responsible = kwargs['acceptable_change_responsible']
    
    # Changing the acceptable_change_responsible list into a numpy array
    acceptable_change_responsible = np.array(acceptable_change_responsible)

    # Getting the names that are not present in the email-package and other resources not assigned
    remaining_change_responsible = np.setdiff1d(acceptable_change_responsible,array_of_resources_involved)

    # Getting the dataframes
    team_availability_sheet = kwargs['team_availability_sheet']
    summary_sheet = kwargs['summary_sheet']
    
    # filling the team_availability_sheet
    team_availability_sheet[len(team_availability_sheet),'S No'] = len(team_availability_sheet)+1
    team_availability_sheet.loc[len(team_availability_sheet),'Month'] = month
    team_availability_sheet.loc[len(team_availability_sheet),'Date'] = datetime.now().strftime("%d-%b-%y")
    team_availability_sheet.loc[len(team_availability_sheet),'Day'] = day
    team_availability_sheet.loc[len(team_availability_sheet),'Day Type'] = day_type

    if(day_type == 'Normal'):
        if(remaining_change_responsible.size > 0):
            team_availability_sheet = not_available_writer(remaining_change_responsible = remaining_change_responsible,
                                                           team_availability_sheet = team_availability_sheet)
        if(array_of_resources_involved.size > 0):
            team_availability_sheet = available_writer(array_of_resources_involved = array_of_resources_involved,
                                                      team_availability_sheet = team_availability_sheet)
    if(day_type == 'Weekend'):
        if(remaining_change_responsible.size > 0):
            team_availability_sheet = empty_space_writer(remaining_change_responsible=remaining_change_responsible,
                                                         team_availability_sheet=team_availability_sheet)
        
        if(array_of_resources_involved.size > 0):
            team_availability_sheet = available_writer(array_of_resources_involved= array_of_resources_involved,
                                                       team_availability_sheet=team_availability_sheet)
        
    if(day_type == 'Holiday Support'):
        if(remaining_change_responsible.size > 0):
            team_availability_sheet = empty_space_writer(remaining_change_responsible=remaining_change_responsible,
                                                         team_availability_sheet=team_availability_sheet)
        
        if(array_of_resources_involved.size > 0):
            team_availability_sheet = available_writer(array_of_resources_involved= array_of_resources_involved,
                                                       team_availability_sheet=team_availability_sheet)
        
    path = kwargs['path']
    writer = pd.ExcelWriter(path=path,
                            engine='openpyxl',
                            mode = 'a',
                            if_sheet_exists='replace')
    pd.to_excel(writer,
                sheet_name = 'Team Availability',
                index = False)
    writer.close()
    del writer

    # Styling the worksheet
    styler(workbook_path=path,
           sheetname='Team Availability')
    
    messagebox.showinfo("    Attendance Workbook Updated!","'Team Availability' worksheet updated!!")

    availability_sheet = team_availability_sheet[team_availability_sheet['Month'] == month]
    normal_working_days_dataframe = availability_sheet[availability_sheet['Day Type'] == 'Normal']
    weekend_working_days_dataframe = availability_sheet[availability_sheet['Day Type'] == 'Weekend']
    holiday_working_days_dataframe = availability_sheet[availability_sheet['Day Type'] == 'Holiday Support']

    # Getting the length of the respective dataframes
    total_normal_working_days = len(normal_working_days_dataframe)
    total_weekend_working_days = len(weekend_working_days_dataframe)
    total_holiday_working_days = len(holiday_working_days_dataframe)

    # Checking if the month name is present in the summary sheet
    month_name_present_in_summary_sheet = False
    
    if(month in summary_sheet['Month'].unique()):
        month_name_present_in_summary_sheet = True
    
    # If month_name_present_in_summary_sheet is False, then creating the rows for the month
    if(not month_name_present_in_summary_sheet):
        summary_sheet.loc[len(summary_sheet),'S No']                    = len(summary_sheet) + 1
        summary_sheet.loc[len(summary_sheet),'Month']                   = month
        summary_sheet.loc[len(summary_sheet),'Total Working Day Count'] = total_normal_working_days
        summary_sheet.loc[len(summary_sheet),'Day Type']                = 'Normal'
        
        summary_sheet.loc[len(summary_sheet),'S No']                    = len(summary_sheet) + 1
        summary_sheet.loc[len(summary_sheet),'Month']                   = month
        summary_sheet.loc[len(summary_sheet),'Total Working Day Count'] = total_weekend_working_days
        summary_sheet.loc[len(summary_sheet),'Day Type']                = 'Weekend'
        
        summary_sheet.loc[len(summary_sheet),'S No']                    = len(summary_sheet) + 1
        summary_sheet.loc[len(summary_sheet),'Month']                   = month
        summary_sheet.loc[len(summary_sheet),'Total Working Day Count'] = total_holiday_working_days
        summary_sheet.loc[len(summary_sheet),'Day Type']                = 'Holiday Support'
        
        summary_sheet.loc[len(summary_sheet),'S No']                    = len(summary_sheet) + 1
        summary_sheet.loc[len(summary_sheet),'Month']                   = month
        summary_sheet.loc[len(summary_sheet),'Total Working Day Count'] = total_normal_working_days
        summary_sheet.loc[len(summary_sheet),'Day Type']                = 'Availability'

    # fixing the indices for different month rows
    length_of_summary_sheet     = len(summary_sheet)
    normal_day_row_index            = length_of_summary_sheet - 4
    weekend_day_row_index           = length_of_summary_sheet - 3
    holiday_day_row_index           = length_of_summary_sheet - 2
    availability_row_index          = length_of_summary_sheet - 1

    # Starting the loop for adding the data in summary sheet
    i = 0
    while(i < acceptable_change_responsible.size):
        change_responsible_selected = acceptable_change_responsible[i]
        summary_sheet.loc[normal_day_row_index,change_responsible_selected]     = len(normal_working_days_dataframe[normal_working_days_dataframe[change_responsible_selected] == 'Available'])
        summary_sheet.loc[weekend_day_row_index,change_responsible_selected]    = len(weekend_working_days_dataframe[weekend_working_days_dataframe[change_responsible_selected] == 'Available'])
        summary_sheet.loc[holiday_day_row_index,change_responsible_selected]    = len(holiday_working_days_dataframe[holiday_working_days_dataframe[change_responsible_selected] == 'Available'])

        percentage = (summary_sheet.iloc[normal_day_row_index]/total_normal_working_days)*100
        summary_sheet.loc[availability_row_index,change_responsible_selected]   = f"{percentage}%"
        i+=1
    
    writer = pd.ExcelWriter(path=path,
                            engine='openpyxl',
                            mode = 'a',
                            if_sheet_exists='replace')
    pd.to_excel(writer,
                sheet_name = 'Summary',
                index=False)
    writer.close()
    del writer

    # Styling the worksheet
    styler(workbook_path=path,
           sheetname='Summary')
    messagebox.showinfo("    Attendance Workbook Updated!","'Summary' worksheet updated!!")


def attendance_workbook_creater(**kwargs):
    path = kwargs['path']
    workbook_to_be_created = Workbook()
    workbook_to_be_created.create_sheet(title = 'Summary',index = 0)
    workbook_to_be_created.create_sheet(title= 'Team Availability',index = 1)
    
    accepted_sheet_list = ['Team Availability','Summary']
    sheets = workbook_to_be_created.sheetnames
    
    i = 0
    while(i < len(sheets)):
        selected_sheet = sheets[i]
        if(not selected_sheet in accepted_sheet_list):
            del workbook_to_be_created[selected_sheet]
        i+=1

    columns_for_team_availability = ['S No','Month','Date','Day','Day Type']
    columns_for_team_availability.extend(kwargs['acceptable_change_responsible'])


    columns_for_summary = ['S No','Month','Total Working Day Count','Day Type']
    columns_for_summary.extend(kwargs['acceptable_change_responsible'])

    summary_sheet = workbook_to_be_created['Summary']
    team_availability_sheet = workbook_to_be_created['Team Availability']

    i = 1
    while(i <= len(columns_for_summary)):
        summary_sheet[f'{get_column_letter(i)}1'].value = columns_for_summary[i-1]
        i+=1
    
    i = 1
    while(i <= len(columns_for_team_availability)):
        team_availability_sheet[f'{get_column_letter(i)}1'] = columns_for_team_availability[i-1]
        i+=1

    workbook_to_be_created.save(path)
    workbook_to_be_created.close()
    del workbook_to_be_created
    
    styler(path,'Summary')
    styler(path,'Team Availability')


def main_function(workbook,**kwargs):
    try:
        if(len(workbook) == 0):
            raise CustomException("    Workbook Not Selected!","Email-Package Workbook not selected ")
        
        # creating a ExcelFile object to read the excel file
        email_package_reader = pd.ExcelFile(workbook,engine= 'openpyxl')
        email_package_reader = pd.read_excel(email_package_reader,'Email-Package')
        
        # Getting the important arrays
        night_shift_lead = kwargs['night_shift_lead']
        buffer_auditor_trainer = kwargs['buffer_auditor_trainer']
        resource_on_automation = kwargs['resource_on_automation']
        acceptable_change_responsible = kwargs['acceptable_change_responsible']
        
        array_of_unique_change_responsible = email_package_reader['Change Responsible'].dropna().unique()
        array_of_unique_technical_validator = email_package_reader['Technical Validator'].dropna().unique()
        
        # deleting the entry for Karan Loomba
        if(not 'Karan Loomba' in array_of_unique_technical_validator):
            acceptable_change_responsible.remove('Karan Loomba')

        array_of_resources_involved = np.append(array_of_unique_change_responsible,array_of_unique_technical_validator)

        # Appending night shift lead, buffer_auditor_trainer and resource on automation in array_of_resources_involved.
        if(night_shift_lead.upper().strip() != 'NA'):
            array_of_resources_involved = np.append(array_of_resources_involved,np.array([night_shift_lead]))
        
        if(buffer_auditor_trainer.upper().strip() != 'NA'):
            array_of_resources_involved = np.append(array_of_resources_involved,np.array([buffer_auditor_trainer]))
        
        if(resource_on_automation.upper().strip() != 'NA'):
            array_of_resources_involved = np.append(array_of_resources_involved,np.array([resource_on_automation]))
        
        #Checking the day type
        response = messagebox.askyesno("    Day-Type Query","Is today a 'Normal' Working Day?")
        day_type = ''
        if(response):
            day_type = 'Normal'
        
        else:
            neo_response = messagebox.askyesno("    Day-Type Query","Is today a 'Weekend' Day?")
            if (neo_response):
                day_type = 'Weekend'
            
            else:
                neo_response = messagebox.askyesno("    Day-Type Query","Is today a 'Holiday'?")
                if(neo_response):
                    day_type = 'Holiday Support'
                else:
                    raise CustomException("    Wrong Selection","You have made the wrong day-type selection!!")

        # Determining the existence of the attendance workbook.
        attendance_workbook = "SRF_MPBN_Team_Availability.xlsx"

        # Folder where the email-package is located
        dirname = os.path.dirname(workbook)

        # condition for checking that attendance workbook is present or not
        if(not os.path.exists(os.path.join(dirname, attendance_workbook))):
            attendance_workbook_creater(path = os.path.join(dirname,attendance_workbook), 
                                        acceptable_change_responsible = acceptable_change_responsible)
        
        attendance_workbook_reader = pd.ExcelFile(attendance_workbook)
        team_availability_sheet = pd.read_excel(attendance_workbook_reader,sheet_name= 'Team Availability')
        summary_sheet = pd.read_excel(attendance_workbook_reader,sheet_name='Summary')

        # getting the current month and day
        current_month = str(datetime.now().__format__("%b'%y"))
        day = str(datetime.now().__format__('%A'))
        
        data_filler(current_month=current_month,
                    day=day,
                    team_availability_sheet = team_availability_sheet,
                    summary_sheet = summary_sheet,
                    day_type = day_type,
                    array_of_resources_involved = array_of_resources_involved,
                    acceptable_change_responsible = acceptable_change_responsible,
                    path = os.path.join(dirname,attendance_workbook))

        attendance_workbook_reader.close()
        email_package_reader.close()
        
    except CustomException:
        flag = 'Unsuccessful'

    finally:
        if(flag == 'Successful'):
            messagebox.showinfo()
        
        return flag