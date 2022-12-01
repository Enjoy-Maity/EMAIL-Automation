#import os
import pandas as pd
from tkinter import messagebox

def evening_task (sender,night_shift_lead,buffer_auditor_trainer,resource_on_automation,workbook):
    wb=pd.ExcelFile(workbook)
    #print(workbook)
    ws=wb.sheet_names
    worksheet=''

    #print(f"\n\n{workbook}\n\n")
    
    for sheet in ws:
        if sheet == 'Email-Package':
            worksheet=sheet
    

    if (len(worksheet) == 0):
        messagebox.showwarning(' Email-Package Worksheet not Present','Kindly Click the Button for Interdomain Kpi Data Prep First!')
        return 'Unsuccessful'
    
    else:
       
        worksheet=pd.read_excel(wb,worksheet)

        if (len(worksheet) == 0):
            messagebox.showwarning(' Email-Package Worksheet Empty','Kindly Click the Button for Interdomain Kpi Data Prep First!')
            return 'Unsuccessful'
        
        total_no_of_crs=len(worksheet)
        total_no_of_resource = 16
        critical = 0  # Level 1
        delhi_critical = 0
        major = 0     # Level 2
        delhi_major = 0

        manual = 0
        create = 0
        enable = 0
        partially_automation = 0

        maintenance_window = f"({worksheet.at[0,'Maintenance Window']})"
        month_dictionary = {
            '01' : 'Jan',
            '02' : 'Feb',
            '03' : 'Mar',
            '04' : 'Apr',
            '05' : 'May',
            '06' : 'Jun',
            '07' : 'Jul',
            '08' : 'Aug',
            '09' : 'Sep',
            '10' : 'Oct',
            '11' : 'Nov',
            '12' : 'Dec'
        }

        exec_date = worksheet.at[0,'Execution Date'].strip().split('-')

        suffixes = { 1: 'st' , 2: 'nd' , 3: 'rd'}
        day = ''    
        if (3 < int(exec_date[0]) < 21) or (23 < int(exec_date[0]) < 31):
            day = f'{int(exec_date[0]):02d}th'
        else:
            day = f'{int(exec_date[0]):02d}{suffixes[int(exec_date[0])%10]}'
        execution_date= f'{day} {month_dictionary[exec_date[1]]} {exec_date[2]}'

        
        resources_occupied_in_night_activities = len(worksheet['Change Responsible'].unique())

        #worksheet['Circle'] = worksheet['Circle'].astype(str).str.replace("\D+","")
        worksheet.fillna("NA", inplace = True)

        for row in range(0,len(worksheet)):
            if worksheet.at[row,'Risk'].strip() == 'Level 1':
                critical+=1
                if worksheet.at[row,'Circle'].strip() == 'DL':
                    delhi_critical+=1
                if (worksheet.at[row,'Execution Projection'].strip().upper() == 'MANUAL') or (worksheet.at[row,'Execution Projection'].strip().upper() == 'MANNUAL'):
                    manual+=1
                if (worksheet.at[row,'Execution Projection'].strip().upper() == 'CREATE') or (worksheet.at[row,'Execution Projection'].strip().upper() == 'CRETA'):
                    create+=1
                if (worksheet.at[row,'Execution Projection'].strip().upper() == 'ENABLE'):
                    enable+=1
                if (worksheet.at[row,'Execution Projection'].upper().__contains__('ENABLE')) and (worksheet.at[row,'Execution Projection'].upper().__contains__('MANUAL')):
                    partially_automation+=1
                if (worksheet.at[row,'Execution Projection'].upper().__contains__('CREATE')) and (worksheet.at[row,'Execution Projection'].upper().__contains__('MANUAL')):
                    partially_automation+=1
            if worksheet.at[row,'Risk'].strip() == 'Level 2':
                major+=1
                if worksheet.at[row,'Circle'].strip() == 'DL':
                    delhi_major+=1
                if (worksheet.at[row,'Execution Projection'].strip().upper() == 'MANUAL') or (worksheet.at[row,'Execution Projection'].strip().upper() == 'MANNUAL'):
                    manual+=1
                if (worksheet.at[row,'Execution Projection'].strip().upper() == 'CREATE') or (worksheet.at[row,'Execution Projection'].strip().upper() == 'CRETA'):
                    create+=1
                if (worksheet.at[row,'Execution Projection'].strip().upper() == 'ENABLE'):
                    enable+=1
                if (worksheet.at[row,'Execution Projection'].upper().__contains__('ENABLE')) and (worksheet.at[row,'Execution Projection'].upper().__contains__('MANUAL')):
                    partially_automation+=1
                if (worksheet.at[row,'Execution Projection'].upper().__contains__('CREATE')) and (worksheet.at[row,'Execution Projection'].upper().__contains__('MANUAL')):
                    partially_automation+=1

        
        resource_on_leave = total_no_of_resource - (2 + resources_occupied_in_night_activities + 1)
        if (resource_on_leave) < 0:
            resource_on_leave = 0
        # 02d ensures that the integer is printed in double digit format
        message = """
Dear Sir,

<<Pre Notification Critical & Major CRs>>
Date : {}
{}
Total : {} ({:02d} Critical {:02d} Major )
| Team : MPBN
• Bharti-I ->
  Critical {:02d} ({:02d} Delhi)
  Major    {:02d} ({:02d} Delhi)
Resource's occupied in night activities : {:02d}
Resource in Day/Planning :: 2
Resource on Comp off/Leave :- {}
Night Shift Lead :: {}
Resource engaged in CLI Preparation :: N/A
Resource on Buffer/Auditor/Training : {}
Resource on Automation : {}
Rollback CR re-attempt count : 0
Partially completed CR re-attempt count : 0
Updated automation CR count
======================
Total CRs                     :{}
CR Planned Manually           :{}
CR Planned via Enable Tool    :{}
CR Planned via CREATE Tool    :{}
CR Planned Partial Automation :{}
======================

Regards,
{}
        """
        message = message.format(execution_date,maintenance_window,total_no_of_crs,critical,major,critical,delhi_critical,major,delhi_major,resources_occupied_in_night_activities,resource_on_leave,night_shift_lead,buffer_auditor_trainer,resource_on_automation,total_no_of_crs,manual,enable,create,partially_automation,sender)
        # print(message)
        
        file_path = workbook.split("/")
        file_path.remove(file_path[-1])
        file_path = '\\'.join(file_path)
        
        #print(f"\n\n{file_path}\n\n")
        file_path = f'{file_path}\\evening message.txt'
        
        #print(f"\n\n{file_path}\n\n")
        #assert os.path.isfile(file_path)
        with open(file_path,'w') as f:
            f.write(message)
        messagebox.showinfo("   Task Completed Successfully",f"Evening Message generated successfully at {file_path}")
        return 'Successful'

#evening_task('Enjoy Maity','','','',"C:/Daily/MPBN Daily Planning Sheet.xlsx")