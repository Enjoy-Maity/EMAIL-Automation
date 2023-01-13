from subprocess import Popen        # Importing for opening applications like notebook through cmd line command.
import pandas as pd                 # Importing for reading Excel Sheet data and Manipulating it.
from tkinter import messagebox      # Impoprting for showing messages.

def evening_task (sender,night_shift_lead,buffer_auditor_trainer,resource_on_automation,workbook):
    wb=pd.ExcelFile(workbook)
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
       # Reading relevant sheet.
        worksheet=pd.read_excel(wb,worksheet)

        # Checking Condition for the data pertaining to today's maintenance date being non-existent.
        if (len(worksheet) == 0):
            messagebox.showwarning(' Email-Package Worksheet Empty','Kindly Click the Button for Interdomain Kpi Data Prep First!')
            return 'Unsuccessful'
        
        total_no_of_crs=len(worksheet)
        total_no_of_resource = 16
        critical = 0        # Risk Level 1
        delhi_critical = 0
        major = 0           # Risk Level 2
        delhi_major = 0

        manual = 0
        create = 0
        enable = 0
        partially_automation = 0

        maintenance_window = f"({worksheet.at[0,'Maintenance Window']})"
        
        # Creating a dictionary to give the subscipt of the month name based on the month number.
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

        # Splitting the date to format it to be sent through the message.
        exec_date = worksheet.at[0,'Execution Date'].strip().split('/')
        
        # Adding suffices to the date.
        suffixes = { 1: 'st' , 2: 'nd' , 3: 'rd'}
        day = ''    
        if (3 < int(exec_date[1]) < 21) or (23 < int(exec_date[1]) < 31):
            day = f'{int(exec_date[1]):02d}th'
        else:
            day = f'{int(exec_date[1]):02d}{suffixes[int(exec_date[1])%10]}'

        execution_date= f'{day} {month_dictionary[exec_date[0]]} {exec_date[2]}'

        
        resources_occupied_in_night_activities = len(worksheet['Change Responsible'].unique())

        # Filling the blank fields in the dataframe with 'NA'.
        worksheet.fillna("NA", inplace = True)

        '''
            Iterating (Looping) over the dataframe for finding the number of critical and major risk level along with the number of CR's done with 
            Automation, Partial-Automation and Manually.
        '''
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

        # Finding the Number of resources that are on leave.
        resource_on_leave = total_no_of_resource - (2 + resources_occupied_in_night_activities + 1)
        if ((resource_on_leave) < 0):
            resource_on_leave = 0       # If the value of Number of resources falls below zero and becomes negative, which isn't possible, setting it to zero.
        
        # 02d ensures that the integer is printed in double digit format
        # Creating the message text that is going to be sent via telegram.
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
Total CRs                       :{}
CR Planned Manually             :{}
CR Planned via Enable Tool      :{}
CR Planned via CREATE Tool      :{}
CR Planned Partial Automation   :{}
======================

Regards,
{}
        """
        # Adding all the other relevant data to the message text by formatting it.
        message = message.format(execution_date,maintenance_window,total_no_of_crs,critical,major,critical,delhi_critical,major,delhi_major,resources_occupied_in_night_activities,resource_on_leave,night_shift_lead,buffer_auditor_trainer,resource_on_automation,total_no_of_crs,manual,enable,create,partially_automation,sender)
        
        # Creating the file path where the text file for the message is being saved.
        file_path = workbook.split("/")
        file_path.remove(file_path[-1])
        file_path = '\\'.join(file_path)
        file_path = f'{file_path}\\evening message.txt'
        
        # Writing the text into the file defined by the file path.
        with open(file_path,'w') as f:
            f.write(message)
        messagebox.showinfo("   Task Completed Successfully",f"Evening Message generated successfully at {file_path}")
        
        # Asking for response, whether thr user wants to check the message being created.
        response = messagebox.askyesno("   Evening Message","Do You want to open the Evening Message text?")
        
        # If the response is positive, then the created text message is opened in notebook via the use of Popen from subprocess module.
        if (response):
            Popen(['notepad.exe',file_path])
        
        else:
            pass

        return 'Successful'

#evening_task('Enjoy Maity','','','',"C:/Daily/MPBN Daily Planning Sheet.xlsx")