from re import T
import sys
from datetime import datetime,timedelta
import subprocess
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Font,Border,Side,PatternFill,Alignment
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import pandas as pd
from tkinter import *
from tkinter import messagebox


class TomorrowDataNotFound(Exception):
    def __init__(self,msg):
        self.msg = msg

class DomainNotFound(Exception):
    def __init__(self,msg):
        self.msg = msg

#####################################################################
#############################    Styling   ##########################
#####################################################################
try: 
    def styling(workbook,sheetname):
        wb  =  load_workbook(workbook)
        ws  =  wb[sheetname]
        font_style  =  Font(color = "FFFFFF",bold = True)
        col_widths = []

        for row in ws.iter_rows(values_only = True):
            for j,value in enumerate(row):
                if len(col_widths)>j:
                    if col_widths[j] < len(str(value)):
                        col_widths[j] = len(str(value))
                else:
                    col_widths.insert(j,len(str(value)))

        for i,column_width in enumerate(col_widths,1):
            if column_width <= 47:
                ws.column_dimensions[get_column_letter(i)].width = column_width+3
            else:
                ws.column_dimensions[get_column_letter(i)].width = 50


        for column in range(1,ws.max_column+1):   # ws.max_column returns the total number of columns present
            col = get_column_letter(column)
            color_fill = PatternFill(start_color = '0033CC',end_color = '0033CC',fill_type = 'solid')
            ws[col+'1'].font = font_style
            ws[col+'1'].fill = color_fill
            ws[col+'1'].alignment = Alignment(horizontal = 'center',vertical = 'center')

        border = Border(top = Side(border_style = 'thick',color = '000000'),bottom = Side(border_style = 'thick',color = '000000'),left = Side(border_style = 'thick',color = '000000'),right = Side(border_style = 'thick',color = '000000'))

        for row in ws:
            for cell in row:
                cell.alignment = Alignment(horizontal = 'center',vertical = 'center',wrap_text=True)
                cell.border = Border(top = Side(border_style = 'medium',color = '000000'),bottom = Side(border_style = 'medium',color = '000000'),left = Side(border_style = 'medium',color = '000000'),right = Side(border_style = 'medium',color = '000000'))

        #rows = ws.max_row
        
        
        wb.save(workbook)




    def quit(event):
        sys.exit(0)
    #####################################################################
    #############################  Paco_cscore  #########################
    #####################################################################

    def paco_cscore(sender,workbook):
        
        
            #user = subprocess.getoutput("echo %username%") # finding the Username of the user where the directory of the file is located 

            #workbook = r"C:\Daily\MPBN Daily Planning Sheet.xlsx" # system path from where the program will take the input
            daily_plan_sheet = pd.read_excel(workbook,'Planning Sheet')
            daily_plan_sheet.fillna("NA",inplace = True)
            tomorrow = datetime.now()+timedelta(1) # getting tomorrow date for data execution
            daily_plan_sheet = daily_plan_sheet[daily_plan_sheet['Execution Date'] == tomorrow.strftime('%Y-%m-%d')]

            if len(daily_plan_sheet) == 0:
                raise TomorrowDataNotFound("Data for tomorrow's date is not present in the MPBN Daily Planning Sheet, kindly check!")
            
            else:
                Email_ID = pd.read_excel(workbook,"Mail Id")

                circle = Email_ID['Circle'].tolist()

                input_error = []
                result_df = pd.DataFrame()

                for i in range(0,len(daily_plan_sheet)):
                    if daily_plan_sheet.at[i,'CR NO'] == "NA":
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Circle'] not in circle):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Activity Title'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Circle'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Risk'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Location'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Change Responsible'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Impact'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Technical Validator'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Activity-Type'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Vendor'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Protocol'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    if (daily_plan_sheet.at[i,'Execution Projection'] == 'NA'):
                        input_error.append(daily_plan_sheet.at[i,'S.NO'])
                        continue
                    else:
                        result_df = pd.concat([result_df,daily_plan_sheet.iloc[i].to_frame().T], ignore_index= True)
                
                result_df.drop_duplicates(keep = 'first', inplace= True)

                del daily_plan_sheet

                daily_plan_sheet = result_df.copy(deep = True)

                del result_df
                
                input_error.sort()

                if (len(input_error) > 0):
                    messagebox.showerror("  Input Errors",f"Input Error in Planning Sheet for S.NO.: {','.join([str(num) for num in input_error])}")
                    return 'Unsuccessful'
                
                else:
                
                    sheetname = "PS Core-Inter Domain"
                    sheetname2 = "CS Core-Inter Domain"
                    sheetname3 = "RAN-Inter Domain"


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
                        if (daily_plan_sheet.iloc[i]['Domain kpi'].upper() == 'PS-CORE') or (daily_plan_sheet.iloc[i]['Domain kpi'].upper() == 'PS') or (daily_plan_sheet.iloc[i]['Domain kpi'].upper() == 'PS_CORE') or (daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith('PACO')) or (daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith("PS")):
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
                        if (daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith("CS")) or (daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith("STP")) or (daily_plan_sheet.iloc[i]['Domain kpi'].upper().startswith("CORE")) :
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
                    dictionary2 = {'CR':mpbn_cr_no,'Maintenance Window':maintenance_window,'CR Category':cr_category,'Impact':impact,'Location':location,'Circle':circle,'MPBN Activity Title':mpbn_activity_title,'CR Owner Domain':cr_owner_domain,'Change Responsible':mpbn_change_responsible_executor,'Technical Validator/Team Lead':validator,'InterDomain':inter_domain,'Impacted Node Details':impacted_node_details,'KPIs to be monitored':Kpis_to_be_monitored}
                    df2 = pd.DataFrame(dictionary2)

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
                        if (daily_plan_sheet.iloc[i]['Domain kpi'].upper().__contains__("RAN")):
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
                            inter_domain.append(daily_plan_sheet.at[i,'Domain kpi'])
                            impacted_node_details.append(daily_plan_sheet.iloc[i]['IMPACTED NODE'])
                            Kpis_to_be_monitored.append(daily_plan_sheet.iloc[i]['KPI DETAILS'])
                            oss_name.append(daily_plan_sheet.iloc[i]['oss name'])
                            oss_IP.append(daily_plan_sheet.iloc[i]['oss ip'])

                    dictionary3 = {'CR':mpbn_cr_no,'Maintenance Window':maintenance_window,'CR Category':cr_category,'Impact':impact,'Location':location,'Circle':circle,'MPBN Activity Title':mpbn_activity_title,'CR Owner Domain':cr_owner_domain,'Change Responsible':mpbn_change_responsible_executor,'Technical Validator/Team Lead':validator,'InterDomain':inter_domain,'Impacted Node Details':impacted_node_details,'KPIs to be monitored':Kpis_to_be_monitored,'OSS Name':oss_name,'OSS IP':oss_IP}
                    df3 = pd.DataFrame(dictionary3)
                    

                    df.reset_index(drop = True,inplace = True)
                    df2.reset_index(drop = True,inplace = True)
                    df3.reset_index(drop = True,inplace = True)



                    
                    # writer = pd.ExcelWriter(workbook,engine = 'xlsxwriter')

                    # daily_plan_sheet.to_excel(writer,sheet_name = 'Planning Sheet',index = False)
                    # df.to_excel(writer,sheet_name = sheetname,index = False)
                    # df2.to_excel(writer,sheet_name = sheetname2,index = False)
                    # df3.to_excel(writer,sheet_name = sheetname3,index = False)
                    # Email_Id.to_excel(writer,sheet_name = 'Mail Id',index = False)


                    writer = pd.ExcelWriter(workbook,engine = "openpyxl",mode = "a",if_sheet_exists = "replace")
                    wb = writer.book
                    # try:
                    #     wb.remove(wb[sheetname])
                    #     wb.remove(wb[sheetname2])
                    #     wb.remove(wb[sheetname3])
                    # # except:
                    # #     pass
                    # finally:
                    df.to_excel(writer,sheet_name = sheetname,index = False,)
                    df2.to_excel(writer,sheet_name = sheetname2,index = False)
                    df3.to_excel(writer,sheet_name = sheetname3,index = False)


                    writer.save()

                    styling(workbook,sheetname)
                    styling(workbook,sheetname2)
                    styling(workbook,sheetname3)

                    messagebox.showinfo("   Successful Completion","Interdomain KPIs Mail Data Preparation Task completed")

                    email_package__sheet_creater(daily_plan_sheet,tomorrow,workbook)

                    return 'Successful'

                
    def email_package__sheet_creater(daily_plan_sheet,tomorrow,workbook):
                #S.NO	Execution Date	Maintenance Window	CR NO	Activity Title	Risk	Location	Circle	"No. of Node Involved"
                #"CR Belongs to Same Activity of Previous CR - Yes/NO"	Change Responsible	Activity Checker	Activity Initiator	Impact	Planning Status	Domain	
                # Final Status	Reason For Rollback / Cancel	Design Availability	Technical Validator	Complexity	Activity-Type	Domain kpi	IMPACTED NODE	KPI DETAILS	oss name	oss ip	Total Time spent on Planned CRs (Mins)	Vendor	Protocol	Execution Projection	
                # Interdomin KPI status	Second Level Validation Status	KPI status	MOP View Status

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

                daily_plan_sheet = daily_plan_sheet[daily_plan_sheet['Execution Date'] == tomorrow.strftime('%Y-%m-%d')]
                daily_plan_sheet_unique_cr = daily_plan_sheet['CR NO'].value_counts().index.to_list()
                for idx,cr in enumerate(daily_plan_sheet_unique_cr):
                    
                    count = daily_plan_sheet['CR NO'].value_counts()[idx]
                    counter = 0
                    
                    execution_date_temp = daily_plan_sheet.at[0,'Execution Date']
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

                    for i in range(0,len(daily_plan_sheet)):
                        if (daily_plan_sheet.at[i,'CR NO'] == cr):
                            if (counter<count):
                                if (count>1):
                                    if (daily_plan_sheet.at[i,'Domain kpi'].upper().__contains__('RAN')):

                                        if (len(daily_plan_sheet.at[i,'IMPACTED NODE'].strip()) == 0) or (str(daily_plan_sheet.at[i,'IMPACTED NODE']).__contains__('NA')) or (str(daily_plan_sheet.at[i,'IMPACTED NODE']).__contains__('na')):
                                            impacted_node_temp = impacted_node_temp
                                        else:
                                            if (len(impacted_node_temp) == 0):
                                                impacted_node_temp = '('+str(daily_plan_sheet.at[i,'Domain kpi'])+' ):- '+str(daily_plan_sheet.at[i,'IMPACTED NODE'])
                                            else:
                                                impacted_node_temp = '('+str(daily_plan_sheet.at[i,'Domain kpi'])+' ):- '+str(daily_plan_sheet.at[i,'IMPACTED NODE'])+' || '+impacted_node_temp
                                        
                                        if (len(domain_kpi_temp) == 0):
                                            domain_kpi_temp = daily_plan_sheet.at[i,'Domain kpi']
                                        elif (len(domain_kpi_temp)>0):
                                            domain_kpi_temp = daily_plan_sheet.at[i,'Domain kpi']+' || '+domain_kpi_temp
                                        
                                        if (len(daily_plan_sheet.at[i,'KPI DETAILS'].strip()) == 0) or (str(daily_plan_sheet.at[i,'KPI DETAILS']).__contains__('NA')) or (str(daily_plan_sheet.at[i,'KPI DETAILS']).__contains__('na')):
                                            kpi_details_temp = kpi_details_temp
                                        else:
                                            if (len(kpi_details_temp) == 0):
                                                kpi_details_temp = '('+str(daily_plan_sheet.at[i,'Domain kpi'])+' ):- '+str(daily_plan_sheet.at[i,'KPI DETAILS'])
                                            elif (len(kpi_details_temp)>0):
                                                kpi_details_temp = '('+str(daily_plan_sheet.at[i,'Domain kpi'])+' ):- '+str(daily_plan_sheet.at[i,'KPI DETAILS'])+' || '+kpi_details_temp
                                        
                                        if (len(str(daily_plan_sheet.at[i,'oss name']).strip()) == 0) or (str(daily_plan_sheet.at[i,'oss name']).__contains__('NA')) :
                                            oss_name_temp = oss_name_temp
                                        else: 
                                            oss_name_temp  =  daily_plan_sheet.at[i,'oss name']
                                        
                                        if (len(str(daily_plan_sheet.at[i,'oss ip']).strip()) == 0) or (str(daily_plan_sheet.at[i,'oss ip']).__contains__('NA')) :
                                            oss_IP_temp = oss_IP_temp
                                        else:
                                            oss_IP_temp  =  daily_plan_sheet.at[i,'oss ip']

                                        if (len(maintenance_window_temp)) == 0:
                                            if (len(str(daily_plan_sheet.at[i,'Maintenance Window']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Maintenance Window']).__contains__('NA')):
                                                maintenance_window_temp = maintenance_window_temp
                                            else:
                                                maintenance_window_temp  =  daily_plan_sheet.at[i,'Maintenance Window']

                                        if(len(activity_title_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity Title']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity Title']).__contains__('NA')):
                                                activity_title_temp = activity_title_temp
                                            else:
                                                activity_title_temp  =  daily_plan_sheet.at[i,'Activity Title']

                                        if(len(risk_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Risk']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Risk']).__contains__('NA')):
                                                risk_temp = risk_temp
                                            else:
                                                risk_temp  =  daily_plan_sheet.at[i,'Risk']

                                        if (len(location_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Location']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Location']).__contains__('NA')):
                                                location_temp = location_temp
                                            else:    
                                                location_temp  =  daily_plan_sheet.at[i,'Location']

                                        if (len(circle_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Circle']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Circle']).__contains__('NA')):
                                                circle_temp = circle_temp
                                            else:
                                                circle_temp  =  daily_plan_sheet.at[i,'Circle']
                                        
                                        if (len(str(no_of_node_involved_temp)) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'No. of Node Involved']).strip()) == 0) or (str(daily_plan_sheet.at[i,'No. of Node Involved']).__contains__('NA')):
                                                no_of_node_involved_temp = no_of_node_involved_temp
                                            else:
                                                no_of_node_involved_temp  =  daily_plan_sheet.at[i,'No. of Node Involved']
                                        
                                        if (len(cr_belongs_to_same_activity_of_previous_cr_yes_no_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'CR Belongs to Same Activity of Previous CR- Yes/NO']).strip()) == 0) or (str(daily_plan_sheet.at[i,'CR Belongs to Same Activity of Previous CR- Yes/NO']).__contains__('NA')):
                                                cr_belongs_to_same_activity_of_previous_cr_yes_no_temp = cr_belongs_to_same_activity_of_previous_cr_yes_no_temp
                                            else:
                                                cr_belongs_to_same_activity_of_previous_cr_yes_no_temp  =  daily_plan_sheet.at[i,'CR Belongs to Same Activity of Previous CR- Yes/NO']
                            
                                        if (len(change_responsible_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Change Responsible']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Change Responsible']).__contains__('NA')):
                                                change_responsible_temp = change_responsible_temp
                                            else:
                                                change_responsible_temp =  daily_plan_sheet.at[i,'Change Responsible']

                                        if (len(activity_checker_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity Checker']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity Checker']).__contains__('NA')):
                                                activity_checker_temp = activity_checker_temp
                                            else:
                                                activity_checker_temp  =  daily_plan_sheet.at[i,'Activity Checker']

                                        if (len(activity_initiator_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity Initiator']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity Initiator']).__contains__('NA')):
                                                activity_initiator_temp = activity_initiator_temp
                                            else:
                                                activity_initiator_temp  =  daily_plan_sheet.at[i,'Activity Initiator']

                                        if (len(impact_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Impact']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Impact']).__contains__('NA')):
                                                impact_temp = impact_temp
                                            else:
                                                impact_temp  =  daily_plan_sheet.at[i,'Impact']

                                        if (len(planning_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Planning Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Planning Status']).__contains__('NA')):
                                                planning_status_temp = planning_status_temp
                                            else:
                                                planning_status_temp  =  daily_plan_sheet.at[i,'Planning Status']

                                        if (len(domain_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Domain']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Domain']).__contains__('NA')):
                                                domain_temp = domain_temp
                                            else:
                                                domain_temp  =  daily_plan_sheet.at[i,'Domain']

                                        if (len(final_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Final Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Final Status']).__contains__('NA')):
                                                final_status_temp = final_status_temp
                                            else:
                                                final_status_temp  =  daily_plan_sheet.at[i,'Final Status']

                                        if (len(reason_for_rollback_cancel_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Reason For Rollback / Cancel']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Reason For Rollback / Cancel']).__contains__('NA')):
                                                reason_for_rollback_cancel_temp  =  reason_for_rollback_cancel_temp
                                            else:
                                                reason_for_rollback_cancel_temp  =  daily_plan_sheet.at[i,'Reason For Rollback / Cancel']

                                        if (len(design_availability_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Design Availability']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Design Availability']).__contains__('NA')):
                                                design_availability_temp  =  design_availability_temp
                                            else:
                                                design_availability_temp  =  daily_plan_sheet.at[i,'Design Availability']

                                        if (len(technical_validator_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Technical Validator']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Technical Validator']).__contains__('NA')):
                                                technical_validator_temp  =  technical_validator_temp
                                            else:
                                                technical_validator_temp  =  daily_plan_sheet.at[i,'Technical Validator']

                                        if (len(complexity_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Complexity']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Complexity']).__contains__('NA')):
                                                complexity_temp  =  complexity_temp
                                            else:
                                                complexity_temp  =  daily_plan_sheet.at[i,'Complexity']

                                        if (len(activity_type_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity-Type']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity-Type']).__contains__('NA')):
                                                activity_type_temp  =  activity_type_temp
                                            else:
                                                activity_type_temp  =  daily_plan_sheet.at[i,'Activity-Type']

                                        if (len(total_time_spent_on_planned_crs_mins_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Total Time spent on Planned CRs (Mins)']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Total Time spent on Planned CRs (Mins)']).__contains__('NA')):
                                                total_time_spent_on_planned_crs_mins_temp  =  total_time_spent_on_planned_crs_mins_temp
                                            else:
                                                total_time_spent_on_planned_crs_mins_temp  =  daily_plan_sheet.at[i,'Total Time spent on Planned CRs (Mins)']

                                        if (len(vendor_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Vendor']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Vendor']).__contains__('NA')):
                                                vendor_temp  =  vendor_temp
                                            else:
                                                vendor_temp  =  daily_plan_sheet.at[i,'Vendor']

                                        if (len(protocol_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Protocol']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Protocol']).__contains__('NA')):
                                                protocol_temp  =  protocol_temp
                                            else:
                                                protocol_temp  =  daily_plan_sheet.at[i,'Protocol']

                                        if (len(execution_projection_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Execution Projection']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Execution Projection']).__contains__('NA')):
                                                execution_projection_temp  =  execution_projection_temp
                                            else:
                                                execution_projection_temp  =  daily_plan_sheet.at[i,'Execution Projection']

                                        if (len(interdomain_kpi_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Interdomain KPI status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Interdomain KPI status']).__contains__('NA')):
                                                interdomain_kpi_status_temp  =  interdomain_kpi_status_temp
                                            else:
                                                interdomain_kpi_status_temp  =  daily_plan_sheet.at[i,'Interdomain KPI status']

                                        if (len(second_level_validation_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Second Level Validation Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Second Level Validation Status']).__contains__('NA')):
                                                second_level_validation_status_temp  =  second_level_validation_status_temp
                                            else:
                                                second_level_validation_status_temp  =  daily_plan_sheet.at[i,'Second Level Validation Status']

                                        if (len(kpi_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'KPI status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'KPI status']).__contains__('NA')):
                                                kpi_status_temp  =  kpi_status_temp
                                            else:
                                                kpi_status_temp  =  daily_plan_sheet.at[i,'KPI status']

                                        if (len(mop_view_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'MOP View Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'MOP View Status']).__contains__('NA')):
                                                mop_view_status_temp  =  mop_view_status_temp
                                            else:
                                                mop_view_status_temp  =  daily_plan_sheet.at[i,'MOP View Status']

                                        
                                    else:
                                        if (len(daily_plan_sheet.at[i,'IMPACTED NODE'].strip()) == 0) or (str(daily_plan_sheet.at[i,'IMPACTED NODE']).__contains__('NA')):
                                            impacted_node_temp = impacted_node_temp
                                        else:
                                            if (len(impacted_node_temp) == 0):
                                                impacted_node_temp = '('+str(daily_plan_sheet.at[i,'Domain kpi'])+' ):- '+str(daily_plan_sheet.at[i,'IMPACTED NODE'])
                                            else:
                                                impacted_node_temp +=  ' || '+'('+str(daily_plan_sheet.at[i,'Domain kpi'])+' ):- '+str(daily_plan_sheet.at[i,'IMPACTED NODE'])
                                        
                                        if (len(domain_kpi_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Domain kpi']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Domain kpi']).__contains__('NA')):
                                                raise DomainNotFound("Domain KPI can't be Empty ")
                                            else:
                                                domain_kpi_temp = daily_plan_sheet.at[i,'Domain kpi']
                                        elif (len(domain_kpi_temp)>0):
                                            domain_kpi_temp +=  ' || '+daily_plan_sheet.at[i,'Domain kpi']
                                        
                                        if (len(daily_plan_sheet.at[i,'KPI DETAILS'].strip()) == 0) or (str(daily_plan_sheet.at[i,'KPI DETAILS']).__contains__('NA')):
                                            kpi_details_temp = kpi_details_temp
                                        else:
                                            if (len(kpi_details_temp) == 0):
                                                kpi_details_temp = '('+str(daily_plan_sheet.at[i,'Domain kpi'])+' ):- '+str(daily_plan_sheet.at[i,'KPI DETAILS'])
                                            elif (len(kpi_details_temp)>0):
                                                kpi_details_temp +=  ' || '+'('+str(daily_plan_sheet.at[i,'Domain kpi'])+' ):- '+str(daily_plan_sheet.at[i,'KPI DETAILS'])
                                        
                                        if (len(str(daily_plan_sheet.at[i,'oss name']).strip()) == 0) or (str(daily_plan_sheet.at[i,'oss name']).__contains__('NA')) :
                                            oss_name_temp = oss_name_temp
                                        else: 
                                            oss_name_temp = oss_name_temp
                                        
                                        if (len(str(daily_plan_sheet.at[i,'oss ip']).strip()) == 0) or (str(daily_plan_sheet.at[i,'oss ip']).__contains__('NA')) :
                                            oss_IP_temp = oss_IP_temp
                                        else:
                                            oss_IP_temp  =  daily_plan_sheet.at[i,'oss ip']

                                        if (len(maintenance_window_temp)) == 0:
                                            if (len(str(daily_plan_sheet.at[i,'Maintenance Window']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Maintenance Window']).__contains__('NA')):
                                                maintenance_window_temp = maintenance_window_temp
                                            else:
                                                maintenance_window_temp  =  daily_plan_sheet.at[i,'Maintenance Window']

                                        if(len(activity_title_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity Title']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity Title']).__contains__('NA')):
                                                activity_title_temp = activity_title_temp
                                            else:
                                                activity_title_temp  =  daily_plan_sheet.at[i,'Activity Title']

                                        if(len(risk_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Risk']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Risk']).__contains__('NA')):
                                                risk_temp = risk_temp
                                            else:
                                                risk_temp  =  daily_plan_sheet.at[i,'Risk']

                                        if (len(location_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Location']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Location']).__contains__('NA')):
                                                location_temp = location_temp
                                            else:    
                                                location_temp  =  daily_plan_sheet.at[i,'Location']

                                        if (len(circle_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Circle']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Circle']).__contains__('NA')):
                                                circle_temp = circle_temp
                                            else:
                                                circle_temp  =  daily_plan_sheet.at[i,'Circle']
                                        
                                        if (len(str(no_of_node_involved_temp)) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'No. of Node Involved']).strip()) == 0) or (str(daily_plan_sheet.at[i,'No. of Node Involved']).__contains__('NA')):
                                                no_of_node_involved_temp = no_of_node_involved_temp
                                            else:
                                                no_of_node_involved_temp  =  daily_plan_sheet.at[i,'No. of Node Involved']
                                        
                                        if (len(cr_belongs_to_same_activity_of_previous_cr_yes_no_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'CR Belongs to Same Activity of Previous CR- Yes/NO']).strip()) == 0) or (str(daily_plan_sheet.at[i,'CR Belongs to Same Activity of Previous CR- Yes/NO']).__contains__('NA')):
                                                cr_belongs_to_same_activity_of_previous_cr_yes_no_temp = cr_belongs_to_same_activity_of_previous_cr_yes_no_temp
                                            else:
                                                cr_belongs_to_same_activity_of_previous_cr_yes_no_temp  =  daily_plan_sheet.at[i,'CR Belongs to Same Activity of Previous CR- Yes/NO']
                            
                                        if (len(change_responsible_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Change Responsible']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Change Responsible']).__contains__('NA')):
                                                change_responsible_temp = change_responsible_temp
                                            else:
                                                change_responsible_temp =  daily_plan_sheet.at[i,'Change Responsible']

                                        if (len(activity_checker_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity Checker']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity Checker']).__contains__('NA')):
                                                activity_checker_temp = activity_checker_temp
                                            else:
                                                activity_checker_temp  =  daily_plan_sheet.at[i,'Activity Checker']

                                        if (len(activity_initiator_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity Initiator']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity Initiator']).__contains__('NA')):
                                                activity_initiator_temp = activity_initiator_temp
                                            else:
                                                activity_initiator_temp  =  daily_plan_sheet.at[i,'Activity Initiator']

                                        if (len(impact_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Impact']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Impact']).__contains__('NA')):
                                                impact_temp = impact_temp
                                            else:
                                                impact_temp  =  daily_plan_sheet.at[i,'Impact']

                                        if (len(planning_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Planning Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Planning Status']).__contains__('NA')):
                                                planning_status_temp = planning_status_temp
                                            else:
                                                planning_status_temp  =  daily_plan_sheet.at[i,'Planning Status']

                                        if (len(domain_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Domain']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Domain']).__contains__('NA')):
                                                domain_temp = domain_temp
                                            else:
                                                domain_temp  =  daily_plan_sheet.at[i,'Domain']

                                        if (len(final_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Final Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Final Status']).__contains__('NA')):
                                                final_status_temp = final_status_temp
                                            else:
                                                final_status_temp  =  daily_plan_sheet.at[i,'Final Status']

                                        if (len(reason_for_rollback_cancel_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Reason For Rollback / Cancel']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Reason For Rollback / Cancel']).__contains__('NA')):
                                                reason_for_rollback_cancel_temp  =  reason_for_rollback_cancel_temp
                                            else:
                                                reason_for_rollback_cancel_temp  =  daily_plan_sheet.at[i,'Reason For Rollback / Cancel']

                                        if (len(design_availability_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Design Availability']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Design Availability']).__contains__('NA')):
                                                design_availability_temp  =  design_availability_temp
                                            else:
                                                design_availability_temp  =  daily_plan_sheet.at[i,'Design Availability']

                                        if (len(technical_validator_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Technical Validator']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Technical Validator']).__contains__('NA')):
                                                technical_validator_temp  =  technical_validator_temp
                                            else:
                                                technical_validator_temp  =  daily_plan_sheet.at[i,'Technical Validator']

                                        if (len(complexity_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Complexity']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Complexity']).__contains__('NA')):
                                                complexity_temp  =  complexity_temp
                                            else:
                                                complexity_temp  =  daily_plan_sheet.at[i,'Complexity']

                                        if (len(activity_type_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity-Type']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity-Type']).__contains__('NA')):
                                                activity_type_temp  =  activity_type_temp
                                            else:
                                                activity_type_temp  =  daily_plan_sheet.at[i,'Activity-Type']

                                        if (len(total_time_spent_on_planned_crs_mins_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Total Time spent on Planned CRs (Mins)']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Total Time spent on Planned CRs (Mins)']).__contains__('NA')):
                                                total_time_spent_on_planned_crs_mins_temp  =  total_time_spent_on_planned_crs_mins_temp
                                            else:
                                                total_time_spent_on_planned_crs_mins_temp  =  daily_plan_sheet.at[i,'Total Time spent on Planned CRs (Mins)']

                                        if (len(vendor_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Vendor']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Vendor']).__contains__('NA')):
                                                vendor_temp  =  vendor_temp
                                            else:
                                                vendor_temp  =  daily_plan_sheet.at[i,'Vendor']

                                        if (len(protocol_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Protocol']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Protocol']).__contains__('NA')):
                                                protocol_temp  =  protocol_temp
                                            else:
                                                protocol_temp  =  daily_plan_sheet.at[i,'Protocol']

                                        if (len(execution_projection_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Execution Projection']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Execution Projection']).__contains__('NA')):
                                                execution_projection_temp  =  execution_projection_temp
                                            else:
                                                execution_projection_temp  =  daily_plan_sheet.at[i,'Execution Projection']

                                        if (len(interdomain_kpi_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Interdomain KPI status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Interdomain KPI status']).__contains__('NA')):
                                                interdomain_kpi_status_temp  =  interdomain_kpi_status_temp
                                            else:
                                                interdomain_kpi_status_temp  =  daily_plan_sheet.at[i,'Interdomain KPI status']

                                        if (len(second_level_validation_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Second Level Validation Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Second Level Validation Status']).__contains__('NA')):
                                                second_level_validation_status_temp  =  second_level_validation_status_temp
                                            else:
                                                second_level_validation_status_temp  =  daily_plan_sheet.at[i,'Second Level Validation Status']

                                        if (len(kpi_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'KPI status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'KPI status']).__contains__('NA')):
                                                kpi_status_temp  =  kpi_status_temp
                                            else:
                                                kpi_status_temp  =  daily_plan_sheet.at[i,'KPI status']

                                        if (len(mop_view_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'MOP View Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'MOP View Status']).__contains__('NA')):
                                                mop_view_status_temp  =  mop_view_status_temp
                                            else:
                                                mop_view_status_temp  =  daily_plan_sheet.at[i,'MOP View Status']
                            
                                elif (count == 1):
                                    if (daily_plan_sheet.at[i,'CR NO'] == cr):
                                        
                                        if (len(daily_plan_sheet.at[i,'IMPACTED NODE'].strip()) == 0) or (daily_plan_sheet.at[i,'IMPACTED NODE'].__contains__('NA')) or (daily_plan_sheet.at[i,'IMPACTED NODE'].__contains__('na')):
                                                    impacted_node_temp = impacted_node_temp
                                        else:
                                            if (len(impacted_node_temp) == 0):
                                                impacted_node_temp = str(daily_plan_sheet.at[i,'IMPACTED NODE'])
                                        
                                        if (len(domain_kpi_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Domain kpi']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Domain kpi']).__contains__('NA')):
                                                domain_kpi_temp = domain_kpi_temp
                                            else:
                                                domain_kpi_temp = daily_plan_sheet.at[i,'Domain kpi']
                                        
                                        if (len(daily_plan_sheet.at[i,'KPI DETAILS'].strip()) == 0) or (daily_plan_sheet.at[i,'KPI DETAILS'].__contains__('NA')) or (daily_plan_sheet.at[i,'KPI DETAILS'].__contains__('na')):
                                            kpi_details_temp = kpi_details_temp
                                        else:
                                            if (len(kpi_details_temp) == 0):
                                                kpi_details_temp = str(daily_plan_sheet.at[i,'KPI DETAILS'])
                                        
                                        if (len(str(daily_plan_sheet.at[i,'oss name']).strip()) == 0) or (str(daily_plan_sheet.at[i,'oss name']).__contains__('NA')) :
                                            oss_name_temp = oss_name_temp   
                                        else: 
                                            oss_name_temp  =  daily_plan_sheet.at[i,'oss name']
                                        
                                        if (len(str(daily_plan_sheet.at[i,'oss ip']).strip()) == 0) or (str(daily_plan_sheet.at[i,'oss ip']).__contains__('NA')) :
                                            oss_IP_temp = oss_IP_temp
                                        else:
                                            oss_IP_temp  =  daily_plan_sheet.at[i,'oss ip']

                                        if (len(maintenance_window_temp)) == 0:
                                            if (len(str(daily_plan_sheet.at[i,'Maintenance Window']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Maintenance Window']).__contains__('NA')):
                                                maintenance_window_temp = maintenance_window_temp
                                            else:
                                                maintenance_window_temp  =  daily_plan_sheet.at[i,'Maintenance Window']

                                        if(len(activity_title_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity Title']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity Title']).__contains__('NA')):
                                                activity_title_temp = activity_title_temp
                                            else:
                                                activity_title_temp  =  daily_plan_sheet.at[i,'Activity Title']

                                        if(len(risk_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Risk']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Risk']).__contains__('NA')):
                                                risk_temp = risk_temp
                                            else:
                                                risk_temp  =  daily_plan_sheet.at[i,'Risk']

                                        if (len(location_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Location']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Location']).__contains__('NA')):
                                                location_temp = location_temp
                                            else:    
                                                location_temp  =  daily_plan_sheet.at[i,'Location']

                                        if (len(circle_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Circle']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Circle']).__contains__('NA')):
                                                circle_temp = circle_temp
                                            else:
                                                circle_temp  =  daily_plan_sheet.at[i,'Circle']
                                        
                                        if (len(str(no_of_node_involved_temp)) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'No. of Node Involved']).strip()) == 0) or (str(daily_plan_sheet.at[i,'No. of Node Involved']).__contains__('NA')):
                                                no_of_node_involved_temp = no_of_node_involved_temp
                                            else:
                                                no_of_node_involved_temp  =  daily_plan_sheet.at[i,'No. of Node Involved']
                                        
                                        if (len(cr_belongs_to_same_activity_of_previous_cr_yes_no_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'CR Belongs to Same Activity of Previous CR- Yes/NO']).strip()) == 0) or (str(daily_plan_sheet.at[i,'CR Belongs to Same Activity of Previous CR- Yes/NO']).__contains__('NA')):
                                                cr_belongs_to_same_activity_of_previous_cr_yes_no_temp = cr_belongs_to_same_activity_of_previous_cr_yes_no_temp
                                            else:
                                                cr_belongs_to_same_activity_of_previous_cr_yes_no_temp  =  daily_plan_sheet.at[i,'CR Belongs to Same Activity of Previous CR- Yes/NO']
                            
                                        if (len(change_responsible_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Change Responsible']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Change Responsible']).__contains__('NA')):
                                                change_responsible_temp = change_responsible_temp
                                            else:
                                                change_responsible_temp =  daily_plan_sheet.at[i,'Change Responsible']

                                        if (len(activity_checker_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity Checker']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity Checker']).__contains__('NA')):
                                                activity_checker_temp = activity_checker_temp
                                            else:
                                                activity_checker_temp  =  daily_plan_sheet.at[i,'Activity Checker']

                                        if (len(activity_initiator_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity Initiator']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity Initiator']).__contains__('NA')):
                                                activity_initiator_temp = activity_initiator_temp
                                            else:
                                                activity_initiator_temp  =  daily_plan_sheet.at[i,'Activity Initiator']

                                        if (len(impact_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Impact']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Impact']).__contains__('NA')):
                                                impact_temp = impact_temp
                                            else:
                                                impact_temp  =  daily_plan_sheet.at[i,'Impact']

                                        if (len(planning_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Planning Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Planning Status']).__contains__('NA')):
                                                planning_status_temp = planning_status_temp
                                            else:
                                                planning_status_temp  =  daily_plan_sheet.at[i,'Planning Status']

                                        if (len(domain_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Domain']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Domain']).__contains__('NA')):
                                                domain_temp = domain_temp
                                            else:
                                                domain_temp  =  daily_plan_sheet.at[i,'Domain']

                                        if (len(final_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Final Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Final Status']).__contains__('NA')):
                                                final_status_temp = final_status_temp
                                            else:
                                                final_status_temp  =  daily_plan_sheet.at[i,'Final Status']

                                        if (len(reason_for_rollback_cancel_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Reason For Rollback / Cancel']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Reason For Rollback / Cancel']).__contains__('NA')):
                                                reason_for_rollback_cancel_temp  =  reason_for_rollback_cancel_temp
                                            else:
                                                reason_for_rollback_cancel_temp  =  daily_plan_sheet.at[i,'Reason For Rollback / Cancel']

                                        if (len(design_availability_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Design Availability']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Design Availability']).__contains__('NA')):
                                                design_availability_temp  =  design_availability_temp
                                            else:
                                                design_availability_temp  =  daily_plan_sheet.at[i,'Design Availability']

                                        if (len(technical_validator_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Technical Validator']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Technical Validator']).__contains__('NA')):
                                                technical_validator_temp  =  technical_validator_temp
                                            else:
                                                technical_validator_temp  =  daily_plan_sheet.at[i,'Technical Validator']

                                        if (len(complexity_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Complexity']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Complexity']).__contains__('NA')):
                                                complexity_temp  =  complexity_temp
                                            else:
                                                complexity_temp  =  daily_plan_sheet.at[i,'Complexity']

                                        if (len(activity_type_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Activity-Type']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Activity-Type']).__contains__('NA')):
                                                activity_type_temp  =  activity_type_temp
                                            else:
                                                activity_type_temp  =  daily_plan_sheet.at[i,'Activity-Type']

                                        if (len(total_time_spent_on_planned_crs_mins_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Total Time spent on Planned CRs (Mins)']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Total Time spent on Planned CRs (Mins)']).__contains__('NA')):
                                                total_time_spent_on_planned_crs_mins_temp  =  total_time_spent_on_planned_crs_mins_temp
                                            else:
                                                total_time_spent_on_planned_crs_mins_temp  =  daily_plan_sheet.at[i,'Total Time spent on Planned CRs (Mins)']

                                        if (len(vendor_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Vendor']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Vendor']).__contains__('NA')):
                                                vendor_temp  =  vendor_temp
                                            else:
                                                vendor_temp  =  daily_plan_sheet.at[i,'Vendor']

                                        if (len(protocol_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Protocol']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Protocol']).__contains__('NA')):
                                                protocol_temp  =  protocol_temp
                                            else:
                                                protocol_temp  =  daily_plan_sheet.at[i,'Protocol']

                                        if (len(execution_projection_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Execution Projection']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Execution Projection']).__contains__('NA')):
                                                execution_projection_temp  =  execution_projection_temp
                                            else:
                                                execution_projection_temp  =  daily_plan_sheet.at[i,'Execution Projection']

                                        if (len(interdomain_kpi_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Interdomain KPI status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Interdomain KPI status']).__contains__('NA')):
                                                interdomain_kpi_status_temp  =  interdomain_kpi_status_temp
                                            else:
                                                interdomain_kpi_status_temp  =  daily_plan_sheet.at[i,'Interdomain KPI status']

                                        if (len(second_level_validation_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'Second Level Validation Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'Second Level Validation Status']).__contains__('NA')):
                                                second_level_validation_status_temp  =  second_level_validation_status_temp
                                            else:
                                                second_level_validation_status_temp  =  daily_plan_sheet.at[i,'Second Level Validation Status']

                                        if (len(kpi_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'KPI status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'KPI status']).__contains__('NA')):
                                                kpi_status_temp  =  kpi_status_temp
                                            else:
                                                kpi_status_temp  =  daily_plan_sheet.at[i,'KPI status']

                                        if (len(mop_view_status_temp) == 0):
                                            if (len(str(daily_plan_sheet.at[i,'MOP View Status']).strip()) == 0) or (str(daily_plan_sheet.at[i,'MOP View Status']).__contains__('NA')):
                                                mop_view_status_temp  =  mop_view_status_temp
                                            else:
                                                mop_view_status_temp  =  daily_plan_sheet.at[i,'MOP View Status']
                                
                            counter +=  1
                    
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
                    'Interdomain KPI status':interdomain_kpi_status,
                    'Second Level Validation Status':second_level_validation_status,
                    'KPI status':kpi_status,
                    'MOP View Status':mop_view_status
                    }
                df = pd.DataFrame(dictionary1)
                df['Execution Date'] = df['Execution Date'].dt.strftime('%d-%m-%Y')
                
                writer = pd.ExcelWriter(workbook,engine = 'openpyxl',mode = 'a',if_sheet_exists = 'replace')
                new_sheetname = 'Email-Package'
                df.index +=  1
                #print(df)
                df.to_excel(writer,sheet_name = new_sheetname,index_label = 'S.NO')
                
                writer.save()
                
                styling(workbook,new_sheetname)

                messagebox.showinfo("   Successful Completion",'Email Package Sheet also created')


                

            


    # except FileNotFoundError:
    #     working_directory = r"C:\Daily"
    #     messagebox.showerror("  File not Found","Check {} for MPBN Daily Planning Sheet.xlsx".format(working_directory))
    #     sys.exit(0)
    
    # except ValueError:
    #     working_directory = r"C:\Daily"
    #     messagebox.showwarning("   Value Error","Check {} for MPBN Daily Planning Sheet.xlsx for all the requirement sheet".format(working_directory))
    #     sys.exit(0)

    

except TomorrowDataNotFound as error:
    messagebox.showerror("  Data for tomorrow's date not found",error)
    sys.exit(0)

except DomainNotFound as error:
    messagebox.showerror("  Domain KPI can't be Empty")

#paco_cscore("Enjoy Maity",r"C:\Daily\MPBN Daily Planning Sheet.xlsx")