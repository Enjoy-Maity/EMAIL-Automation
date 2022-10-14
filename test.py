from unittest import result
import pandas as pd 

workbook = r"C:\Daily\MPBN Daily Planning Sheet.xlsx"
workbook = pd.ExcelFile(workbook)
worksheet =  pd.read_excel(workbook,"Planning Sheet")

#worksheet.reset_index(drop= True,inplace = True)
result_df = pd.DataFrame()
#print(worksheet_unique_cr)
worksheet = worksheet[['S.NO','Execution Date','Maintenance Window','CR NO','Activity Title','Risk','Location','Circle']]
worksheet_unique_cr = worksheet['CR NO'].value_counts().index.to_list()

input_error = []
for index,cr_no in enumerate(worksheet_unique_cr):
    counter= worksheet['CR NO'].value_counts().to_list()[index]
    temp_df = worksheet[worksheet["CR NO"] == cr_no]
    if (counter > 1):
        if (len(temp_df['Circle'].unique()) > 1):
            for i in range(0, len(temp_df)):
                input_error.append(temp_df.at[i,'S.NO'])
            continue
        if (len(temp_df['Circle'].unique()) == 1):
            result_df = pd.concat([result_df,temp_df.iloc[0].to_frame().T],ignore_index= True)
    else:
        result_df = pd.concat([result_df,temp_df.iloc[0].to_frame().T],ignore_index= True)

new_result_df = pd.DataFrame()
for i in range(0,len(result_df)):
    if (result_df.at[i,'Activity Title'] == 'NaN'):
        continue
    
    if (result_df.at[i,'CR NO'] == 'NaN'):
        continue

    if (result_df.at[i,'Circle'] == 'NaN'):
        continue

    else:
        new_result_df = pd.concat([new_result_df,result_df.iloc[i].to_frame().T],ignore_index = True)



#result_df = worksheet
print(input_error)
    
#print(result_df)
print("\n")