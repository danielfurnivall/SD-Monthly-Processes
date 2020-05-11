"""This script reads through all the files in the W:/Staff Downloads directory, opening each file and appending it
    to a master file. It then drops all duplicates, keeping the most recent - thus providing an up-to-date list
    of all historical employees during the recorded duration, including the final position of leavers."""

import pandas as pd
import os, psutil
from datetime import datetime

today = datetime.today().strftime('%Y-%m-%d')
Leaver_file = pd.read_excel(r'\\ntserver5\generaldb\workforcedb\Starters & Leavers\Starters and Leavers - Apr 16 - Feb 20.xlsx',
                            sheet_name='Leavers - Apr-16 - Present')
df = pd.DataFrame(columns=['Pay_Number', 'Area', 'Sector/Directorate/HSCP_Code', 'Sector/Directorate/HSCP',
                           'Sub-Directorate 1', 'Sub-Directorate 2', 'department', 'Cost_Centre',
                           'Surname', 'Forename', 'Base', 'Job_Family_Code',
                           'Job_Family', 'Sub_Job_Family', 'Post_Descriptor', 'Conditioned_Hours',
                           'Contracted_Hours', 'WTE', 'Contract_Description', 'NI_Number', 'Age',
                           'Date_of_Birth', 'Date_Started', 'Contract Planned Contract End Date',
                           'Annual_Salary', 'Date_To_Grade', 'Date_Superannuation_Started',
                           'SB_Number', 'Sick_Date_Entitlement_From', 'Description',
                           'Marital_Status', 'Sex', 'Job_Description', 'Grade', 'Group_Code',
                           'Pay_Scale', 'Pay_Band', 'Scale_Point', 'Pay_Point', 'Incremental Date',
                           'Address_Line_1', 'Address_Line_2', 'Address_Line_3', 'Postcode',
                           'Area_Pay_Division', 'Mental_Health_Y/N'])


for filename in os.listdir('//ntserver5/generalDB/WorkforceDB/Staff Downloads/'):
    print(filename)
    current = pd.read_excel('//ntserver5/generalDB/WorkforceDB/Staff Downloads/' + filename)
    pid = os.getpid()
    ps=psutil.Process(pid)
    print(ps.memory_info())
    df = pd.concat([current, df])
    df.drop_duplicates('Pay_Number', keep='last', inplace=True)
    print(len(df))
df = df.merge(Leaver_file[['Pay_Number', 'Date Left', 'Leaving Description']], on='Pay_Number', how='left')

df.to_csv('W://Historical Employees/uniques ' + str(today) + '.csv', index=False)

print('file available at W://Historical Employees/uniques ' + str(today) + '.csv')