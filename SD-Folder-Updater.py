"""
This file looks in the W:/Staff Downloads folder for the files in the "dates" list. If it finds them, nothing
happens. If nothing is found for a specific date, it will pull the file from the relevant folder and perform some basic
manipulation on columns to make them more manageable (e.g. moving pay number field to far left)
"""
import pandas as pd
import os

dates = ["Jul-15", "Aug-15", "Sep-15", "Oct-15", "Nov-15", "Dec-15", "Jan-16",
         "Feb-16", "Mar-16", "Apr-16", "May-16", "Jun-16", "Jul-16", "Aug-16",
         "Sep-16", "Oct-16", "Nov-16", "Dec-16", "Jan-17", "Feb-17", "Mar-17",
         "Apr-17", "May-17", "Jun-17", "Jul-17", "Aug-17", "Sep-17", "Oct-17",
         "Nov-17", "Dec-17", "Jan-18", "Feb-18", "Mar-18", "Apr-18", "May-18",
         "Jun-18", "Jul-18", "Aug-18", "Sep-18", "Oct-18", "Nov-18", "Dec-18",
         "Jan-19", "Feb-19", "Mar-19", "Apr-19", "May-19", "Jun-19", "Jul-19",
         "Aug-19", "Sep-19", "Oct-19", "Nov-19", "Dec-19", "Jan-20", "Feb-20",
         "Mar-20", "Apr-20"]  # add dates in this format for each new month


def excelfile(x):
    # list all files in relevant download dir
    files = os.listdir('//ntserver5/generalDB/WorkforceDB/Workforce Monthly Reports/Monthly_Reports/' + str(x) +
                       ' Snapshot/Staff Download/')
    # deal with the various old naming formats
    for i in files:
        if "GGC" in i:
            return ('//ntserver5/generalDB/WorkforceDB/Workforce Monthly Reports/Monthly_Reports/' + str(x) +
                    ' Snapshot/Staff Download/' + str(i))
        elif "GG&C" in i:
            return ('//ntserver5/generalDB/WorkforceDB/Workforce Monthly Reports/Monthly_Reports/' + str(x) +
                    ' Snapshot/Staff Download/' + str(i))
        elif "Staff Download - " + str(x) in i:
            return ('//ntserver5/generalDB/WorkforceDB/Workforce Monthly Reports/Monthly_Reports/' + str(x) +
                    ' Snapshot/Staff Download/' + str(i))


for i in dates:
    # some strange stuff below from before i discovered pd.to_datetime...
    x = pd.Period(i)
    x = str(x).replace('0001-', "")
    x = x.replace("-", "-20")
    x = str.split(x, '-')
    x = x[1] + '-' + x[0] + ' - Staff Download'
    if os.path.isfile('//ntserver5/generalDB/WorkforceDB/Staff Downloads/' + str(x) + '.xlsx'):
        print(x + " - found")
    else:
        print(x + " - no - Creating New File")  # creates new staff download if corresponding staff dl not found above
        df = pd.read_excel(excelfile(i))

        df.set_index('Pay_Number')
        if 'Area_Pay_Division' not in df.columns:
            df['Area_Pay_Division'] = '<Added Later>'
        if 'Contract Planned Contract End Date' not in df.columns:
            df['Contract Planned Contract End Date'] = '<Added Later>'
            # the mental health flag was added later on, so we need to add this placeholder for use in historical
            # employees script
        if 'Mental_Health_Y/N' not in df.columns:
            df['Mental_Health_Y/N'] = '<Added Later>'
            # same as above
        if 'Contract_Description' not in df.columns:
            df['Contract_Description'] = '<Added Later>'

        if 'Division' in df.columns:
            print('renaming columns')
            df = df.rename(columns={'Division': 'Sector/Directorate/HSCP',
                                    'Division_Code': 'Sector/Directorate/HSCP_Code',
                                    'Directorate': 'Sub-Directorate 1',
                                    'sub_directorate': 'Sub-Directorate 2'})
        df = df[['Pay_Number', 'Area', 'Sector/Directorate/HSCP_Code', 'Sector/Directorate/HSCP',
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
                 'Area_Pay_Division', 'Mental_Health_Y/N']]

        writer = pd.ExcelWriter('//ntserver5/generalDB/WorkforceDB/Staff Downloads/' + str(x) + '.xlsx')
        # a1:at1        df.to_excel(writer, 'Sheet1', index=False)
        writer.save()
print("Complete")
