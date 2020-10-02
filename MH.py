import pandas as pd
from tkinter.filedialog import askopenfilename
import requests
from requests_ntlm import HttpNtlmAuth
from collections import defaultdict
import numpy as np

pd.set_option('display.width', 420)
pd.set_option('display.max_columns', 10)

start = pd.Timestamp.now()
learnpro_runtime = input("when was this data pulled from learnpro? (format = dd-mm-yy)")
learnpro_date = pd.to_datetime(learnpro_runtime, format='%d-%m-%y')
# todo SCOPING
# get the scopes file from same place as before, but this time aim to go for the M&H bits

username = 'xggc\\' + input("GGC username?")
password = input("GGC password")


def getScope():
    response = requests.get(
        'http://teams.staffnet.ggc.scot.nhs.uk/teams/CorpSvc/HR/WorkInfo/HSE%20Resources/HSE%20Scope%20List.xlsx',
        auth=HttpNtlmAuth(username, password))

    with open('W:/LearnPro/HSEScope-current.xlsx', 'wb') as f:
        f.write(response.content)

    curr_scope = pd.read_excel("W:/LearnPro/HSEScope-current.xlsx", skiprows=1)
    print(curr_scope.columns)

    curr_scope = curr_scope[['Cost Centre', 'Administrative Services.1', 'Allied Health Profession.1',
                             'Dental Support.1', 'Executives.1', 'Healthcare Sciences.1',
                             'Medical And Dental.1', 'Medical Support.1', 'Nursing and Midwifery.1',
                             'Other Therapeutic.1', 'Personal and Social Care.1', 'Support Services.1']]
    renames = {'Administrative Services.1': 'Administrative Services',
               'Allied Health Profession.1': 'Allied Health Profession',
               'Dental Support.1': 'Dental Support',
               'Healthcare Sciences.1': 'Healthcare Sciences',
               'Executives.1': 'Executives',
               'Medical And Dental.1': 'Medical and Dental',
               'Medical Support.1': 'Medical Support',
               'Nursing and Midwifery.1': 'Nursing and Midwifery',
               'Other Therapeutic.1': 'Other Therapeutic',
               'Personal and Social Care.1': 'Personal and Social Care',
               'Support Services.1': 'Support Services'}
    curr_scope.rename(columns=renames, inplace=True)
    print(curr_scope.columns)
    return curr_scope


def workwithScopes(df):
    """This func goes through the HSE Scope sheet and returns a concatenated instruction of who to include in scope"""
    print(len(df))

    fams = ['Administrative Services', 'Allied Health Profession',
            'Dental Support', 'Executives', 'Healthcare Sciences',
            'Medical and Dental', 'Medical Support', 'Nursing and Midwifery',
            'Other Therapeutic', 'Personal and Social Care', 'Support Services']

    df.set_index('Cost Centre', inplace=True)



    df.dropna(subset=fams, inplace=True, how='all')
    print(len(df))
    output = []
    for i in df.itertuples():
        count = 0
        for j in i[1:]:
            count = count + 1
            if isinstance(j, str):
                data = (i[0] + fams[count - 1])
                output.append(data)
            else:
                continue

    return output


def merge_scopes(concat):
    sd = pd.read_excel('W:/Staff Downloads/2020-08 - Staff Download.xlsx')

    sd['Concat'] = sd['Cost_Centre'] + sd['Job_Family']
    sd['Scope'] = 1
    sd['Scope'].loc[~sd['Concat'].isin(concat)] = 0
    print(sd.columns)

    return sd

def exclusions(sd):
    marisa_011020 = ['G0530247', 'G0001226', 'G9200959', 'G5820162', 'G5843480', 'G9490728',
                     'G9293493', 'G5898005', 'G5913004', 'G5925134', 'G5925045', 'G5926807',
                     'G1094505', 'G0988448', 'G0001854', 'G9356738', 'G9860558', 'G9867672',
                     'G5878527', 'G0638080', 'G0510297', 'G0799661']

    sd['Scope'].loc[(sd['Pay_Number'].isin(marisa_011020))] = 0
    print(sd['Scope'].loc[sd['Pay_Number'] == 'G0988448'])
    return sd


# todo EMPOWER
def empower():
    emp_data = pd.read_excel('W:/Learnpro/M&H/Empower Data.xlsx')
    #emp_data.drop_duplicates(subset=['payroll_number', 'course_name'], inplace=True)
    emp_data = defaultdict(lambda:0, dict(emp_data['payroll_number'].value_counts()))
    emp_assessors = pd.read_excel('W:/Learnpro/M&H/Empower Assessors.xlsx')
    #emp_assessors.drop_duplicates(subset=['payroll_number', 'course_name'], inplace=True)
    emp_assessors = defaultdict(lambda:0, dict(emp_assessors['payroll_number'].value_counts()))
    print(f'Data file length: {len(emp_data)}, Assessors file length: {len(emp_assessors)}')
    print(f'G9860587 : {emp_data.get("G0833983")}')
    print(f'G9878496: {emp_assessors.get("G0833983")}')
    return emp_data, emp_assessors


def eESS():
    filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/M&H',
                               title="Choose the relevant M&H Training Extract file"
                               )
    eESS_data = pd.read_csv(filename, encoding='utf-16', sep='\t')
    lookup = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data',
                               title="Choose the pay number lookup file"
                               )
    pay_number_lookup = pd.read_csv(lookup, encoding='utf-16', sep='\t')
    pay_number_lookup.rename(columns={'Payroll Number':'Pay Number'}, inplace=True)
    eESS_data = eESS_data.merge(pay_number_lookup, how='left', on='Assignment Number')

    #eESS_data = eESS_data.drop_duplicates(subset=['Pay Number', 'Course Name'])
    eESS_data = defaultdict(lambda:0,dict(eESS_data['Pay Number'].value_counts()))
    filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/M&H',
                               title="Choose the relevant eESS Assessors file"
                               )
    eESS_assessors = pd.read_excel(filename)
    print(eESS_assessors.columns)
    eESS_assessors = defaultdict(lambda: 0, dict(eESS_assessors['Payroll Number'].value_counts()))
    print(eESS_assessors)
    return eESS_data, eESS_assessors

def map_data(sd):
    sd['Assessed-Empower'] = sd['Pay_Number'].map(emp_data)
    sd['Assessed-eESS'] = sd['Pay_Number'].map(eESS_data)
    sd['Assessor Course-eESS'] = sd['Pay_Number'].map(eESS_assessors)
    sd['Assessor Course-Empower'] = sd['Pay_Number'].map(emp_assessors)
    sd['Assessor - Both'] = sd['Assessor Course-eESS'] + sd['Assessor Course-Empower']
    sd['Assessed-Total'] = sd['Assessed-Empower'] + sd['Assessed-eESS']
    sd['Assessed Category'] = ''
    sd['Assessed Category'] = pd.cut(sd['Assessed-Total'], bins=[-1,0, 1, 2, 1000], labels=['0', '1', '2', '3 or more'])
    sd['Assessed Category'].cat.add_categories('Out of scope', inplace=True)
    sd['Assessed Category'].loc[sd['Scope'] == 0] = 'Out of scope'
    print(sd.columns)

    correct_headings = ['Sector/Directorate/HSCP_Code', 'Scope',
       'Assessor Course-eESS', 'Assessor Course-Empower', 'Assessor - Both',
       'Assessed-Empower', 'Assessed-eESS', 'Assessed-Total',
       'Assessed Category', 'Area', 'Sector/Directorate/HSCP',
       'Sub-Directorate 1', 'Sub-Directorate 2', 'department', 'Cost_Centre',
       'Pay_Number', 'Surname', 'Forename', 'Base', 'Job_Family_Code',
       'Job_Family', 'Sub_Job_Family', 'Post_Descriptor', 'Conditioned_Hours',
       'Contracted_Hours', 'WTE', 'Contract_Description', 'NI_Number', 'Age',
       'Date_of_Birth', 'Date_Started', 'Contract Planned Contract End Date',
       'Annual_Salary', 'Date_To_Grade', 'Date_Superannuation_Started',
       'SB_Number', 'Sick_Date_Entitlement_From', 'Description',
       'Marital_Status', 'Sex', 'Job_Description', 'Grade', 'Group_Code',
       'Pay_Scale', 'Pay_Band', 'Scale_Point', 'Pay_Point', 'Incremental Date',
       'Address_Line_1', 'Address_Line_2', 'Address_Line_3', 'Postcode',
       'Area_Pay_Division', 'Mental_Health_Y/N']
    sd['Assessor - Both'] = np.where(sd['Assessor - Both'] > 0, 1, 0)
    sd = sd[correct_headings]
    mh_piv = pd.pivot_table(sd[sd['Scope'] == 1],index=['Area','Sector/Directorate/HSCP'],
                               columns='Assessed Category',
                               values='Pay_Number', aggfunc='count', fill_value=0, margins=True,
                               margins_name='All Staff')
    assessor_piv = pd.pivot_table(sd[sd['Scope'] == 1], index=['Area', 'Sector/Directorate/HSCP'],
                                  columns='Assessor - Both',
                                  values='Pay_Number', aggfunc='count', fill_value=0, margins=True,
                                  margins_name='All Staff')
    with pd.ExcelWriter('C:/Tong/test.xlsx') as writer:
        sd.to_excel(writer, index=False, sheet_name='data')
        mh_piv.to_excel(writer, sheet_name='Assessments Pivot')
        assessor_piv.to_excel(writer, sheet_name='Assessor Pivot')
    writer.save()
    #drop cols for sharing
    sd = sd[['Scope', 'Assessor Course-eESS', 'Assessor Course-Empower',
       'Assessor - Both', 'Assessed-Empower', 'Assessed-eESS',
       'Assessed-Total', 'Assessed Category', 'Area',
       'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2',
       'department', 'Cost_Centre', 'Surname', 'Forename', 'Base',
       'Job_Family_Code', 'Job_Family', 'Sub_Job_Family']]
    with pd.ExcelWriter('W:/Learnpro/Named Lists HSE/'+learnpro_date.strftime('%Y-%m-%d')+'/'+
                        learnpro_date.strftime('%Y%m%d')+' - HSE MovingAndHandling.xlsx') as writer:
        sd.to_excel(writer, index=False, sheet_name='data')
        mh_piv.to_excel(writer, sheet_name='Assessments Pivot')
        assessor_piv.to_excel(writer, sheet_name='Assessor Pivot')
    writer.save()
# todo make a version for M&H folder as well as a more censored version for sharepoint

eESS_data, eESS_assessors = eESS()
emp_data, emp_assessors = empower()

scope = getScope()
concats = workwithScopes(scope)
print(len(concats))
sd = merge_scopes(concats)
sd = exclusions(sd)
map_data(sd)


end = pd.Timestamp.now()
print(end - start)
