"""This file takes in stat mand, sharps, falls, m&H compliance and plonks them in a single sheet which is used to inform
the emailer file (uKan) developed by Benjamin Law"""

import pandas as pd
from tkinter.filedialog import askopenfilename
import numpy as np

# This is used to revert back to the correct naming schema for the email files
SM_Lookup = {'SM2': 'Fire Awareness', 'SM3': 'Health, Safety & Welfare', 'SM9': 'Violence and Aggression',
 'SM1': 'Equality, Diversity and Human Rights', 'SM6': 'Manual Handling', 'SM4': 'Infection Control',
 'SM7': 'Public Protection', 'SM8': 'Security and Threat', 'SM5': 'Information Governance'}


def take_in_stat_mand():


    filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/Named Lists',
                           title="Choose the relevant stat/mand file - "
                                 "it should be the 'GGC Pay Nums' file for the relevant month."
                           )

    df = pd.read_excel(filename)

    df.rename(columns={'ID Number': 'Pay_Number', 'First': 'Forename', 'Last': 'Surname'}, inplace=True)
    print(df.columns)
    df.drop(columns=SM_Lookup.values(), inplace=True)
    df.rename(columns=SM_Lookup, inplace=True)
    print(df.columns)
    print(df['Fire Awareness'].value_counts())
    print(len(df))
    df['LP Account'] = np.where((~(df['Pay_Number'].isin(users_file()))) & (df[list(SM_Lookup.values())].sum(axis=1) == 0), 0, 1)
    print(df['LP Account'].value_counts())
    premerge = len(df)
    df = df.merge(staff_download(), on='Pay_Number', how='left')
    df = df.merge(falls(), on='Pay_Number', how='left')
    df = df.merge(sharps(), on='Pay_Number', how='left')
    df = df.merge(MH(), on='Pay_Number', how='left')
    df = df.merge(manager_emails(), how='left')
    postmerge = len(df)
    df['SectorLookup'] = df['Sector/Directorate/HSCP']
    print(f'Pre-merge: {premerge}, Post-merge:{postmerge}')
    df = df[['Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1',
       'Sub-Directorate 2', 'department', 'Cost_Centre', 'Pay_Number',
       'Surname', 'Forename', 'Base', 'Job_Family', 'Sub_Job_Family',
       'Post_Descriptor', 'WTE', 'Contract_Description', 'NI_Number',
       'Date_of_Birth', 'Date_Started', 'Work Email Address',
       'Manager Payroll', 'LP Account', 'Equality, Diversity and Human Rights',
       'Fire Awareness', 'Health, Safety & Welfare', 'Infection Control',
       'Information Governance', 'Manual Handling', 'Public Protection',
       'Security and Threat', 'Violence and Aggression',
       'Equality, Diversity and Human Rights expires on...',
       'Fire Awareness expires on...',
       'Health, Safety & Welfare expires on...',
       'Infection Control expires on...',
       'Information Governance expires on...', 'Manual Handling expires on...',
       'Public Protection expires on...', 'Security and Threat expires on...',
       'Violence and Aggression expires on...', 'SectorLookup', 'GGC Module',
       'GGC Date', 'NES Module', 'NES Date', 'Falls Compliant', 'Falls Date',
       'M and H Compliant', 'M and H Date']]
    print("If it reaches here, all cols correct")
    df.to_excel('W:/ukan/danny_staff_files/'+pd.Timestamp.now().strftime('%Y%m%d')+' - Staff File.xlsx', index=False)




def users_file():
    filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data',
                           title="Choose the relevant stat/mand USERS file - "
                           )
    df = pd.read_csv(filename, skiprows=11, sep='\t')
    print(df.columns)
    return df['ID Number'].unique().tolist()

def staff_download():
    filename = askopenfilename(initialdir = 'W:/Staff Downloads',
                               title = 'Choose the current staff download')
    df = pd.read_excel(filename)
    print(df.columns)
    df = df[['Pay_Number','Base','Post_Descriptor', 'WTE', 'Contract_Description', 'NI_Number',
       'Date_of_Birth', 'Date_Started']]
    return df

def falls():
    filename = askopenfilename(initialdir='W:/Learnpro/Named Lists HSE',
                               title='Choose the current falls file (this must be the one with pay number!')
    df = pd.read_excel(filename)
    df = df[['Pay_Number', 'Falls Compliant', 'Falls Date']]

    df['Falls Compliant'].loc[~df['Falls Compliant'].isin(['Out of scope', 'Complete'])] = '0'
    df['Falls Compliant'].loc[df['Falls Compliant'] == 'Complete'] = '1'
    print(df['Falls Compliant'].value_counts())
    df['Falls Date'] = df['Falls Date'].dt.strftime('%d/%m/%Y')
    print(df['Falls Date'].head(5))
    return df

def sharps():
    filename = askopenfilename(initialdir='W:/Learnpro/Named Lists HSE',
                               title='Choose the current Sharps file (this must be the one with pay number!')
    df = pd.read_excel(filename)
    print(df.columns)
    print(len(df))
    df = df[['Pay_Number', 'GGC Module', 'GGC Date', 'NES Module', 'NES Date']]
    print(df['NES Date'].head(5))
    print(df['GGC Date'].head(5))
    mods = ['GGC Module', 'NES Module']
    for i in mods:
        df[i].loc[~df[i].isin(['Out Of Scope', 'Complete'])] = 0
        df[i].loc[df[i] == 'Complete'] = 1
    df['GGC Date'] = df['GGC Date'].dt.strftime('%d/%m/%Y')
    df['NES Date'] = df['NES Date'].dt.strftime('%d/%m/%Y')
    return df

def MH():
    filename = askopenfilename(initialdir='W:/Learnpro/M&H/Staff Download (working files)',
                               title='Choose the current M&H file')
    df = pd.read_excel(filename)
    print(df.columns)
    df = df[['Pay_Number', 'Assessed Category']]
    print(df['Assessed Category'].value_counts())
    df['Assessed Category'].loc[df['Assessed Category'].isin(['3 or more', 2, 1 ])] = 'Complete'
    df['Assessed Category'].loc[df['Assessed Category'] == 0] = 'Not Undertaken'
    df = df.rename(columns={'Assessed Category':'M and H Compliant'})
    df['M and H Date'] = ''
    print(df['M and H Compliant'].value_counts())
    print(df.columns)
    return df

def manager_emails():
    filename = askopenfilename(initialdir= 'W:/ukan/eESS email extracts',
                               title='Choose the correct manager email file')
    df = pd.read_excel(filename)
    print(df.columns)
    manager_lookup = dict(zip(df['Employee Number'], df['Payroll Number']))
    print(len(manager_lookup))
    df['Manager Payroll'] = df['Supervisor ID'].map(manager_lookup)
    df.rename(columns={'Payroll Number':'Pay_Number'}, inplace=True)
    df = df[['Pay_Number','Work Email Address', 'Manager Payroll']]
    return df




# emails = manager_emails()
# # sd = staff_download()
# # users = users_file()
# # falls = falls()
# # sharps = sharps()
# # mh = MH()
stat_mand = take_in_stat_mand()


ukan_cols = ['Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1',
       'Sub-Directorate 2', 'department', 'Cost_Centre', 'Pay_Number',
       'Surname', 'Forename', 'Base', 'Job_Family', 'Sub_Job_Family',
       'Post_Descriptor', 'WTE', 'Contract_Description', 'NI_Number',
       'Date_of_Birth', 'Date_Started', 'Work Email Address',
       'Manager Payroll', 'LP Account', 'Equality, Diversity and Human Rights',
       'Fire Awareness', 'Health, Safety & Welfare', 'Infection Control',
       'Information Governance', 'Manual Handling', 'Public Protection',
       'Security and Threat', 'Violence and Aggression',
       'Equality, Diversity and Human Rights expires on...',
       'Fire Awareness expires on...',
       'Health, Safety & Welfare expires on...',
       'Infection Control expires on...',
       'Information Governance expires on...', 'Manual Handling expires on...',
       'Public Protection expires on...', 'Security and Threat expires on...',
       'Violence and Aggression expires on...', 'SectorLookup', 'GGC Module',
       'GGC Date', 'NES Module', 'NES Date', 'Falls Compliant', 'Falls Date',
       'M and H Compliant', 'M and H Date']