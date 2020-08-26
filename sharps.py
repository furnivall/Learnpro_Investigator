"""
This file builds on the work already done with the statutory/mandatory portion of the learnpro process and produces a
further report on Sharps & Skins data.
"""

import pandas as pd
from tkinter.filedialog import askdirectory
import os
import datetime as dt
import numpy as np
import requests
from requests_ntlm import HttpNtlmAuth
import configparser

sharps_courses = ['GGC: Management of Needlestick & Similar Injuries',
                  # 'GGC: Prevention & Management of Occ. Exposure Pt 2',
                  'NES: Prev. and Mgmt. of Occ. Exposure (Assessment)']

username = 'xggc\\' + input("GGC username?")
password = input("GGC password")
print(username)
print(password)


def getHSEScopeFile():
    response = requests.get(
        'http://teams.staffnet.ggc.scot.nhs.uk/teams/CorpSvc/HR/WorkInfo/HSE%20Resources/Staff%20Scope%20List.xlsx',
        auth=HttpNtlmAuth(username, password))

    with open('W:/LearnPro/HSEScope-current.xlsx', 'wb') as f:
        f.write(response.content)

    curr_scope = pd.read_excel("W:/LearnPro/StaffScope-current.xlsx")
    ggc_oos = curr_scope[curr_scope['GGC Module'] == "Out Of Scope"]
    nes_oos = curr_scope[curr_scope['NES Module'] == 'Out Of Scope']

    return ggc_oos['Pay_Number'].tolist(), nes_oos['Pay_Number'].tolist()


def check_users():
    df = pd.read_excel()


def NESScope(df):
    print('Generating NES Scope')
    # Estates & Facilities
    df = df[~(df['Sector/Directorate/HSCP'] == 'Estates and Facilities')]

    # Support Services
    df = df[~(df['Job_Family'] == 'Support Services')]

    # AHPs
    df = df[~(df['Sub_Job_Family'].isin(['Diagnostic Radiography', 'Therapeutic Radiography', 'Orthotics']))]

    df = df[~((df['Area'] == 'Partnerships') & df['Sub_Job_Family'].isin(['Physiotherapy', 'Occupational Therapy']))]

    # Chief Operating Officer - this is already exclude by Acute Corporate in GGC Scope but I've left it here
    df = df[~(df['Sub-Directorate 1'] == 'Chief Operating Officer')]

    # Inverclyde Children and Families
    df = df[~(df['Sub-Directorate 1'] == 'Inverclyde Hscp: Children & Families')]

    # Community Engagement Team - this is already excluded due to job families but it's left in
    df = df[~(df['department'] == 'Community Engagement Team')]

    # Research Management
    df = df[~(df['Sub-Directorate 2'] == 'Research Management')]

    # Biochem Lab
    df = df[~(df['department'] == 'Gri -Biochemistry Lab')]

    print(f'NES Inscope: {len(df)}')
    df.to_excel('C:/Learnpro_Extracts/sharps/nestest.xlsx')
    return df


def GGCScope(df):
    print('Generating GGC Scope')

    df_orig = df

    # Sector Level Exclusions
    df = df[~df['Sector/Directorate/HSCP'].isin(['Acute Corporate', 'Board Medical Director', 'Board Administration',
                                                 'Centre For Population Health', 'Corporate Communications', 'eHealth',
                                                 'Finance', 'Public Health'])]
    df = df[~((df['Sector/Directorate/HSCP'] == 'HR and OD') & (df['Sub-Directorate 1'] != 'Occupational Health'))]

    # Job Family Exclusions
    df = df[~df['Job_Family'].isin(['Administrative Services', 'Personal and Social Care', 'Executive'])]
    df = df[~((df['Job_Family'] == 'Other Therapeutic') & (df['Sub_Job_Family'] != 'Optometry'))]

    # Sub job family exclusions
    df = df[~(df['Sub_Job_Family'].isin(
        ['Speech and Language Therapy', 'Speech And Language Therapy', 'Arts Therapies', 'Dietetics',
         'AHP Training/Administration', 'Generic Therapies']))]

    # Health Visitors
    df = df[~((df['Sub_Job_Family'] == 'Health Visitor Nursing') & (df['Pay_Band'] != '5'))]

    # School Nurses
    df = df[~((df['Sub_Job_Family'] == 'School Nursing') & (df['Pay_Band'] != '5'))]

    # Clinical Physics
    df = df[~(df['Sub-Directorate 2'].isin(['Equipment Management', 'Clinical Physics Radiotherapy Physics',
                                            'Clinical Physics Core']))]

    # Haematology
    df = df[~(df['Sub-Directorate 2'] == 'Haematology')]

    # Dept. level exclusions
    df = df[~((df['Sector/Directorate/HSCP'] == 'Board Nurse Director') &
              (df['department'].isin(
                  ['Nur -Clinical Co-Ordinators', 'Nur -Practice Development', 'Iv Fluids Protocol'])))]

    print(f'GGC Overall: {len(df_orig)},  GGC Inscope: {len(df)}')

    return df


def take_in_dir(list_of_modules):
    """Takes in directory with prompt then selects all learnpro files within that folder"""

    # counters for lp/nonlp files within dir
    lp_count = 0
    non_lp_count = 0

    # prompt for dir

    # TODO swap this back when testing is done
    dirname = askdirectory(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data',
                           title="Choose a directory full of learnpro files."
                           )

    # initialise master dataframe
    master = pd.DataFrame()
    modules = []
    # read in historic empower data

    user_ids = []

    # iterate through directory
    for file in os.listdir(dirname):
        # operate on the LEARNPRO files first
        if 'LEARNPRO' in file:

            print(file)
            lp_count += 1
            # read file into df
            df = pd.read_csv(dirname + "/" + file, skiprows=14, sep="\t")
            print("Initial length: " + str(len(df)))
            modules.append(df['Module'].unique().tolist())
            # drop fails
            df = df[df['Passed'] == "Yes"]
            print("Removed fails - new length: " + str(len(df)))
            df = df[['ID Number', 'Module', 'Assessment Date']]
            # TODO this would be a good place to add any other cleaning steps
            # drop modules not in list
            df = df[df['Module'].isin(list_of_modules)]
            print("Removed extra modules - new length: " + str(len(df)))

            # add data to master file
            master = master.append(df, ignore_index=True)

            print(file + " added to master, current size= " + str(len(master)))
        elif 'users' in file:
            df = pd.read_csv(dirname + "/" + file, skiprows=11, sep="\t")
            user_ids = df['ID Number'].unique().tolist()


        else:
            print("not lp")
            non_lp_count += 1
    master['GGC Source'] = 'Learnpro'
    # log outputs below:
    print(str(lp_count + non_lp_count) + " files read. " + str(lp_count) + " contained learnpro data, " + str(
        non_lp_count) +
          " did not.")
    master['Module'] = master['Module'].astype('category')

    master['Assessment Date'] = pd.to_datetime(master['Assessment Date'], format='%d/%m/%y %H:%M')

    # TODO empower and eESS functions removed - restore them here if necessary

    # create master list of id numbers
    df_users = master['ID Number'].unique().tolist()

    # beautiful and elegant nested list comprehension
    modules = [item for sublist in modules for item in sublist]
    modules = list(dict.fromkeys(modules))

    with open('C:/Learnpro_Extracts/listfile.txt', 'w') as filehandler:
        for listitem in modules:
            filehandler.write('%s\n' % listitem)

    print(master.columns)
    print(master.dtypes)
    master.to_csv('C:/Learnpro_Extracts/bigfile.csv', index=False)
    # exit()
    return master, user_ids


def build_user_compliance_dates(df):
    """Takes in a dataframe with lots of learnpro pass data and produces a dataset of test dates as an output"""
    users = df['ID Number'].drop_duplicates().sort_values().to_frame()
    initial_length = len(users)
    for module in sharps_courses:
        print(module)
        df1 = df[df['Module'] == module]
        df1 = df1.drop(columns="Module")
        df1 = df1.rename(columns={"Assessment Date": module + " Date"})
        print("This module has " + str(len(df1)) + "users")
        users = users.merge(df1, on="ID Number", how="left")
        users = users[users['ID Number'].notnull()]
        # this is necessary because otherwise it'll create dupes for some reason
        users.drop_duplicates(subset='ID Number', inplace=True, keep='last')
        print("Current dataset size: " + str(len(users)) + " users")
        print("Dataframe shape: " + str(users.shape))
    # for debug
    # users.to_excel("C:/Learnpro_Extracts/dates.xlsx", index=False)

    print("Initial length: " + str(initial_length))
    print("Final length: " + str(len(users['ID Number'].drop_duplicates())))
    print(users.columns)

    return users


def sd_merge(df):
    """This function merges in the Staff Download data to let us work with identifiable stuff for our pivot"""

    # legacy - good for excel pivots
    df['Headcount'] = 1

    # TODO point this somewhere else
    sd = pd.read_excel('W:/Staff Downloads/2020-07 - Staff Download.xlsx')

    # Cleaning step - vital for merge to work properly - ID number must be on both sides of merge
    sd = sd.rename(columns={'Pay_Number': 'ID Number', 'Forename': 'First', 'Surname': 'Last'})

    # cut down to useful cols
    sd = sd[['ID Number', 'NI_Number', 'Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2',
             'department', 'Cost_Centre', 'First', 'Last', 'Job_Family', 'Sub_Job_Family', 'Base',
             'Contract_Description', 'Job_Description', 'Post_Descriptor', 'WTE', 'Pay_Band',
             'Date_Started']]

    # merge
    df = df.merge(sd, on='ID Number', how='right')

    return df


def produce_files(df):
    """Builds final files for named list"""
    # List of columns in final file from Ben's sheet:

    print(df.columns)

    df = df.rename(columns={'First':'Forename', 'Last':'Surname', 'GGC':'GGC Module', 'NES':'NES Module',
                            'ID Number':'Pay_Number'})
    df['Compliant'] = 'DO THIS LATER'
    df['GGC Source'] = 'Learnpro'
    df = df[['Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1',
             'Sub-Directorate 2', 'department', 'Cost_Centre', 'Pay_Number',
             'Surname', 'Forename', 'Base', 'Job_Family', 'Sub_Job_Family',
             'Post_Descriptor', 'WTE', 'Contract_Description', 'NI_Number',
             'Date_Started', 'Job_Description', 'Pay_Band', 'GGC Module', 'GGC Date',
             'GGC Source', 'NES Module', 'NES Date', 'Compliant']]

    # write to book
    with pd.ExcelWriter('C:/Learnpro_Extracts/sharps/namedList.xlsx') as writer:
        df.to_excel(writer, sheet_name='data', index=False)
        # TODO add pivot
        # piv.to_excel(writer, sheet_name='pivot')
    writer.save()

    # for debugging - make csv file with all columns and data


def ni_fix(df):
    """This little function does a lot - it gets all NI numbers with duplicates, then creates a dataframe for each dupe,
    checks the compliance for each course across all relevant pay numbers, then gets the most recent course completion
    date then applies these to all of the pay numbers"""

    # find dupes and put in list
    dups = df[df.duplicated(subset='NI_Number')]['NI_Number'].drop_duplicates().to_list()

    # this is to make df.update() be able to compare indices and amend data
    df.set_index('ID Number')

    courses = ['GGC', 'NES']

    # loop through courses, first making all compliant if necessary, then editing expiry dates
    for number, i in enumerate(dups):
        dfx = df[df['NI_Number'] == i]
        for j in courses:
            if 'Complete' in dfx[j].values:
                dfx.loc[:, j] = 'Complete'
            dfx.loc[:, j + ' expires on...'] = dfx[j + ' expires on...'].max()
        df.update(dfx)
    df.reset_index()
    return df


def check_compliance(df, users):
    """Takes in a df with test dates for each of the stat/mand modules, including both versions of safe info handling.
        It will produce a completed named list dataset ala the old CompliancePro"""

    print(df.columns)

    # iterate through modules, grabbing their test dates as appropriate
    for module in sharps_courses:
        columnnames = {'GGC: Management of Needlestick & Similar Injuries': 'GGC',
                       # 'GGC: Prevention & Management of Occ. Exposure Pt 2':'GGC 2',
                       'NES: Prev. and Mgmt. of Occ. Exposure (Assessment)': 'NES'
                       }
        # the courses all have a shorter name to look nicer (see above dictionary)
        short_name = columnnames.get(module)
        df[short_name] = ""
        df[short_name + ' Date'] = df[module + ' Date']
        # attach compliance to all except infogov and fire. It is likely this will need to be changed to capture more
        earliestCompliance = pd.Timestamp.now() - pd.DateOffset(years=2)
        # df[short_name] = np.where(df[module + ' Date'] > earliestCompliance, "Complete", "Not Compliant")
        df[short_name].loc[df[module + ' Date'] > earliestCompliance] = 'Complete'
        df[short_name].loc[(df[module + ' Date'].notnull()) & (df[module + ' Date'] < earliestCompliance)] = 'Expired'

        df[str(short_name) + ' expires on...'] = df[module + ' Date'] + pd.DateOffset(years=2)
        df[short_name + ' expires on...'].loc[df[short_name] == 'Not Compliant'] = None
        print(users)

    # merge staff download cols into dataset
    df = sd_merge(df)

    # fix for people with more than one pay number

    df = ni_fix(df)

    # df['NES - Scope'] = ""
    # df['GGC - Scope'] = ""
    df['NES'].loc[(df['NES'].isnull()) & (~df['ID Number'].isin(nes['Pay_Number'].unique().tolist()))] = 'Out Of Scope'
    df['GGC'].loc[(df['GGC'].isnull()) & (~df['ID Number'].isin(ggc['Pay_Number'].unique().tolist()))] = 'Out Of Scope'
    df['NES'].loc[(~df['ID Number'].isin(users)) & (df['NES'].isnull())] = 'No Account'
    df['GGC'].loc[(~df['ID Number'].isin(users)) & (df['GGC'].isnull())] = 'No Account'
    # wrap up and produce final files
    produce_files(df)


sd = pd.read_excel('W:/Staff Downloads/2020-07 - Staff Download.xlsx')
sd_list = sd['Pay_Number'].tolist()

ggc_oos, nes_oos = getHSEScopeFile()

ggc = GGCScope(sd)
ggc_excl = ggc['Pay_Number'].tolist()
ggc_oos_excl = []
nes = NESScope(ggc)
nes_excl = nes['Pay_Number'].tolist()
nes_oos_excl = []
for i in sd_list:
    if i not in ggc_excl:
        ggc_oos_excl.append(i)
    if i not in nes_excl:
        nes_oos_excl.append(i)
print(f'Out of scope from GGC exclusions: {len(ggc_oos_excl)}, Out of scope from Scopes File: {len(ggc_oos)}')
print(f'Out of scope from NES exclusions: {len(nes_oos_excl)}, Out of scope from Scopes file: {len(nes_oos)}')

print(f'GGC from exclusions: {len(ggc)}, GGC from Scopes file: {40800 - len(ggc_oos)}')
print(f'NES from exclusions: {len(nes)}, NES from Scopes file: {40800 - len(nes_oos)}')

ggc_diff = set(ggc_oos).difference(ggc_oos_excl)
print(len(ggc_diff))
nes_diff = set(nes_oos).difference(nes_oos_excl)
print(len(nes_diff))
print(nes_diff)
ggc_exclusions = pd.DataFrame(list(ggc_diff), columns=['Pay_Number'])
nes_exclusions = pd.DataFrame(list(nes_diff), columns=['Pay_Number'])

ggc_exclusions = ggc_exclusions.merge(sd, on='Pay_Number', how='left')
nes_exclusions = nes_exclusions.merge(sd, on='Pay_Number', how='left')
print(ggc_exclusions.columns)
ggc_exclusions = ggc_exclusions[['Pay_Number', 'Area', 'Sector/Directorate/HSCP_Code',
                                 'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2',
                                 'department', 'Cost_Centre', 'Surname', 'Forename', 'Base',
                                 'Job_Family_Code', 'Job_Family', 'Sub_Job_Family', 'Pay_Band']]
nes_exclusions = nes_exclusions[['Pay_Number', 'Area', 'Sector/Directorate/HSCP_Code',
                                 'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2',
                                 'department', 'Cost_Centre', 'Surname', 'Forename', 'Base',
                                 'Job_Family_Code', 'Job_Family', 'Sub_Job_Family', 'Pay_Band']]

with pd.ExcelWriter('C:/Learnpro_Extracts/sharps/differences.xlsx') as writer:
    ggc_exclusions.to_excel(writer, sheet_name='GGC', index=False)
    nes_exclusions.to_excel(writer, sheet_name='NES', index=False)
writer.save()
# exit()


master_data, user_list = take_in_dir(sharps_courses)
dates_frame = build_user_compliance_dates(master_data)
dates_frame.to_csv('C:/Learnpro_Extracts/sharps/date_debug.csv', index=False)
check_compliance(dates_frame, user_list)
print(dates_frame.dtypes)

# print(sd['NES Module'].value_counts())
