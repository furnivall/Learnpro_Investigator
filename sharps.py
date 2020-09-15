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

starttime = pd.Timestamp.now()
learnpro_runtime = input("when was this data pulled from learnpro? (format = dd-mm-yy)")
learnpro_date = pd.to_datetime(learnpro_runtime, format='%d-%m-%y')
sharps_courses = ['GGC: Management of Needlestick & Similar Injuries',
                  # 'GGC: Prevention & Management of Occ. Exposure Pt 2',
                  'NES: Prev. and Mgmt. of Occ. Exposure (Assessment)']

username = 'xggc\\' + input("GGC username?")
password = input("GGC password")


def NotatWork(date):
    date = date.strftime('%Y-%m-%d')
    df = pd.read_excel('W:/Daily_Absence/' + date + '.xls', skiprows=4)
    print(df.columns)
    print(f'Total absences: {len(df)}')
    mat = df[df['AbsenceReason Description'] == 'Maternity Leave']['Pay No'].tolist()
    df = df[~df['Pay No'].isin(mat)]
    susp = df[df['Absence Type'] == 'Suspended']['Pay No'].tolist()
    df = df[~df['Pay No'].isin(susp)]
    secondment = df[df['AbsenceReason Description'] == 'Secondment']['Pay No'].tolist()
    df = df[~df['Pay No'].isin(secondment)]

    df['Absence Episode Start Date'] = pd.to_datetime(df['Absence Episode Start Date'], format='%Y-%m-%d')

    day28 = pd.Timestamp.now() - pd.DateOffset(days=28)
    long_abs = df[df['Absence Episode Start Date'] < day28]['Pay No'].tolist()
    return mat, susp, secondment, long_abs


mat, susp, secondment, long_abs = NotatWork(learnpro_date - pd.DateOffset(days=1))


def getHSEScopeFile():
    response = requests.get(
        'http://teams.staffnet.ggc.scot.nhs.uk/teams/CorpSvc/HR/WorkInfo/HSE%20Resources/Staff%20Scope%20List.xlsx',
        auth=HttpNtlmAuth(username, password))

    with open('W:/LearnPro/HSEScope-current.xlsx', 'wb') as f:
        f.write(response.content)

    curr_scope = pd.read_excel("W:/LearnPro/HSEScope-current.xlsx")
    ggc_oos = curr_scope[curr_scope['GGC Module'] == "Out Of Scope"]
    nes_oos = curr_scope[curr_scope['NES Module'] == 'Out Of Scope']

    return ggc_oos['Pay_Number'].tolist(), nes_oos['Pay_Number'].tolist()


def check_users():
    df = pd.read_excel()


def empower(date):
    """ This function takes in a date and reduces the empower extract to only those within the 2 year expiry period"""
    df = pd.read_excel('C:/Learnpro_Extracts/Empower_data/final_sharps.xlsx')

    df['Assessment Date'] = pd.to_datetime(df['Assessment Date'], format='%d-%b-%y 00:00:00')
    df = df[df['Assessment Date'] > pd.to_datetime(date) - pd.DateOffset(years=2)]
    return df


def eESS(file, date):
    df = pd.read_csv(file, sep='\t', encoding='utf-16')

    eess_courses = ['GGC E&F Sharps - Disposal of Sharps (Toolbox Talks)',
                    'GGC E&F Sharps - Inappropriate Disposal of Sharps',
                    'GGC E&F Sharps - Management of Injuries (Toolbox Talks)']
    df = df[df['Course Name'].isin(eess_courses)]
    with open("C:/Learnpro_Extracts/eesslookup.txt", "r") as file:
        data = file.read()
    eesslookup = eval(data)
    df['Pay Number'] = df['Employee Number'].map(eesslookup)
    df['Course'] = 'GGC: Management of Needlestick & Similar Injuries'
    df = df.rename(
        columns={'Pay Number': 'ID Number', 'Course': 'Module', 'Course End Date': 'Assessment Date'})

    df = df[['ID Number', 'Module', 'Assessment Date']]
    print(len(df['Assessment Date']))

    df['Assessment Date'] = pd.to_datetime(df['Assessment Date'], format='%Y/%m/%d')
    df = df[df['Assessment Date'] > pd.to_datetime(date) - pd.DateOffset(years=2)]
    df['GGC Source'] = 'eESS'
    df['Assessment Date'] = df['Assessment Date'].dt.strftime('%d/%m/%y %H:%M')
    return df


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

    # Megan Fileman - request from Linda McCall - remedied on 01-09-2020
    df = df[~(df['Pay_Number'] == 'G9838100')]

    print(f'NES Inscope: {len(df)}')
    df.to_excel('C:/Learnpro_Extracts/sharps/nestest.xlsx')
    return df


def GGCScope(df):
    print('Generating GGC Scope')

    df_orig = df

    # In response to query from Arwel Williams (28/08/20), removed certain training grades neuro docs
    arwel_neuro = ['G9854092', 'G9843249', 'G9853174', 'G9854035', 'G9853491', 'G9853144', 'G9853581', 'G9853542']
    df = df[~df['Pay_Number'].isin(arwel_neuro)]

    # In response to query from Marisa McAllister (14/09/20), removing Michele Barrett - quality manager - non-clinical
    df = df[~(df['Pay_Number'] == 'G9830586')]

    # In response to Natalie Mcmillan / Margaret Anderson retirement notice - 01/09/2020
    df = df[~(df['Pay_Number'] == 'G0009520')]

    # In response to Amanda Parker - estates + facilities sharps - 15/09/2020
    amanda_parker = ['G9857762', 'G9858380', 'C6508413', 'C3007189']
    df = df[~(df['Pay_Number'].isin(amanda_parker))]


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

            lp_count += 1
            # read file into df
            df = pd.read_csv(dirname + "/" + file, skiprows=14, sep="\t")
            print("Initial length: " + str(len(df)))
            modules.append(df['Module'].unique().tolist())
            # drop fails
            df = df[df['Passed'] == "Yes"]
            print("Removed fails - new length: " + str(len(df)))
            df = df[['ID Number', 'Module', 'Assessment Date']]
            df['GGC Source'] = 'Learnpro'
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

        elif 'eESS' in file:
            eESS_data = eESS(dirname + '/' + file, learnpro_date)
            master = master.append(eESS_data, ignore_index=True)

        else:
            print("not lp")
            non_lp_count += 1

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
    # deal with empower
    empower_data = empower(learnpro_date)
    master = master.append(empower_data, ignore_index=True)

    with open('C:/Learnpro_Extracts/listfile.txt', 'w') as filehandler:
        for listitem in modules:
            filehandler.write('%s\n' % listitem)

    master.sort_values(by='Assessment Date', inplace=True)
    master.to_csv('C:/Learnpro_Extracts/bigfile.csv', index=False)
    # exit()
    return master, user_ids


def build_user_compliance_dates(df):
    """Takes in a dataframe with lots of learnpro pass data and produces a dataset of test dates as an output"""
    users = df[['ID Number']].drop_duplicates(subset='ID Number').sort_values(by='ID Number')
    initial_length = len(users)
    for module in sharps_courses:
        print(module)
        df1 = df[df['Module'] == module]
        df1 = df1.drop(columns="Module")
        df1 = df1.rename(columns={"Assessment Date": module + " Date"})
        print(f"This module has {len(df1)} users")
        users = users.merge(df1, on="ID Number", how="left")
        users = users[users['ID Number'].notnull()]
        # this is necessary because otherwise it'll create dupes for some reason
        users.drop_duplicates(subset='ID Number', inplace=True, keep='last')

    ggc_source = df[['ID Number', 'GGC Source']].drop_duplicates(subset='ID Number', keep='last')
    users = users.merge(ggc_source, on='ID Number', how='left')

    print("Initial length: " + str(initial_length))
    print("Final length: " + str(len(users['ID Number'].drop_duplicates())))

    # for debug
    users.to_excel("C:/Learnpro_Extracts/sharps/dates.xlsx", index=False)

    return users


def sd_merge(df):
    """This function merges in the Staff Download data to let us work with identifiable stuff for our pivot"""

    # legacy - good for excel pivots
    df['Headcount'] = 1

    # TODO point this somewhere else
    sd = pd.read_excel('W:/Staff Downloads/2020-08 - Staff Download.xlsx')

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

    df = df.rename(columns={'First': 'Forename', 'Last': 'Surname', 'GGC': 'GGC Module', 'NES': 'NES Module',
                            'ID Number': 'Pay_Number'})

    df = df[['Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1',
             'Sub-Directorate 2', 'department', 'Cost_Centre', 'Pay_Number',
             'Surname', 'Forename', 'Base', 'Job_Family', 'Sub_Job_Family',
             'Post_Descriptor', 'WTE', 'Contract_Description', 'NI_Number',
             'Date_Started', 'Job_Description', 'Pay_Band', 'GGC Module', 'GGC Date',
             'GGC Source', 'NES Module', 'NES Date', 'Compliant']]

    ggc_piv = pd.pivot_table(df, index=['Area', 'Sector/Directorate/HSCP'], columns='GGC Module', values='Pay_Number',
                             aggfunc='count', margins=True,fill_value=0, margins_name='All Staff')


    ggc_piv['Not at work ≥ 28 days'] = ggc_piv['Maternity Leave'] + ggc_piv['Suspended'] + ggc_piv['Secondment'] + \
                             ggc_piv['≥28 days Absence']

    ggc_piv['In scope'] = ggc_piv['Complete'] + ggc_piv['Expired'] + ggc_piv['No Account'] + ggc_piv['Not undertaken'] + \
                          ggc_piv['Not at work ≥ 28 days']
    ggc_piv['Compliance %'] = ((ggc_piv['Complete'] + ggc_piv['Not at work ≥ 28 days']) / ggc_piv['In scope'] * 100).round(2)
    ggc_piv.drop(inplace=True, columns=['All Staff', 'Secondment', 'Suspended', 'Out Of Scope', 'Maternity Leave',
                                        '≥28 days Absence'])
    ggc_piv = ggc_piv[ggc_piv['Compliance %'] > 0.01]

    nes_piv = pd.pivot_table(df, index=['Area', 'Sector/Directorate/HSCP'], columns='NES Module', values='Pay_Number',
                             aggfunc='count', margins=True, margins_name='All Staff', fill_value=0)


    nes_piv['Not at work ≥ 28 days'] = nes_piv['Maternity Leave'] + nes_piv['Suspended'] + nes_piv['Secondment'] +\
                             nes_piv['≥28 days Absence']

    nes_piv['In scope'] = (
                nes_piv['Complete'] + nes_piv['Expired'] + nes_piv['No Account'] + nes_piv['Not undertaken'] +
                nes_piv['Not at work ≥ 28 days']).round(2)

    nes_piv['Compliance %'] = ((nes_piv['Complete'] + nes_piv['Not at work ≥ 28 days']) / nes_piv['In scope'] * 100).round(2)
    nes_piv.drop(inplace=True, columns=['All Staff', 'Secondment', 'Suspended', 'Out Of Scope', 'Maternity Leave',
                                        '≥28 days Absence'])
    nes_piv = nes_piv[nes_piv['Compliance %'] > 0.01]


    # These lines below hide the type of absence for privacy reasons. If you want to debug, then comment these out.
    df.loc[((df['NES Module'].isin(['Secondment', 'Out of Scope', 'Maternity Leave', 'Suspended']))),
           'NES Module'] = 'Not at work ≥ 28 days'
    df.loc[((df['GGC Module'].isin(['Secondment', 'Out of Scope', 'Maternity Leave', 'Suspended']))),
           'GGC Module'] = 'Not at work ≥ 28 days'

    # write to book
    with pd.ExcelWriter('W:/Learnpro/HSE Sharps and Skins/'+learnpro_date.strftime('%Y%m%d')+' - HSE Sharps.xlsx') as writer:
        df.to_excel(writer, sheet_name='Export', index=False)
        # TODO add pivot
        # piv.to_excel(writer, sheet_name='pivot')
        ggc_piv.to_excel(writer, sheet_name='GGC Pivot')
        nes_piv.to_excel(writer, sheet_name='NES Pivot')
    writer.save()

    # HSE Named Lists sharepoint upload:
    df.drop(columns=['Pay_Number', 'WTE', 'Contract_Description', 'NI_Number', 'Date_Started', 'Job_Description', 'Post_Descriptor',
                     'Pay_Band'], inplace=True)
    with pd.ExcelWriter('W:/Learnpro/Named Lists HSE/'+learnpro_date.strftime('%Y-%m-%d')+'/'+learnpro_date.strftime('%Y%m%d')+' - HSE Sharps.xlsx') as writer:
        df.to_excel(writer, sheet_name='Export', index=False)
        # TODO add pivot
        # piv.to_excel(writer, sheet_name='pivot')
        ggc_piv.to_excel(writer, sheet_name='GGC Pivot')
        nes_piv.to_excel(writer, sheet_name='NES Pivot')
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
        earliestCompliance = learnpro_date - pd.DateOffset(years=2)
        # df[short_name] = np.where(df[module + ' Date'] > earliestCompliance, "Complete", "Not Compliant")
        df[short_name].loc[df[module + ' Date'] > earliestCompliance] = 'Complete'
        df[short_name].loc[(df[module + ' Date'].notnull()) & (df[module + ' Date'] < earliestCompliance)] = 'Expired'

        df[str(short_name) + ' expires on...'] = df[module + ' Date'] + pd.DateOffset(years=2)
        df[short_name + ' expires on...'].loc[df[short_name] == 'Not Compliant'] = None

    # merge staff download cols into dataset
    df = sd_merge(df)

    # fix for people with more than one pay number

    df = ni_fix(df)

    # df['NES - Scope'] = ""
    # df['GGC - Scope'] = ""
    # #UNCOMMENT IF YOU WANT TO USE RULES
    # df['NES'].loc[(df['NES'].isnull()) & (~df['ID Number'].isin(nes['Pay_Number'].unique().tolist()))] = 'Out Of Scope'
    # df['GGC'].loc[(df['GGC'].isnull()) & (~df['ID Number'].isin(ggc['Pay_Number'].unique().tolist()))] = 'Out Of Scope'

    # UNCOMMENT IF YOU WANT TO USE GGC SCOPE
    df['NES'].loc[(df['ID Number'].isin(nes_oos) | df['ID Number'].isin(nes_oos_excl))] = 'Out Of Scope'
    df['GGC'].loc[(df['ID Number'].isin(ggc_oos) | df['ID Number'].isin(ggc_oos_excl))] = 'Out Of Scope'

    df['NES'].loc[(~df['ID Number'].isin(users)) & (df['NES'].isnull())] = 'No Account'
    df['GGC'].loc[(~df['ID Number'].isin(users)) & (df['GGC'].isnull())] = 'No Account'
    df['GGC'].loc[(df['GGC'].isnull()) | (df['GGC'] == "")] = 'Not undertaken'
    df['NES'].loc[(df['NES'].isnull()) | (df['NES'] == "")] = 'Not undertaken'


    # deal with not at work staff
    df.loc[((df['ID Number'].isin(long_abs)) & (
        ~df['GGC'].isin(['Complete', 'Out Of Scope']))), 'GGC'] = '≥28 days Absence'
    df.loc[((df['ID Number'].isin(mat)) & (
        ~df['GGC'].isin(['Complete', 'Out Of Scope']))), 'GGC'] = 'Maternity Leave'
    df.loc[((df['ID Number'].isin(secondment)) & (
        ~df['GGC'].isin(['Complete', 'Out Of Scope']))), 'GGC'] = 'Secondment'
    df.loc[((df['ID Number'].isin(susp)) & (
        ~df['GGC'].isin(['Complete', 'Out Of Scope']))), 'GGC'] = 'Suspended'
    df.loc[((df['ID Number'].isin(long_abs)) & (
        ~df['NES'].isin(['Complete', 'Out Of Scope']))), 'NES'] = '≥28 days Absence'
    df.loc[((df['ID Number'].isin(mat)) & (
        ~df['NES'].isin(['Complete', 'Out Of Scope']))), 'NES'] = 'Maternity Leave'
    df.loc[((df['ID Number'].isin(secondment)) & (
        ~df['NES'].isin(['Complete', 'Out Of Scope']))), 'NES'] = 'Secondment'
    df.loc[((df['ID Number'].isin(susp)) & (
        ~df['NES'].isin(['Complete', 'Out Of Scope']))), 'NES'] = 'Suspended'

    df['Compliant'] = ''
    df['Compliant'].loc[(df['GGC'] == 'Complete') & ((df['NES'].isin(['Out Of Scope', 'Complete'])))] = 1
    df['Compliant'].loc[(df['Compliant'] == '')] = 0

    # wrap up and produce final files
    produce_files(df)


sd = pd.read_excel('W:/Staff Downloads/2020-08 - Staff Download.xlsx')
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
print(f'nes from exclusions {nes_oos_excl[0:10]}')
print(f'nes from file{nes_oos[0:10]}')

print(f'GGC from exclusions: {len(ggc)}, GGC from Scopes file: {40800 - len(ggc_oos)}')
print(f'NES from exclusions: {len(nes)}, NES from Scopes file: {40800 - len(nes_oos)}')

ggc_diff = set(ggc_oos).difference(ggc_oos_excl)
print(f'GGC Diff = {len(ggc_diff)}')
nes_diff = set(nes_oos).difference(nes_oos_excl)
print(f'NES Diff = {len(nes_diff)}')

ggc_exclusions = pd.DataFrame(list(ggc_diff), columns=['Pay_Number'])
nes_exclusions = pd.DataFrame(list(nes_diff), columns=['Pay_Number'])

ggc_exclusions = ggc_exclusions.merge(sd, on='Pay_Number', how='left')
nes_exclusions = nes_exclusions.merge(sd, on='Pay_Number', how='left')

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

# print(sd['NES Module'].value_counts())
endtime = pd.Timestamp.now()
print(f'Total time elapsed : {endtime - starttime}')
