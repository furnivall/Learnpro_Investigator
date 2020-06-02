"""
This is the new master learnpro file.
It takes a directory of learnpro-related files (including eESS and empower files), concatenates them into a
single dataset, then crunches them into a full compliance named list for statutory/mandatory training compliance.
"""
import pandas as pd
from tkinter.filedialog import askdirectory
import os
import datetime as dt
import numpy as np

# these are the modules we're looking at specifically to capture compliance
stat_mand = ['GGC: 001 Fire Safety',
             'GGC: Health and Safety, an Introduction',
             'GGC: 003 Reducing Risks of Violence & Aggression',
             'GGC: Equality, Diversity and Human Rights',
             'GGC: Manual Handling Theory',
             'Child Protection - Level 1',
             'Adult Support & Protection',
             'GGC: Standard Infection Control Precautions ',
             'GGC: 008 Security & Threat',
             'GGC: 009 Safe Information Handling',
             'Safe Information Handling',
             'Fire Emergency within the Ward', 'Fire Fighting Equipment', 'Fire Prevention',
              'Introduction and General Fire Safety', 'Specialist Roles'
             ]


def sd_pull():
    """Read in staff download, then chop it down to ID and NI number."""
    # TODO replace this direct link with a prompt later - using this for speed currently
    df = pd.read_excel('W:/Staff Downloads/2020-04 - Staff Download.xlsx')
    df = df[['Pay_Number', 'NI_Number']]
    # rename column for easy merging
    df.columns = ['ID Number', 'NI Number']
    return df


def users_file():
    """Read in learnpro users file, then chop it to relevant identifiables"""
    # TODO replace to find specific users file.
    df = pd.read_csv('C:/Learnpro_Extracts/lpdata-28-05-20/USERS.xls', skiprows=11, sep="\t")
    df = df[['First Name', 'Last Name', 'Email / Username', 'ID Number', 'Registration Date']]
    return df


def take_in_dir(list_of_modules):
    """Takes in directory with prompt then selects all learnpro files within that folder"""

    # counters for lp/nonlp files within dir
    lp_count = 0
    non_lp_count = 0

    # prompt for dir

    # TODO swap this back when testing is done
    # dirname = askdirectory(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data',
    #                        title="Choose a directory full of learnpro files."
    #                        )
    dirname = 'C:/Learnpro_Extracts/LP_testData_15-05-20'

    # initialise master dataframe
    master = pd.DataFrame()
    modules = []

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
        else:
            # TODO deal with empower, eESS etc
            print("not lp")
            non_lp_count += 1
    # log outputs below:
    print(str(lp_count + non_lp_count) + " files read. " + str(lp_count) + " contained learnpro data, " + str(
        non_lp_count) +
          " did not.")
    master['Module'] = master['Module'].astype('category')

    master['Assessment Date'] = pd.to_datetime(master['Assessment Date'], format='%d/%m/%y %H:%M').dt.normalize()

    df_users = master['ID Number'].unique().tolist()
    modules = [item for sublist in modules for item in sublist]
    modules = list(dict.fromkeys(modules))

    with open('C:/Learnpro_Extracts/listfile.txt', 'w') as filehandler:
        for listitem in modules:
            filehandler.write('%s\n' % listitem)


    return master, df_users


def get_temp_accounts(df):
    """Captures all temporary accounts (denoted with a T at the beginning)"""
    df_temps = df[df['ID Number'].str.contains("T", na=False)]

    print(str(len(df_temps)) + " temporary accounts found.")
    return df_temps


def build_user_compliance_dates(df):
    """Takes in a dataframe with lots of learnpro pass data and produces a dataset of test dates as an output"""
    users = df['ID Number'].drop_duplicates().sort_values().to_frame()
    initial_length = len(users)

    for module in stat_mand:
        print(module)
        df1 = df[df['Module'] == module]
        df1 = df1.drop(columns="Module")
        df1 = df1.rename(columns={"Assessment Date": module + " Date"})
        print("This module has " + str(len(df1)) + "users")
        users = users.merge(df1, on="ID Number", how="left")
        users = users[users['ID Number'].notnull()]
        # this is necessary because otherwise it'll create dupes for some reason
        users.drop_duplicates(subset='ID Number', inplace=True)
        print("Current dataset size: " + str(len(users)) + " users")
        print("Dataframe shape: " + str(users.shape))

    # users.to_csv("C:/Learnpro_Extracts/namedList.csv", index=False)
    print("Initial length: " + str(initial_length))
    print("Final length: " + str(len(users['ID Number'].drop_duplicates())))
    return users


def check_compliance(df):
    """Takes in a df with test dates for each of the stat/mand modules, including both versions of safe info handling.
    It will produce a completed named list dataset ala the old CompliancePro"""
    print(df.columns)
    # dealing with multiple types of safe info handling files

    for module in stat_mand:
        columnnames = {'GGC: 001 Fire Safety': 'Fire Awareness',
                       'GGC: Health and Safety, an Introduction': 'Health, Safety & Welfare',
                       'GGC: 003 Reducing Risks of Violence & Aggression': 'Violence and Aggression',
                       'GGC: Equality, Diversity and Human Rights': 'Equality, Diversity and Human Rights',
                       'GGC: Manual Handling Theory': 'Manual Handling',
                       'GGC: Standard Infection Control Precautions ': 'Infection Control',
                       'GGC: 008 Security & Threat': 'Security and Threat',
                       }
        if module == 'GGC: 001 Fire Safety':
            continue


        if module in ['GGC: 009 Safe Information Handling',
                      'Safe Information Handling']:
            df[module + ' Date'] = df[module + ' Date'] + pd.DateOffset(years=3)
            df['Information Governance Date'] = df[
                ['GGC: 009 Safe Information Handling Date', 'Safe Information Handling Date']].max(axis=1)
            df['Information Governance'] = np.where((df['GGC: 009 Safe Information Handling Date'].notnull())
                                                    | (df['Safe Information Handling Date'].notnull()), "Complete",
                                                    "Not Compliant")
            df['Information Governance expires on...'] = df['Information Governance Date'] + pd.DateOffset(years=3)
            continue

        short_name = columnnames.get(module)
        print(df[module + ' Date'][100])


        df[short_name] = np.where(df[module + ' Date'].notnull(), "Complete", "Not Compliant")
        df[str(short_name) + ' expires on...'] = df[module + ' Date'] + pd.DateOffset(years=3)


        print(type(df[module + ' Date'][0]))
        print(module)

    # TODO public protection

    df['Public Protection'] = np.where(df['Adult Support & Protection Date'].notnull() &
                                       df['Child Protection - Level 1 Date'].notnull(), "Complete", "Not Compliant")
    df['Public Protection expires on...'] = (df[['Adult Support & Protection Date',
                                        'Child Protection - Level 1 Date']].min(axis=1)) + pd.DateOffset(years=3)

    df['Public Protection expires on...'].loc[df['Public Protection'] == 'Not Compliant'] = ''

    df['Fire Awareness expires on...'] = df['GGC: 001 Fire Safety Date'] + pd.DateOffset(years=1)
    startdate_fire = pd.Timestamp.now()
    enddate_fire = startdate_fire + pd.DateOffset(years=1)
    date_range = (df['Fire Awareness expires on...'] > startdate_fire) &\
                 (df['Fire Awareness expires on...'] <= enddate_fire)
    df['Fire Awareness'] = np.where(date_range, "Complete", "Not Compliant")

    SM_lookup = {'Fire Awareness':'SM1', 'Health, Safety & Welfare': 'SM2', 'Violence and Aggression':'SM3',
                 'Equality, Diversity and Human Rights':'SM4', 'Manual Handling':'SM5', 'Infection Control':'SM6',
                 'Public Protection':'SM7','Security and Threat':'SM8', 'Information Governance':'SM9'}
    for i in SM_lookup:
        print(SM_lookup.get(i))
        df[SM_lookup.get(i)] = np.where(df[i] == 'Complete', 1, 0)
    df['Headcount'] = 1
    print(df.columns)

    sd = pd.read_excel('W:/Staff Downloads/2020-04 - Staff Download.xlsx')
    sd = sd.rename(columns={'Pay_Number': 'ID Number'})
    sd = sd[['ID Number','Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2', 'department', 'Cost_Centre',
             'Forename', 'Surname', 'Job_Family', 'Sub_Job_Family']]
    print(len(sd))
    print(len(df))

    df = df.merge(sd, on='ID Number', how='right')
    print(len(df))
    df = df.rename(columns={'Forename':'First', 'Surname':'Last'})
    # df = df[['Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1',
    #    'Sub-Directorate 2', 'department', 'Cost_Centre', 'First', 'Last',
    #    'Job_Family', 'Sub_Job_Family', 'Equality, Diversity and Human Rights',
    #    'Fire Awareness', 'Health, Safety & Welfare', 'Infection Control',
    #    'Information Governance', 'Manual Handling', 'Public Protection',
    #    'Security and Threat', 'Violence and Aggression',
    #    'Equality, Diversity and Human Rights expires on...',
    #    'Fire Awareness expires on...',
    #    'Health, Safety & Welfare expires on...',
    #    'Infection Control expires on...',
    #    'Information Governance expires on...', 'Manual Handling expires on...',
    #    'Public Protection expires on...', 'Security and Threat expires on...',
    #    'Violence and Aggression expires on...', 'SM1', 'SM2', 'SM3', 'SM4',
    #    'SM5', 'SM6', 'SM7', 'SM8', 'SM9', 'Headcount']]


    df.to_csv("C:/Learnpro_Extracts/namedList.csv", index=False)
    # TODO calculate fire safety dates
    # TODO add compliance for all other files


master_data, user_number = take_in_dir(stat_mand)
print(type(master_data['Assessment Date'][0]))
dates_frame = build_user_compliance_dates(master_data)
check_compliance(dates_frame)

exit()
users = users_file()
print("Learnpro Users: " + str(len(users)) + " users.")
print(users.columns)

sd = sd_pull()
print("Staff Download: " + str(len(sd)) + " users.")
