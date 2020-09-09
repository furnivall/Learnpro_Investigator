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

learnpro_runtime = input("when was this data pulled from learnpro? (format = dd-mm-yy)")
startdate_fire = pd.to_datetime(learnpro_runtime, format='%d-%m-%y')
starttime = pd.Timestamp.now()
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
             # these are used to capture old fire compliance, but do not seem necessary any more
             'Fire Emergency within the Ward', 'Fire Fighting Equipment', 'Fire Prevention',
             'Introduction and General Fire Safety', 'Specialist Roles',
             # eESS ones
             'GGC E&F StatMand - Equality, Diversity & Human Rights (face to face session)',
             'GGC E&F StatMand - Health & Safety an Induction (face to face session)',
             'GGC E&F StatMand - Manual Handling Theory (face to face session)',
             'Public Protection - eESS',
             'GGC E&F StatMand - Reducing Risks of Violence & Aggression(face to face session)',
             'GGC E&F StatMand - Safe Information Handling Foundation (face to face session)',
             'GGC E&F StatMand - Security & Threat (face to face session)',
             'GGC E&F StatMand - Standard Infection Control Precautions (face to face session)',
             'GGC E&F StatMand - General Awareness Fire Safety Training (face to face session)'
             ]


def sd_pull():  # UNUSED CURRENTLY
    """Read in staff download, then chop it down to ID and NI number."""
    # TODO replace this direct link with a prompt later - using this for speed currently
    df = pd.read_excel('W:/Staff Downloads/2020-08 - Staff Download.xlsx')
    df = df[['Pay_Number', 'NI_Number']]
    # rename column for easy merging
    df.columns = ['ID Number', 'NI Number']
    return df


def users_file():  # UNUSED CURRENTLY
    """Read in learnpro users file, then chop it to relevant identifiables"""
    # TODO replace to find specific users file.
    df = pd.read_csv('C:/Learnpro_Extracts/lpdata-28-05-20/USERS.xls', skiprows=11, sep="\t")
    df = df[['First Name', 'Last Name', 'Email / Username', 'ID Number', 'Registration Date']]
    return df


def empower(date):
    """ This function takes in a date and reduces the empower extract to only those within the 3 year expiry period"""
    df = pd.read_excel('C:/Learnpro_Extracts/LP_testData_31-05-20/empower_final.xlsx')
    print(df['Assessment Date'].iloc[0])
    df = df[df['Assessment Date'] > pd.to_datetime(date) - pd.DateOffset(years=3)]
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

    dirname = 'C:/Learnpro_Extracts/20200904-auto'

    # initialise master dataframe
    master = pd.DataFrame()
    modules = []
    # read in historic empower data

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
            df['Module'].loc[df['Module'] == 'Adult Support and Protection Act'] = 'Adult Support and Protection'
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

    master['Assessment Date'] = pd.to_datetime(master['Assessment Date'], format='%d/%m/%y %H:%M')
    eess = eESS('C:/Learnpro_Extracts/20200904-auto/CompliancePro_Extract.csv')
    master = master.append(eess, ignore_index=True)

    # deal with empower
    empower_data = empower(startdate_fire)
    master = master.append(empower_data)

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
    return master, df_users


def eESS(file):
    df = pd.read_csv(file, sep='\t', encoding='utf-16')
    eess_courses = ['GGC E&F StatMand - Equality, Diversity & Human Rights (face to face session)',
                    'GGC E&F StatMand - General Awareness Fire Safety Training (face to face session)',
                    'GGC E&F StatMand - Health & Safety an Induction (face to face session)',
                    'GGC E&F StatMand - Manual Handling Theory (face to face session)',
                    'GGC E&F StatMand - Public Protection (face to face session)',
                    'GGC E&F StatMand - Reducing Risks of Violence & Aggression(face to face session)',
                    'GGC E&F StatMand - Safe Information Handling Foundation (face to face session)',
                    'GGC E&F StatMand - Security & Threat (face to face session)',
                    'GGC E&F StatMand - Standard Infection Control Precautions (face to face session)']
    df = df[df['Course Name'].isin(eess_courses)]
    with open("C:/Learnpro_Extracts/eesslookup.txt", "r") as file:
        data = file.read()
    eesslookup = eval(data)
    df['Pay Number'] = df['Employee Number'].map(eesslookup)
    df['Course'] = df['Course Name'].map(
        {
            'GGC E&F StatMand - Equality, Diversity & Human Rights (face to face session)': 'GGC: Equality, Diversity and Human Rights',
            'GGC E&F StatMand - General Awareness Fire Safety Training (face to face session)': 'GGC: 001 Fire Safety',
            'GGC E&F StatMand - Health & Safety an Induction (face to face session)': 'GGC: Health and Safety, an Introduction',
            'GGC E&F StatMand - Manual Handling Theory (face to face session)': 'GGC: Manual Handling Theory',
            'GGC E&F StatMand - Public Protection (face to face session)':'Public Protection - eESS',
            'GGC E&F StatMand - Reducing Risks of Violence & Aggression(face to face session)': 'GGC: 003 Reducing Risks of Violence & Aggression',
            'GGC E&F StatMand - Safe Information Handling Foundation (face to face session)': 'GGC: 009 Safe Information Handling',
            'GGC E&F StatMand - Security & Threat (face to face session)': 'GGC: 008 Security & Threat',
            'GGC E&F StatMand - Standard Infection Control Precautions (face to face session)': 'GGC: Standard Infection Control Precautions '})

    df = df.rename(
        columns={'Pay Number': 'ID Number', 'Course': 'Module', 'Course End Date': 'Assessment Date'})

    df = df[['ID Number', 'Module', 'Assessment Date']]
    print(len(df['Assessment Date']))

    df['Assessment Date'] = pd.to_datetime(df['Assessment Date'], format='%Y/%m/%d')

    return df


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
        users.drop_duplicates(subset='ID Number', inplace=True, keep='last')
        print("Current dataset size: " + str(len(users)) + " users")
        print("Dataframe shape: " + str(users.shape))
    # for debug
    # users.to_excel("C:/Learnpro_Extracts/dates.xlsx", index=False)

    print("Initial length: " + str(initial_length))
    print("Final length: " + str(len(users['ID Number'].drop_duplicates())))
    return users


def ni_fix(df):
    """This little function does a lot - it gets all NI numbers with duplicates, then creates a dataframe for each dupe,
    checks the compliance for each course across all relevant pay numbers, then gets the most recent course completion
    date then applies these to all of the pay numbers"""

    # find dupes and put in list
    dups = df[df.duplicated(subset='NI_Number')]['NI_Number'].drop_duplicates().to_list()

    # this is to make df.update() be able to compare indices and amend data
    df.set_index('ID Number')

    courses = ['Equality, Diversity and Human Rights', 'Fire Awareness', 'Health, Safety & Welfare',
               'Infection Control',
               'Information Governance', 'Manual Handling', 'Public Protection', 'Security and Threat',
               'Violence and Aggression']

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


def SM_num_maker(df):
    """This function adds the SM1-9 columns, which exist in the old named list format produced by CompliancePro"""

    SM_lookup = {'Fire Awareness': 'SM2', 'Health, Safety & Welfare': 'SM3', 'Violence and Aggression': 'SM9',
                 'Equality, Diversity and Human Rights': 'SM1', 'Manual Handling': 'SM6', 'Infection Control': 'SM4',
                 'Public Protection': 'SM7', 'Security and Threat': 'SM8', 'Information Governance': 'SM5'}

    for i in SM_lookup:
        df[SM_lookup.get(i)] = np.where(df[i] == 'Complete', 1, 0)

    return df


def sd_merge(df):
    """This function merges in the Staff Download data to let us work with identifiable stuff for our pivot"""

    # legacy - good for excel pivots


    # TODO point this somewhere else
    sd = pd.read_excel('W:/Staff Downloads/2020-08 - Staff Download.xlsx')

    # Cleaning step - vital for merge to work properly - ID number must be on both sides of merge
    sd = sd.rename(columns={'Pay_Number': 'ID Number', 'Forename': 'First', 'Surname': 'Last'})

    # cut down to useful cols
    sd = sd[['ID Number', 'NI_Number', 'Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2',
             'department', 'Cost_Centre', 'First', 'Last', 'Job_Family', 'Sub_Job_Family']]

    # merge
    df = df.merge(sd, on='ID Number', how='right')

    return df


def produce_files(df):
    """Builds final files for named list"""
    df['Headcount'] = 1
    # subset for final named list
    df2 = df[['ID Number', 'Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2', 'department',
              'Cost_Centre', 'First', 'Last', 'Job_Family', 'Sub_Job_Family', 'Equality, Diversity and Human Rights',
              'Fire Awareness', 'Health, Safety & Welfare', 'Infection Control', 'Information Governance',
              'Manual Handling', 'Public Protection', 'Security and Threat', 'Violence and Aggression',
              'Equality, Diversity and Human Rights expires on...', 'Fire Awareness expires on...',
              'Health, Safety & Welfare expires on...', 'Infection Control expires on...',
              'Information Governance expires on...', 'Manual Handling expires on...',
              'Public Protection expires on...',
              'Security and Threat expires on...', 'Violence and Aggression expires on...', 'SM1', 'SM2', 'SM3', 'SM4',
              'SM5', 'SM6', 'SM7', 'SM8', 'SM9', 'Headcount']]

    # build pivot tab for comparison
    piv = pd.pivot_table(df, index='Sector/Directorate/HSCP',
                         values=['SM1', 'SM2', 'SM3', 'SM4', 'SM5', 'SM6', 'SM7', 'SM8', 'SM9'],
                         aggfunc=np.sum, margins=True)
    # produce compliance percentage row
    piv.loc['% Compliance'] = round(piv.iloc[-1] / len(df) * 100, 1)

    # write to book
    with pd.ExcelWriter('C:/Learnpro_Extracts/namedList.xlsx') as writer:
        df2.to_excel(writer, sheet_name='data', index=False)
        piv.to_excel(writer, sheet_name='pivot')

    # for debugging - make csv file with all columns and data
    df.to_csv("C:/Learnpro_Extracts/namedList.csv")


def check_compliance(df):
    """Takes in a df with test dates for each of the stat/mand modules, including both versions of safe info handling.
    It will produce a completed named list dataset ala the old CompliancePro"""

    print(df.columns)

    # dealing with multiple types of fire safety files - not actually relevant any more but doesn't slow anything down
    fire_mods = ['GGC: 001 Fire Safety',
                 'GGC E&F StatMand - General Awareness Fire Safety Training (face to face session)',
                 'Fire Emergency within the Ward', 'Fire Fighting Equipment', 'Fire Prevention',
                 'Introduction and General Fire Safety', 'Specialist Roles']

    # iterate through modules, grabbing their test dates as appropriate
    for module in stat_mand:
        columnnames = {'GGC: 001 Fire Safety': 'Fire Awareness',
                       'GGC: Health and Safety, an Introduction': 'Health, Safety & Welfare',
                       'GGC: 003 Reducing Risks of Violence & Aggression': 'Violence and Aggression',
                       'GGC: Equality, Diversity and Human Rights': 'Equality, Diversity and Human Rights',
                       'GGC: Manual Handling Theory': 'Manual Handling',
                       'GGC: Standard Infection Control Precautions ': 'Infection Control',
                       'GGC: 008 Security & Threat': 'Security and Threat',
                       }
        if module in fire_mods:
            # add 1 year to fire test date to get expiry, then go to start of loop again
            df[module + ' Date'] = df[module + ' Date'] + pd.DateOffset(years=1)
            continue

        # deal with info gov
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

        # the courses all have a shorter name to look nicer (see above dictionary)
        short_name = columnnames.get(module)

        # attach compliance to all except infogov and fire. It is likely this will need to be changed to capture more
        df[short_name] = np.where(df[module + ' Date'].notnull(), "Complete", "Not Compliant")
        df[str(short_name) + ' expires on...'] = df[module + ' Date'] + pd.DateOffset(years=3)

    # pub prot
    df['Public Protection'] = np.where(df['Adult Support & Protection Date'].notnull() &
                                       df['Child Protection - Level 1 Date'].notnull(), "Complete", "Not Compliant")


    # deal with the eESS public prot (covers both adult and child)
    pubprot_mask = (df['Public Protection - eESS Date'].notnull())
    df.loc[pubprot_mask, 'Public Protection'] = 'Complete'

    fire_mods_dates = [i + ' Date' for i in fire_mods[1:]]

    df['Public Protection expires on...'] = (df[['Adult Support & Protection Date',
                                                 'Child Protection - Level 1 Date']].min(axis=1)) + pd.DateOffset(
        years=3)

    # give eEss public prot people an expiry date
    df.loc[pubprot_mask, 'Public Protection expires on...'] = df['Public Protection - eESS Date'] + pd.DateOffset(
        years=3)


    df['Public Protection expires on...'].loc[df['Public Protection'] == 'Not Compliant'] = None

    # get most recent fire pass from all modules
    df['Fire Awareness expires on...'] = np.where(df['GGC: 001 Fire Safety Date'].notnull(),
                                                  df['GGC: 001 Fire Safety Date'], df[fire_mods_dates].min(axis=1))


    # TODO change this to max import date

    enddate_fire = startdate_fire + pd.DateOffset(years=1)
    date_range = (df['Fire Awareness expires on...'] > startdate_fire) & \
                 (df['Fire Awareness expires on...'] <= enddate_fire)
    df['Fire Awareness'] = np.where(date_range, "Complete", "Not Compliant")

    # merge staff download cols into dataset
    df = sd_merge(df)

    # fix for people with more than one pay number
    df = ni_fix(df)

    # add cols for SM1-9 to keep legacy shape ala compliancepro
    df = SM_num_maker(df)

    # wrap up and produce final files
    produce_files(df)


master_data, user_number = take_in_dir(stat_mand)
print(type(master_data['Assessment Date'][0]))
dates_frame = build_user_compliance_dates(master_data)
check_compliance(dates_frame)

endtime = pd.Timestamp.now()
print(f'Runtime = {endtime - starttime}')