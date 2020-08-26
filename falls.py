import pandas as pd
import requests
from requests_ntlm import HttpNtlmAuth
import os
from tkinter.filedialog import askdirectory

start = pd.Timestamp.now()

pd.set_option('display.width', 420)
pd.set_option('display.max_columns', 10)

falls_courses = ['An Introduction to Falls', 'The Falls Bundle of Care', 'What to do when your patient falls',
                 'Risk Factors for Falls (Part 1)', 'Risk Factors for Falls (Part 2)', 'Falls - Bedrails ']

username = 'xggc\\' + input("GGC username?")
password = input("GGC password")




def getHSEScopeFile():
    response = requests.get(
        'http://teams.staffnet.ggc.scot.nhs.uk/teams/CorpSvc/HR/WorkInfo/HSE%20Resources/HSE%20Scope%20List.xlsx',
        auth=HttpNtlmAuth(username, password))

    with open('W:/LearnPro/HSEScope-current.xlsx', 'wb') as f:
        f.write(response.content)

    curr_scope = pd.read_excel("W:/LearnPro/HSEScope-current.xlsx", skiprows=1)
    print(curr_scope.columns)
    curr_scope = curr_scope[['Cost Centre', 'Administrative Services', 'Allied Health Profession',
       'Dental Support', 'Executives', 'Healthcare Sciences',
       'Medical And Dental', 'Medical Support', 'Nursing and Midwifery',
       'Other Therapeutic', 'Personal and Social Care', 'Support Services']]
    curr_scope.rename(columns={'Medical And Dental':'Medical and Dental'}, inplace=True)
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
        count=0
        for j in i[1:]:
            count = count+1
            if isinstance(j, str):
                data = (i[0] + fams[count-1])
                output.append(data)
            else:
               continue

    return output

def merge_scopes(concat):
    sd = pd.read_excel('W:/Staff Downloads/2020-07 - Staff Download.xlsx')
    print(sd.columns)
    sd['Concat'] = sd['Cost_Centre'] + sd['Job_Family']
    sd['Falls Compliant'] = ""
    sd['Falls Compliant'].loc[~sd['Concat'].isin(concat)] = 'Out of scope'
    print(sd['Falls Compliant'].value_counts())
    return sd

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

    with open('C:/Learnpro_Extracts/falls/listfile.txt', 'w') as filehandler:
        for listitem in modules:
            filehandler.write('%s\n' % listitem)

    print(master.columns)
    print(master.dtypes)
    master.to_csv('C:/Learnpro_Extracts/falls/bigfile-falls.csv', index=False)
    # exit()
    return master, user_ids

def build_user_compliance_dates(df):
    """Takes in a dataframe with lots of learnpro pass data and produces a dataset of test dates as an output"""
    users = df['ID Number'].drop_duplicates().sort_values().to_frame()
    initial_length = len(users)
    for module in falls_courses:
        print(module)
        df1 = df[df['Module'] == module]
        df1 = df1.drop(columns="Module")
        df1 = df1.rename(columns={"Assessment Date": module})
        print("This module has " + str(len(df1)) + "users")
        users = users.merge(df1, on="ID Number", how="left")
        users = users[users['ID Number'].notnull()]
        # this is necessary because otherwise it'll create dupes for some reason
        users.drop_duplicates(subset='ID Number', inplace=True, keep='last')
        print("Current dataset size: " + str(len(users)) + " users")
        print("Dataframe shape: " + str(users.shape))
    # for debug
    users.to_excel("C:/Learnpro_Extracts/falls/dates.xlsx", index=False)

    print("Initial length: " + str(initial_length))
    print("Final length: " + str(len(users['ID Number'].drop_duplicates())))
    print(users.columns)
    return users

def produce_files(df):
    """Builds final files for named list"""
    # List of columns in final file from Ben's sheet:

    print(df.columns)

    df = df.rename(columns={'First':'Forename', 'Last':'Surname', 'ID Number':'Pay_Number'})
    df['Compliant'] = 'DO THIS LATER'
    df['GGC Source'] = 'Learnpro'
    df = df[['Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1',
             'Sub-Directorate 2', 'department', 'Cost_Centre', 'Pay_Number',
             'Surname', 'Forename', 'Base', 'Job_Family', 'Sub_Job_Family',
             'Post_Descriptor', 'WTE', 'Contract_Description', 'NI_Number',
             'Date_Started', 'Job_Description', 'Pay_Band','Falls Compliant', 'Falls Date',
             'An Introduction to Falls', 'The Falls Bundle of Care', 'What to do when your patient falls',
             'Risk Factors for Falls (Part 1)', 'Risk Factors for Falls (Part 2)', 'Falls - Bedrails ']]

    # write to book
    with pd.ExcelWriter('C:/Learnpro_Extracts/falls/namedList.xlsx') as writer:
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


    # loop through courses, first making all compliant if necessary, then editing expiry dates
    for number, i in enumerate(dups):
        dfx = df[df['NI_Number'] == i]
        dfx.loc[:, 'Falls Compliant'] = min(dfx['Falls Compliant'].tolist(), key=len)
        for j in falls_courses:
            dfx.loc[:, j] = dfx[j].max()
        df.update(dfx)
    df.reset_index()
    return df

def check_compliance(df, users):
    """Takes in a df with test dates for each of the stat/mand modules, including both versions of safe info handling.
        It will produce a completed named list dataset ala the old CompliancePro"""

    print(df.columns)

    # iterate through modules, grabbing their test dates as appropriate





    # merge staff download cols into dataset
    df = sd_merge(df, sd)
    df['Falls Date'] = df[falls_courses].min(axis=1)
    subset = df[falls_courses]
    df['Falls Date'].loc[subset.isnull().any(axis=1)] = ''
    df['Falls Compliant'].loc[(df['Falls Date'] != '') & (df['Falls Compliant'] == "")] = 'Complete'
    df['Falls Compliant'].loc[(df['Falls Compliant'] =="") & (~df['ID Number'].isin(users))] = 'No Account'
    df['Falls Compliant'].loc[(subset.isnull().all(axis=1)) & (df['Falls Compliant'] == "")] = 'Not Undertaken'
    df['Falls Compliant'].loc[(subset.isnull().any(axis=1)) & (df['Falls Compliant'] == "")] = 'Not Complete'
    # fix for people with more than one pay number

    df = ni_fix(df)

    # df['NES - Scope'] = ""
    # df['GGC - Scope'] = ""
    # df['NES'].loc[(df['NES'].isnull()) & (~df['ID Number'].isin(nes['Pay_Number'].unique().tolist()))] = 'Out Of Scope'
    # df['GGC'].loc[(df['GGC'].isnull()) & (~df['ID Number'].isin(ggc['Pay_Number'].unique().tolist()))] = 'Out Of Scope'
    # df['NES'].loc[(~df['ID Number'].isin(users)) & (df['NES'].isnull())] = 'No Account'
    # df['GGC'].loc[(~df['ID Number'].isin(users)) & (df['GGC'].isnull())] = 'No Account'
    # # wrap up and produce final files
    produce_files(df)

def sd_merge(df, sd):
    """This function merges in the Staff Download data to let us work with identifiable stuff for our pivot"""

    # legacy - good for excel pivots
    df['Headcount'] = 1


    # Cleaning step - vital for merge to work properly - ID number must be on both sides of merge
    sd = sd.rename(columns={'Pay_Number': 'ID Number', 'Forename': 'First', 'Surname': 'Last'})

    # # cut down to useful cols
    # sd = sd[['ID Number', 'NI_Number', 'Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2',
    #          'department', 'Cost_Centre', 'First', 'Last', 'Job_Family', 'Sub_Job_Family', 'Base',
    #          'Contract_Description', 'Job_Description', 'Post_Descriptor', 'WTE', 'Pay_Band',
    #          'Date_Started']]

    # merge
    df = df.merge(sd, on='ID Number', how='right')

    return df

scope = getHSEScopeFile()
concats = workwithScopes(scope)
sd = merge_scopes(concats)
all_learnpro, users = take_in_dir(falls_courses)
dates_frame = build_user_compliance_dates(all_learnpro)
check_compliance(dates_frame, users)

end = pd.Timestamp.now()
print(end - start)
