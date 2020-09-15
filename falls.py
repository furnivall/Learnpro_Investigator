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

learnpro_runtime = input("when was this data pulled from learnpro? (format = dd-mm-yy)")
learnpro_date = pd.to_datetime(learnpro_runtime, format='%d-%m-%y')

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
    sd = pd.read_excel('W:/Staff Downloads/2020-08 - Staff Download.xlsx')
    print(sd.columns)
    sd['Concat'] = sd['Cost_Centre'] + sd['Job_Family']
    sd['Falls Compliant'] = ""
    sd['Falls Compliant'].loc[~sd['Concat'].isin(concat)] = 'Out of scope'

    # In response to query from Pauline Simpson (28/08/2020), we have removed all Obstetrics staff from the scope list
    sd['Falls Compliant'].loc[sd['Sub-Directorate 2'] == 'Obstetrics'] = 'Out of scope'

    # In response to a query from Stephanie Mckay (08/09/2020), we have added "Qeuh-neuro + Omfs Opd" to the
    # scope for Falls
    sd['Falls Compliant'].loc[sd['department'] == 'Qeuh-neuro + Omfs Opd'] = ''

    # in response to Cameron's look of shock, we have removed two cost centres with a single in-scope AHP as they
    # incorrectly added
    sd['Falls Compliant'].loc[(sd['Cost_Centre'].isin(['G51126', 'G69311']))] = 'Out of scope'

    # in response to Craig Broadfoot (14/09/20), a large list of consultants has been removed from scope
    list_of_consultants = ['G9879385', 'G3014304', 'G3012239', 'G2965208', 'G3012247', 'G9463496', 'G950611X',
                           'G9842847', 'G5906911', 'G9134654', 'G0003267', 'G3016420', 'G3033821', 'G9876054',
                           'G9328939', 'G9877304', 'G9884616', 'G9886307', 'G9861715', 'G9557253', 'G5864593',
                           'G4925688', 'G492035X', 'G9878810', 'G2966808', 'G2951797', 'G3009955', 'G297200X',
                           'G9276769', 'G2980533', 'G2988801', 'G9854590', 'G9846003', 'G9219668', 'G9186832',
                           'G0007361', 'G0285242', 'G0006354', 'G0003450', 'G3029395', 'G9871839', 'G3034925',
                           'G301603X', 'G9496572', 'G9878294', 'G9876461', 'G9877777', 'G9536086', 'G0000003',
                           'G9888845', 'G9883687', 'G9882775', 'G9868017', 'G910366X', 'G5828341', 'G3044947',
                           'G9477098', 'G2994143', 'G9844086', 'G9849161', 'G9182640', 'G1081756', 'G3039471',
                           'G3024466', 'G9364730', 'G934795X', 'G9530223', 'G9527850', 'G9530029', 'G9885350',
                           'G9886988', 'G9869609', 'G2997606', 'G2988232', 'G9849287', 'G3040895', 'G9838092',
                           'G9837737', 'G9874536', 'G9830944', 'G9831123', 'G9886234', 'G9867879', 'G9864751',
                           'G9124578', 'G9849270', 'G9831099', 'G3010996', 'G2984032', 'G9178678', 'G9853537',
                           'G3001725', 'G9472290', 'G9391770', 'G9149325', 'G0006196', 'G0446874']
    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(list_of_consultants))] = 'Out of scope'

    # in response to an email from Marisa McAllister (15/09/20), two staff in Oncology have been removed from scope.
    oncology_150920 = ['G9850476', 'G9848847']
    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(oncology_150920))] = 'Out of scope'
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

    falls_piv = pd.pivot_table(df[df['Job_Family'].isin(['Nursing and Midwifery', 'Allied Health Profession',
                               'Medical and Dental'])], index=['Job_Family','Sector/Directorate/HSCP'],
                               columns='Falls Compliant',
                               values='Pay_Number', aggfunc='count', fill_value=0, margins=True, margins_name='All Staff')
    falls_piv['Not at work ≥ 28 days'] = falls_piv['Secondment'] + falls_piv['Maternity Leave'] + \
                               falls_piv['≥28 days Absence'] + falls_piv['Suspended']


    falls_piv['In scope'] = falls_piv['Complete'] + falls_piv['No Account'] \
                            + falls_piv['Not Undertaken'] + falls_piv['Not Complete'] + falls_piv['Not at work ≥ 28 days']
    falls_piv['Compliance %'] = ((falls_piv['Complete'] + falls_piv['Not at work ≥ 28 days']) / falls_piv['In scope'] * 100).round(2)
    falls_piv = falls_piv[falls_piv['Compliance %'] > 5]
    falls_piv.drop(columns=['Maternity Leave', 'Suspended', '≥28 days Absence', 'Secondment', 'Out of scope',
                            'All Staff'], inplace=True)
    falls_piv = falls_piv[['Complete', 'Not at work ≥ 28 days', 'Not Undertaken', 'Not Complete', 'No Account', 'In scope', 'Compliance %']]
    # To protect privacy of those who are off work >28 days. For debugging of absence type, comment this line.
    df.loc[((df['Falls Compliant'].isin(['Secondment', 'Out of Scope', 'Maternity Leave', 'Suspended']))),
           'Falls Compliant'] = 'Not at work ≥ 28 days'

    # write to book
    with pd.ExcelWriter('W:/Learnpro/HSE Falls/'+learnpro_date.strftime('%Y%m%d')+' - HSE Falls.xlsx') as writer:
        df.to_excel(writer, sheet_name='Export', index=False)
        falls_piv.to_excel(writer, sheet_name='Summary')
                # TODO add pivot
        # piv.to_excel(writer, sheet_name='pivot')
    writer.save()

    #write to Named Lists HSE folder
    df.drop(columns=['Pay_Number', 'WTE', 'Contract_Description', 'NI_Number', 'Date_Started', 'Job_Description', 'Post_Descriptor',
                     'Pay_Band'], inplace=True)
    with pd.ExcelWriter('W:/Learnpro/Named Lists HSE/'+learnpro_date.strftime('%Y-%m-%d')+'/'+learnpro_date.strftime('%Y%m%d')+' - HSE Falls.xlsx') as writer:
        df.to_excel(writer, sheet_name='Export', index=False)
        falls_piv.to_excel(writer, sheet_name='Summary')
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

    df.loc[((df['ID Number'].isin(long_abs)) & (
        ~df['Falls Compliant'].isin(['Complete', 'Out of scope']))), 'Falls Compliant'] = '≥28 days Absence'
    df.loc[((df['ID Number'].isin(mat)) & (
        ~df['Falls Compliant'].isin(['Complete', 'Out of scope']))), 'Falls Compliant'] = 'Maternity Leave'
    df.loc[((df['ID Number'].isin(secondment)) & (
        ~df['Falls Compliant'].isin(['Complete', 'Out of scope']))), 'Falls Compliant'] = 'Secondment'
    df.loc[((df['ID Number'].isin(susp)) & (
        ~df['Falls Compliant'].isin(['Complete', 'Out of scope']))), 'Falls Compliant'] = 'Suspended'


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
