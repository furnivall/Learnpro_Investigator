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
    sd = pd.read_excel('W:/Staff Downloads/2020-11 - Staff Download.xlsx')
    print(sd.columns)
    sd['Concat'] = sd['Cost_Centre'] + sd['Job_Family']
    sd['Falls Compliant'] = ""
    sd['Falls Compliant'].loc[~sd['Concat'].isin(concat)] = 'Out of scope'

    # Due to some weird new cost centres being created by Clyde for all training grades docs,
    # we've now moved these into scope.
    sd['Falls Compliant'].loc[sd['Cost_Centre'].str.startswith('GZCL')] = ''
    # Same as above but for GDSG grades
    sd['Falls Compliant'].loc[sd['Cost_Centre'].str.startswith('GDSG')] = ''

    sd['Falls Compliant'].loc[sd['Cost_Centre']=='G16169'] = ''

    # In response to query from Pauline Simpson (28/08/2020), we have removed all Obstetrics staff from the scope list
    sd['Falls Compliant'].loc[sd['Sub-Directorate 2'] == 'Obstetrics'] = 'Out of scope'

    # In response to a query from Stephanie Mckay (08/09/2020), we have added "Qeuh-neuro + Omfs Opd" to the
    # scope for Falls
    sd['Falls Compliant'].loc[sd['department'] == 'Qeuh-neuro + Omfs Opd'] = ''

    sd['Falls Compliant'].loc[sd['Pay_Number'] == 'G9829789'] = ''



    # in response to query from Aline Williams (19/10/2020)
    sd['Falls Compliant'].loc[sd['department'] == 'Rah - Pain Service'] = 'Out of scope'
    sd['Falls Compliant'].loc[sd['Pay_Number'] == 'G3025365'] = 'Out of scope'
    # In response to a query from Joyce Brown (15/10/2020) we have removed someone from scope
    sd['Falls Compliant'].loc[sd['Pay_Number'] == 'G9851266'] = 'Out of scope'

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
                           'G3001725', 'G9472290', 'G9391770', 'G9149325', 'G0006196', 'G0446874', 'G3046141']
    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(list_of_consultants))] = 'Out of scope'

    # in response to an email from Marisa McAllister (15/09/20), two staff in Oncology have been removed from scope.
    oncology_150920 = ['G9850476', 'G9848847']
    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(oncology_150920))] = 'Out of scope'

    # A second email from Marisa McAllister (16/09/20) removed many more staff.
    oncology_160920 = ['G0530247', 'G9834415', 'G9848926', 'G9864463', 'C1105078', 'G3015033', 'G9423079', 'G9382534',
                       'G9489770', 'G9312749', 'G5947219', 'G5947251', 'G9530576', 'G956313X', 'C1149369', 'G9857313',
                       'G9868320', 'G4926544', 'G4922964', 'G9200959', 'G9873026', 'G9169105', 'G0937959']
    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(oncology_160920))] = 'Out of scope'

    # Email from Patricia Lang (18/09/20) removed another few staff.
    plastics180920 = ['G0000367', 'G3889035', 'G3906051', 'G9249109', 'G9862696', 'G949197X', 'G9278737', 'G0004951',
                      'G9231153', 'G9841302', 'G9841603', 'G5920647','G9839432', 'G9843224', 'G9862696', 'G9157476',
                      'G9264469', 'G9850002', 'G9842187', 'G9842896', 'G9278737', 'G9659684', 'G3880435', 'G0000367',
                      'G9864203', 'G388970X', 'G5845696', 'G9249109', 'G09852914', 'G3906051', 'G9851616', 'G9880676',
                      'G5877628', 'G9853044', 'G949197X', 'G3804704', 'G9835614', 'G3889025', 'G9865031', 'G3858707',
                      'G00004951', 'G3813061']

    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(plastics180920))] = 'Out of scope'

    #john carson 23-11-20
    jc231120 = ['G9356274', 'G9862852', 'G9873827', 'G3821978', 'G9184333', 'G9870974', 'G9321047', 'C1064568',
                'G9878923', 'G7109830', 'G7109849', 'G0014206', 'G1055453', 'G7141637', 'C3003922', 'G9181830',
                'G9849400', 'G9404007', 'C9523502', 'G9838544', 'G9867335']
    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(jc231120))] = 'Out of scope'

    renal210920 = ['G0000367', 'G3889035', 'G3906051', 'G9249109', 'G9862696', 'G949197X', 'G9278737', 'G0004951',
                   'G9231153', 'G9841302', 'G9841603', 'G5920647', 'G9556788', 'G591597X', 'G3853551', 'G9858884',
                   'G9843276', 'G5836840', 'G107458X', 'G9867883', 'G5891507', 'G385454X', 'G9473025', 'G5866057',
                   'G5912511', 'G9875028', 'G9140689', 'G9217142', 'G9868363', 'G0005791', 'G5865980', 'G391111X',
                   'G9882128', 'G9443274', 'G9520147', 'G9843279']
    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(renal210920))] = 'Out of scope'

    neurology250920 = ['G0000003', 'G0003450', 'G9846003', 'G9878810', 'G2966808', 'G9854590', 'G9186832', 'G9871839',
                       'G9878294', 'G9888845', 'G9882775', 'G3044947', 'G9831099', 'G9839077', 'G9849270', 'G1120417',
                       'G9864697', 'G9830781', 'G9842969']
    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(neurology250920))] = 'Out of scope'



    gri_radiology = ['G9566090', 'G3890759', 'G1098217', 'G3850226', 'G3882802', 'G3889629', 'G9336532', 'G3895777',
                     'G1035800', 'G108433X', 'G0674540', 'G4957792', 'G3905462', 'G1098748', 'G3884260', 'G3891968',
                     'G3875199', 'G9887197', 'G9885618', 'G9246916', 'G9534482', 'G0714321', 'G9874312', 'G9869381',
                     'G9863409', 'G9855443', 'G3895343']
    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(gri_radiology))] = 'Out of scope'

    sach_radiology = ['G0018392', 'G4905105', 'G1537652', 'G9405461', 'G9392106', 'G9241809', 'G4903501', 'G4915712',
                      'G9847152', 'G0685089', 'G9847151', 'G9847152', 'G0464627', 'G9194428']
    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(sach_radiology))] = 'Out of scope'

    # Heather Richardson quick 3 - 27-10-20
    heather_271020 = ['G9854861', 'G9836783', 'G0003491', 'G9566090']
    sd['Falls Compliant'].loc[(sd['Pay_Number'].isin(heather_271020))] = 'Out of scope'

    # craig broadfoot 28-10-20 - time sensitive
    craig_broadfoot281020 = ['G9860114', 'G9888845', 'G9844999', 'G9844920', 'G9843272', 'G9851460', 'G9847464', 'G3018970']
    if learnpro_date < pd.to_datetime('01-01-21', dayfirst=True):
        sd['Falls Compliant'].loc[(sd['Pay_Number']).isin(craig_broadfoot281020)] = 'Out of scope'

    kim_kilgour281020 = ['G9831479', 'G9838034', 'G9842367']
    if learnpro_date < pd.to_datetime('01-01-21', dayfirst=True):
        sd['Falls Compliant'].loc[(sd['Pay_Number']).isin(kim_kilgour281020)] = 'Out of scope'

    #Pauline Simpson requested that consultants in her depts be removed from falls
    sd['Falls Compliant'].loc[(sd['Sub-Directorate 2']=='Gynaecology') & (sd['Sub_Job_Family'] == 'Consultant')] = 'Out of scope'

    #Heather Richardson requested that all training grades in this cost centre be removed from scope
    sd['Falls Compliant'].loc[(sd['Cost_Centre'] == 'G40520') & (sd['Sub_Job_Family'] == 'Training Grades')] = 'Out of scope'

    # Jean Still
    sd['Falls Compliant'].loc[sd['Cost_Centre'] == 'G67062'] = 'Out of scope'
    sd['Falls Compliant'].loc[sd['Pay_Number'].isin(['G7132980', 'G7160968'])] = 'Out of scope'

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
    falls_piv['Not at work ≥ 28 days'] = falls_piv['Maternity Leave'] + \
                                        +falls_piv['≥28 days Absence'] \
                                         + falls_piv['Suspended']
                                            # falls_piv['Secondment'] +\

    falls_piv['In scope'] = falls_piv['Complete'] + falls_piv['No Account'] \
                            + falls_piv['Not Undertaken'] + falls_piv['Not Complete'] + falls_piv['Not at work ≥ 28 days']
    falls_piv['Compliance %'] = ((falls_piv['Complete'] + falls_piv['Not at work ≥ 28 days']) / falls_piv['In scope'] * 100).round(2)
    falls_piv = falls_piv[falls_piv['Compliance %'] > 5]
    falls_piv.drop(columns=['Maternity Leave', 'Suspended', '≥28 days Absence',  'Out of scope', #'Secondment',
                            'All Staff'], inplace=True)
    falls_piv = falls_piv[['Complete', 'Not at work ≥ 28 days', 'Not Undertaken', 'Not Complete', 'No Account', 'In scope', 'Compliance %']]
    # To protect privacy of those who are off work >28 days. For debugging of absence type, comment this line.
    df.loc[((df['Falls Compliant'].isin(['Secondment', 'Out of Scope', 'Maternity Leave', 'Suspended', '≥28 days Absence']))),
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
