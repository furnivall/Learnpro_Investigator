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
    # df = pd.read_excel(file)
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
    df.to_csv('C:/Tong/eESS_data_test.csv', index=False)
    return df


def NESScope(df):
    print('Generating NES Scope')
    # Estates & Facilities
    df = df[~(df['Sector/Directorate/HSCP'] == 'Estates and Facilities')]

    # Support Services
    df = df[~(df['Job_Family'] == 'Support Services')]

    # Specific exemptions within radiography:
    radiographers_16092020 = ['G5887283', 'G0799173', 'G5907004', 'G9172343', 'G0013919']
    df = df[~(df['Pay_Number'].isin(radiographers_16092020))]

    df = df[~(df['Pay_Number'] == 'G9231447')]

    # Linda Arthur
    linda_arthur_202021106 = ['G9376291', 'G9376305', 'G9376313', 'G9852132', 'G7094930', 'G7077491', 'G9829789']

    # AHPs - removed then re-added diagnostic and therapeutic radiography on 16/09/2020
    df = df[~(df['Sub_Job_Family'].isin(['Orthotics']))]

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

    # Sarah Dickson - req from Gillian Davie - 22-09-20
    df = df[~(df['Pay_Number'] == 'G9836615')]

    # Claire Goodfellow - 22-10-20
    df = df[~(df['Pay_Number'].isin(
        ['G1111922', 'G0006757', 'G9834581', 'G0006757', 'G9871222', 'G0001494', 'G9871035', 'G9888166']))]

    # Gillian Gall
    df = df[df['Pay_Number'] != 'G9838705']

    # Christine Brownlie
    df = df[df['Pay_Number'] != 'G0445649']

    # Lynn Smith
    df = df[df['Pay_Number'] != 'G9841051']

    # Catherine Nivison
    df = df[~(df['Pay_Number'].isin(['G9838979', 'G9841872', 'G9838926', 'G9828935', 'G9872210']))]

    # Alison Leiper
    df = df[df['Pay_Number'] != 'G9839094']

    print(f'NES Inscope: {len(df)}')

    return df


def GGCScope(df):
    print('Generating GGC Scope')

    df_orig = df

    # donna mooney
    dm241120 = ['G7158327', 'G7123698', 'G9853847', 'G9170294', 'G9838146', ' G9876217', 'G3043002', 'G7155328']
    df = df[~(df['Pay_Number'].isin(dm241120))]

    # Jillian Pounds
    df = df[~(df['Pay_Number'] == 'G9838216')]

    # Tom Quinn
    tq_19112020 = ['G9424490', 'G9828619', 'G9830171', 'G9832581', 'G9848211', 'G9828738', 'G9830329', 'G9830328',
                   'G9828710',
                   'G1537831', 'G9829089', 'G9830169', ' G9874561']
    df = df[~(df['Pay_Number'].isin(tq_19112020))]

    # Kevin Tolland
    df = df[df['Pay_Number'] != 'G9830111']

    # Janice Niven
    janice_niven_20201028 = ['G9872242', 'G9844137', 'G0185116', 'G928317X', 'G9844137', 'G0185116', 'G9832754',
                             'G0588709']
    df = df[~(df['Pay_Number']).isin(janice_niven_20201028)]

    # Geraldine Cairns 28-10-20
    ger_cairns20201028 = ['G9197885' 'G9867991' 'C1158031' 'G9170472' 'G9847251' 'G9847253'
                          'G9839964' 'G9846775' 'G1087576' 'G7123744' 'G9532226' 'G9859180'
                          'G9866627' 'G7142250' 'G9877625' 'G9867976' 'G9376291' 'G9376305'
                          'G9376313' 'G9852132' 'G7094930' 'G7077491' 'G9879429' 'G9887236'
                          'G911887X' 'G9837798']
    df = df[~(df['Pay_Number'].isin(ger_cairns20201028))]

    # Christina Heuston - 28-10-20
    df = df[df['Pay_Number'] != 'G9848514']
    df = df[df['Pay_Number'] != 'G936448X']
    df = df[~(df['Pay_Number'].isin(['G0135569', 'G9536671', 'G7077939']))]

    # Joyce Brown - 15/10/20
    df = df[df['Pay_Number'] != 'G9851266']

    # Jacqueline Pollock
    jp20201124 = ['G9401113', 'C9513639', 'G9855505', 'G4958780', 'G9111549', 'G926485X', 'G9863835']
    df = df[~(df['Pay_Number'].isin(jp20201124))]

    # Allison Morrison
    df = df[df['Pay_Number'] != 'G9829089']

    # Triple P
    df = df[
        ~(df['Pay_Number'].isin(['G9877625', 'G9867976', 'G9881708', 'G9866999', 'G9354247', 'G9864202', 'G9864175']))]

    # Heather Richardson - 28-10-20
    df = df[df['Pay_Number'] != 'G9566090']

    # Stephen Wallace - 30-10-20
    df = df[~(df['Cost_Centre'].isin(['G43217', 'G43221', 'G49009', 'G22515']))]

    # Catriona Reid - 28-10-20
    cat_reid281020 = ['G9837681', 'G7161069', 'G4950259', 'G914109X', 'G9849451', 'G7154224', 'G9840589', 'G9190503',
                      'G9852573']
    df = df[~df['Pay_Number'].isin(cat_reid281020)]

    # In response to query from Aline Williams (19/10/20)
    df = df[df['department'] != 'Gri Macmillan Bereavement Cntr']

    # In response to query from Cindy Wallis, (16/08/20), removed some of the Primary Care MH team in East Ren
    pcmh20201016 = ['G7154224', 'G9849451', 'G9840589', 'G9840356', 'G9340696']
    df = df[~df['Pay_Number'].isin(pcmh20201016)]

    # In response to query from Arwel Williams (28/08/20), removed certain training grades neuro docs
    arwel_neuro = ['G9854092', 'G9843249', 'G9853174', 'G9854035', 'G9853491', 'G9853144', 'G9853581', 'G9853542',
                   'G9830229']
    df = df[~df['Pay_Number'].isin(arwel_neuro)]

    # In response to query from Marisa McAllister (14/09/20), removing Michele Barrett - quality manager - non-clinical
    df = df[~(df['Pay_Number'] == 'G9830586')]

    # In response to Natalie Mcmillan / Margaret Anderson retirement notice - 01/09/2020
    df = df[~(df['Pay_Number'] == 'G0009520')]

    # In response to Amanda Parker - estates + facilities sharps - 15/09/2020
    amanda_parker = ['G9857762', 'G9858380', 'C6508413', 'C3007189']
    df = df[~(df['Pay_Number'].isin(amanda_parker))]

    # Another Marisa McAllister one - 22-09-20
    df = df[~(df['Pay_Number'] == 'G921741X')]

    # Another Marisa McAllister one - 20-10-20
    df = df[~(df['Pay_Number'].isin(['G9856432', 'G9828708']))]

    # mary mcfarlane - 25-09-20
    gri_radiology = ['G9566090', 'G3890759', 'G1098217', 'G3850226', 'G3882802', 'G3889629', 'G9336532', 'G3895777',
                     'G1035800', 'G108433X', 'G0674540', 'G4957792', 'G3905462', 'G1098748', 'G3884260', 'G3891968',
                     'G3875199', 'G9887197', 'G9885618', 'G9246916', 'G9534482', 'G0714321', 'G9874312', 'G9869381',
                     'G9863409', 'G9855443', 'G3895343', 'G3803090']
    df = df[~(df['Pay_Number'].isin(gri_radiology))]

    sach_radiology = ['G0018392', 'G4905105', 'G1537652', 'G9405461', 'G9392106', 'G9241809', 'G4903501', 'G4915712',
                      'G9847152', 'G0685089', 'G9847151', 'G9847152', 'G0464627']
    df = df[~(df['Pay_Number'].isin(sach_radiology))]

    # Amanda Parker - 03-11-20
    ap20201103 = ['C1074326', 'C3020681', 'C3901769', 'C3904474', 'C390721X', 'C391755X', 'C9540725', 'G7055986',
                  'G7071760',
                  'G7089937', 'G7090862', 'G7091230', 'G7092555', 'G7098030', 'G7105126', 'G7112300', 'G7113188',
                  'G7127308',
                  'G7137540', 'G7152132', 'G7155077', 'G9193928', 'G9227288', 'G9380493', 'G9435816', 'G9446036',
                  'G9515720',
                  'G9858545', 'G9874129', 'G5892562', 'G5924391', 'G9837163', 'G9561056', 'G9107487', 'G3875997',
                  'G5825806',
                  'G0880205', 'G0860425', 'G0866733', 'C9540687', 'G7071949']
    df = df[~(df['Pay_Number'].isin(ap20201103))]

    # Sector Level Exclusions
    df = df[~df['Sector/Directorate/HSCP'].isin(['Acute Corporate', 'Board Medical Director', 'Board Administration',
                                                 'Centre For Population Health', 'Corporate Communications', 'eHealth',
                                                 'Finance', 'Public Health'])]
    df = df[~((df['Sector/Directorate/HSCP'] == 'HR and OD') & (df['Sub-Directorate 1'] != 'Occupational Health'))]

    # Sandra Quinn
    df = df[~(df['Pay_Number'].isin(['G913090X', 'C3018830', 'G0171344', 'G0005079', 'G9209972', 'G9862157',
                                     'G3854426', 'C2090112', 'G9850559', 'G5866448', 'G9850556', 'G9835914',
                                     'G2354128', 'G9835172', 'G9840490', 'C2058464', 'G9205594', 'G9828395',
                                     'G9849073', 'G9236074', 'G9220410', 'G9865462', 'G2998335', 'G0433187',
                                     'G9829115']))]

    # Lisa Dorrian - non time sensitive
    df = df[~(df['Pay_Number'].isin(['G9828449', 'G9828443', 'G9829252', 'G9828922']))]

    # craig broadfoot 28-10-20 - time sensitive
    craig_broadfoot281020 = ['G9860114', 'G9888845', 'G9844999', 'G9844920', 'G9843272', 'G9851460', 'G9847464',
                             'G3018970']
    if learnpro_date < pd.to_datetime('01-01-21', dayfirst=True):
        df = df[~(df['Pay_Number'].isin(craig_broadfoot281020))]

    kim_kilgour281020 = ['G9831479', 'G9838034', 'G9842367']
    if learnpro_date < pd.to_datetime('01-01-21', dayfirst=True):
        df = df[~(df['Pay_Number'].isin(kim_kilgour281020))]
        df = df[~(df['Pay_Number'].isin(['G7134622', 'G0129828']))]  # Annette Gibson
        df = df[~(df['Pay_Number'] == 'G0028657')]  # Janice Naven
        df = df[~(df['Pay_Number'] == 'C9561935')]  # Biereonwu, Isobel
        df = df[~(df['Pay_Number'].isin(['G5880378', 'G7112750', 'G9448098', 'G9854574']))]  # Lorna Hill

    df = df[~(df['Pay_Number'] == 'G9379711')]
    df = df[~(df['Pay_Number'] == 'G7135025')]

    fiona_taylor261120 = ['G9877976', 'G7094892', 'G0401110']
    df = df[~(df['Pay_Number'].isin(fiona_taylor261120))]

    alison_shields261120 = ['G9829869', 'G9828925', 'G9831118']

    janis_young_temp = ['G9864540', 'G2348306', 'G933923X', 'G9833063', 'G9831963', 'G9832354', 'G9873283', 'G9833333',
                        'G9832692', 'G9839500', 'G9832897', 'G9835222', 'G9833074', 'G9833065', 'G9832850', 'G3001687',
                        'G9830338', 'G9832700', 'G9848801', 'G9836970', 'G9849219', 'G9373020']
    if learnpro_date < pd.to_datetime('01-01-21', dayfirst=True):
        df = df[~(df['Pay_Number'].isin(janis_young_temp))]

    # Rosie Cherry 30-10-20
    df = df[~(df['Pay_Number'].isin(['G8006431', 'G2986604', 'G7068212']))]

    # Suzanne Catterall
    suz_c_20201208 = ['G9837902', 'G0124915', 'G1107542', 'G7085680', 'G9832075', 'G7055153', 'G9168915', 'G9471901',
                      'G9845794',
                      'G9174362', 'G9869165', 'G7075995', 'G7135718', 'G9830410', 'G9837863', 'G9156879', 'G9860398',
                      'G9324895', 'G9828157',
                      'G9870472', 'G7116748', 'G1010379', 'G9433147', 'G9263942', 'G7090269', 'G9828858', 'G9152717',
                      'G9225293',
                      'G7140509', 'G1503456', 'G0006956', 'G9872219', 'G9849710', 'G7107781', 'G916197X', 'G1573322',
                      'G9847439',
                      'G9862629', 'G9870767']
    df = df[~(df['Pay_Number'].isin(suz_c_20201208))]

    # Patricia Doherty
    pat_d_20201209 = ['G7137486', 'G0414956', 'G9874636', 'G2965941', 'G9838232', 'G92993562', 'G9878297', 'G0002894']
    df = df[~(df['Pay_Number'].isin(pat_d_20201209))]

    # Jean Still
    df = df[~(df['Cost_Centre'] == 'G67062')]
    df = df[~(df['Pay_Number'].isin(['G7132980', 'G7160968', 'G9828861']))]

    # Margaret Valenti 30-10-20
    df = df[~(df['Cost_Centre'].isin(['G20515', 'G22515']))]

    # Job Family Exclusions
    df = df[~df['Job_Family'].isin(['Administrative Services', 'Personal and Social Care', 'Executive'])]
    df = df[~((df['Job_Family'] == 'Other Therapeutic') & (df['Sub_Job_Family'] != 'Optometry'))]

    df = df[~((df['Area'] == 'Partnerships') & (
        df['Sub_Job_Family'].isin(['Speech and Language Therapy', 'Speech And Language Therapy'])))]

    # Sub job family exclusions
    df = df[~(df['Sub_Job_Family'].isin(
        ['Arts Therapies', 'Dietetics', 'AHP Training/Administration', 'Generic Therapies']))]

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

    df = df[~(df['Pay_Number'].isin(['G7132980', 'G7160968', 'G9828861', 'G9828708']))]
    # Another Marisa McAllister one - 20-10-20
    df = df[~(df['Pay_Number'].isin(['G9856432', 'G0405019']))]

    # re-adding someone to scope test
    print(df_orig[df_orig['Pay_Number'] == 'G0001291'])
    df = df.append(df_orig[df_orig['Pay_Number'] == 'G0001291'], ignore_index=True)
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

    # update eESS dictionary

    eESS_lookup = pd.read_csv(dirname + '/' + 'Pay Number to Assignment Number.csv', encoding='utf-16', sep='\t')
    eESS_lookup['Assignment Number'] = eESS_lookup['Assignment Number'].str.slice(0, 8)
    eESS_lookup.dropna(axis=0, inplace=True)
    eESS_lookup['Assignment Number'] = eESS_lookup['Assignment Number'].astype(int)
    eESS_lookup = dict(zip(eESS_lookup['Assignment Number'], eESS_lookup['Payroll Number']))
    with open('C:/Learnpro_Extracts/eesslookup.txt', 'w') as filehandler:
        print(eESS_lookup, file=filehandler)

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

        elif 'CompliancePro Extract' in file:
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
    sd = pd.read_excel('W:/Staff Downloads/2020-11 - Staff Download.xlsx')

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
             'GGC Source', 'NES Module', 'NES Date']]

    ggc_piv = pd.pivot_table(df, index=['Area', 'Sector/Directorate/HSCP'], columns='GGC Module', values='Pay_Number',
                             aggfunc='count', margins=True, fill_value=0, margins_name='All Staff')

    ggc_piv['Not at work ≥ 28 days'] = ggc_piv['Maternity Leave'] + ggc_piv['Suspended'] + ggc_piv[
        '≥28 days Absence']  # ggc_piv['Secondment'] + \

    ggc_piv['In scope'] = ggc_piv['Complete'] + ggc_piv['Expired'] + ggc_piv['No Account'] + ggc_piv['Not undertaken'] + \
                          ggc_piv['Not at work ≥ 28 days']
    ggc_piv['Compliance %'] = (
                (ggc_piv['Complete'] + ggc_piv['Not at work ≥ 28 days']) / ggc_piv['In scope'] * 100).round(2)
    ggc_piv.drop(inplace=True, columns=['All Staff', 'Suspended', 'Out Of Scope', 'Maternity Leave',  # 'Secondment',
                                        '≥28 days Absence'])
    ggc_piv = ggc_piv[ggc_piv['Compliance %'] > 0.01]

    nes_piv = pd.pivot_table(df, index=['Area', 'Sector/Directorate/HSCP'], columns='NES Module', values='Pay_Number',
                             aggfunc='count', margins=True, margins_name='All Staff', fill_value=0)

    nes_piv['Not at work ≥ 28 days'] = nes_piv['Maternity Leave'] + nes_piv['Suspended'] + nes_piv['Secondment'] + \
                                       nes_piv['≥28 days Absence']

    nes_piv['In scope'] = (
            nes_piv['Complete'] + nes_piv['Expired'] + nes_piv['No Account'] + nes_piv['Not undertaken'] +
            nes_piv['Not at work ≥ 28 days']).round(2)

    nes_piv['Compliance %'] = (
                (nes_piv['Complete'] + nes_piv['Not at work ≥ 28 days']) / nes_piv['In scope'] * 100).round(2)
    nes_piv.drop(inplace=True, columns=['All Staff', 'Suspended', 'Out Of Scope', 'Maternity Leave',  # 'Secondment',
                                        '≥28 days Absence'])
    nes_piv = nes_piv[nes_piv['Compliance %'] > 0.01]

    # These lines below hide the type of absence for privacy reasons. If you want to debug, then comment these out.
    df.loc[
        ((df['NES Module'].isin(['Secondment', 'Out of Scope', 'Maternity Leave', 'Suspended', '≥28 days Absence']))),
        'NES Module'] = 'Not at work ≥ 28 days'
    df.loc[
        ((df['GGC Module'].isin(['Secondment', 'Out of Scope', 'Maternity Leave', 'Suspended', '≥28 days Absence']))),
        'GGC Module'] = 'Not at work ≥ 28 days'

    df['Compliant'] = ''

    df['Compliant'].loc[(df['GGC Module'].isin(['Complete', 'Not at work ≥ 28 days'])) & (
    (df['NES Module'].isin(['Out Of Scope', 'Complete',
                            'Not at work ≥ 28 days'])))] = 1
    df['Compliant'].loc[(df['Compliant'] == '')] = 0

    # write to book
    with pd.ExcelWriter(
            'W:/Learnpro/HSE Sharps and Skins/' + learnpro_date.strftime('%Y%m%d') + ' - HSE Sharps.xlsx') as writer:
        df.to_excel(writer, sheet_name='Export', index=False)
        # TODO add pivot
        # piv.to_excel(writer, sheet_name='pivot')
        ggc_piv.to_excel(writer, sheet_name='GGC Pivot')
        nes_piv.to_excel(writer, sheet_name='NES Pivot')
    writer.save()

    df_ernie = df.drop(columns=['WTE', 'Contract_Description', 'NI_Number', 'Date_Started', 'Job_Description',
                                'Post_Descriptor',
                                'Pay_Band', 'Compliant'])
    df_ernie.to_excel('W:/Learnpro/ernie_' + learnpro_date.strftime('%Y%m%d') + '.xlsx', index=False)

    # HSE Named Lists sharepoint upload - "Compliant Column removed per Cameron's instructions on 09-10-20:
    df.drop(columns=['Pay_Number', 'WTE', 'Contract_Description', 'NI_Number', 'Date_Started', 'Job_Description',
                     'Post_Descriptor',
                     'Pay_Band', 'Compliant'], inplace=True)
    with pd.ExcelWriter(
            'W:/Learnpro/Named Lists HSE/' + learnpro_date.strftime('%Y-%m-%d') + '/' + learnpro_date.strftime(
                    '%Y%m%d') + ' - HSE Sharps.xlsx') as writer:
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

    # wrap up and produce final files
    produce_files(df)


sd = pd.read_excel('W:/Staff Downloads/2020-11 - Staff Download.xlsx')
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
