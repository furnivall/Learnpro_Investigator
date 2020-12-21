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
    sd = pd.read_excel('W:/Staff Downloads/2020-11 - Staff Download.xlsx')

    sd['Concat'] = sd['Cost_Centre'] + sd['Job_Family']
    sd['Scope'] = 1
    sd['Scope'].loc[~sd['Concat'].isin(concat)] = 0
    print(sd.columns)

    return sd

def exclusions(sd):
    marisa_011020 = ['G0530247', 'G0001226', 'G9200959', 'G5820162', 'G5843480', 'G9490728',
                     'G9293493', 'G5898005', 'G5913004', 'G5925134', 'G5925045', 'G5926807',
                     'G1094505', 'G0988448', 'G0001854', 'G9356738', 'G9860558', 'G9867672',
                     'G5878527', 'G0638080', 'G0510297', 'G0799661', 'G074414X', 'G1009095']

    sd['Scope'].loc[(sd['Pay_Number'].isin(marisa_011020))] = 0

    # Fiona Smith 29-10-20
    sd['Scope'].loc[(sd['Pay_Number'] == 'G9379711')] =0

    # Robert Macfarlane - 19-11-20
    sd['Scope'].loc[(sd['Pay_Number'] == 'G9887181')] = 0


    # kim kilgour 28-10-20 - time sensitive
    kim_kilgour281020 = ['G9831479', 'G9838034', 'G9842367']
    if learnpro_date < pd.to_datetime('01-01-21', dayfirst=True):
        sd['Scope'].loc[(sd['Pay_Number']).isin(kim_kilgour281020)] = 0


    # craig broadfoot 28-10-20 - time sensitive
    craig_broadfoot281020 = ['G9860114', 'G9888845', 'G9844999', 'G9844920', 'G9843272', 'G9851460', 'G3018970']
    if learnpro_date < pd.to_datetime('01-01-21', dayfirst=True):
        sd['Scope'].loc[(sd['Pay_Number']).isin(craig_broadfoot281020)] = 0

    sd['Scope'].loc[(sd['Pay_Number'] == 'G9489983')] = 0

    HPN_081020 = ['G2979179', 'G9884902', 'G9173730', 'G9871651', 'G0092665', 'G0415367', 'G0636363', 'G0890243',
                  'G155462X', 'G0021369', 'G9878834', 'G0862231', 'G0867209', 'G0534471', 'G3893979', 'G3814874',
                  'G3900649', 'G1554018', 'G1561588', 'G1535226', 'G0579939', 'G0865087', 'G9181571', 'G0849227',
                  'G9102558', 'G3903788', 'G9111336', 'G3913449', 'G9883796', 'G0004598', 'G9124098', 'G0003759',
                  'G9881147', 'G9288023', 'G9304878', 'G9888665', 'G9871650', 'G9336087', 'G9871624', 'G9521305',
                  'G9881150', 'G0001790', 'G9888644', 'G9527621', 'G9535160', 'G9860943', 'G9873935', 'G9545336',
                  'G0001400', 'G0003775', 'G0000660', 'G0000914', 'G9884552', 'G0006964', 'G9887636', 'G9867411',
                  'G9872532', 'G9880637', 'G9878395', 'G9874404', 'G9869257', 'G9867466', 'G9858430', 'G9878582',
                  'G9532730', 'G9850598', 'G9850600', 'G9850398', 'G9842512', 'G9842405', 'G9841219', 'G9841226',
                  'G9838376', 'G9833812', 'G9832195', 'G9850612', 'C3916200', 'C1177877', 'G9392319', 'C1174592',
                  'C9535357', 'C4007441', 'C1005022', 'G9873407', 'C302234X', 'G3821528', 'C1074261', 'C1023853',
                  'C1093312', 'C1002600', 'C208418X', 'G1525891', 'G0003030', 'C9548548', 'C954853X', 'G9859167',
                  'C1178563', 'G9878791', 'C9564152', 'C9562400', 'G9866804', 'G9867430', 'G9887614', 'G9877381',
                  'G9880814', 'G9874624', 'G9867109', 'G9860237', 'C3905462', 'G9858237', 'G9854484', 'G9850620',
                  'G9850619', 'G9841156', 'G9840539', 'G9840325', 'G9837753', 'G9834706', 'G9833830', 'G153730X',
                  'G1543997', 'G1573063', 'G1093339', 'G1521640', 'G0865516', 'G9877510', 'G9155090', 'G919858X',
                  'G9328319', 'G9871028', 'G9869720', 'G9872849', 'G9867943', 'G9867406', 'G9859513', 'G9859505',
                  'G9850639', 'G9850218', 'G9830615', 'G9155104', 'G1570862', 'G1517937', 'G1520830', 'G1551353',
                  'G3017885', 'G915521X', 'G1543946', 'G9338187', 'G9874538', 'G1522019', 'G9863254', 'G1557343',
                  'G1532480', 'G1545310', 'G0388939', 'G0923133', 'G0917028', 'G0882356', 'G0915203', 'G0873438',
                  'G0102369', 'G0872024', 'G091374X', 'G0764892', 'G1007718', 'G0443565', 'G0873969', 'G0921300',
                  'G0765910', 'G917141X', 'G0925004', 'G1053485', 'G0882771', 'G1034006', 'G109338X', 'G0862355',
                  'G9304193', 'G1522833', 'G1520776', 'G9329714', 'G1507648', 'G1557289', 'G0385492', 'G1552503',
                  'G1548948', 'G044748X', 'G0769983', 'G1509284', 'C1035827', 'G1535706', 'G0291226', 'G0834807',
                  'G0769916', 'G0385972', 'G0835668', 'G9868723', 'G1573764', 'G1570137', 'G3901777', 'G2364034',
                  'G1504924', 'G9864560', 'G9533206', 'G1571656', 'G9536051', 'C1122088', 'G1570145', 'G9234012',
                  'G1573519', 'G0004825', 'G9521011', 'G9110712', 'G1573918', 'G9398783', 'G9520996', 'G9147934',
                  'G9377603', 'G9154663', 'G9878667', 'G9409351', 'G0002408', 'G9233385', 'G9409114', 'G9233822',
                  'G9239316', 'G9263179', 'G927555X', 'G9328114', 'G932822X', 'G9329285', 'G9521445', 'G9507310',
                  'G9508554', 'G9508643', 'G9529012', 'G9531718', 'G0002818', 'G9535128', 'G9536256', 'G9539417',
                  'G9562923', 'G0002186', 'G9869247', 'G0000585', 'G9878069', 'G9878076', 'G0001987', 'G0003856',
                  'G9860172', 'G9880866', 'G9859633', 'G9888696', 'G9888141', 'G9881055', 'G9887878', 'G9887877',
                  'G9887687', 'G9867440', 'G9880704', 'G9880000', 'G9874374', 'G9875599', 'G9874625', 'G9874060',
                  'G9873828', 'G9872637', 'G9870680', 'G9863256', 'G9867685', 'G9867446', 'G9867444', 'G9867443',
                  'G9864763', 'G9863852', 'G9859642', 'G9859635', 'G9859632', 'G9859620', 'G9859372', 'G9867578',
                  'G934313X', 'G9870679', 'G9859604', 'G9859640', 'G9856875', 'G9854734', 'G9854214', 'G9851182',
                  'G9851185', 'G9850728', 'G9850727', 'G9850726', 'G9850715', 'G9850714', 'G9850711', 'G9850707',
                  'G9850706', 'G9850666', 'G9850725', 'G9850723', 'G9850710', 'G9850705', 'G9850667', 'G9850407',
                  'G9848968', 'G9842065', 'G9841748', 'G9841415', 'G9841150', 'G9841147', 'G9841148', 'G9839389',
                  'G9839388', 'G9834868', 'G1540327', 'G2367009']
    sd['Scope'].loc[(sd['Pay_Number'].isin(HPN_081020))] = 0

    vent_serv = ['G9843023', 'G5901545', 'G9447873', 'G1556312', 'G1532863', 'G1515403', 'G1565281', 'G1553402',
                 'G9529187', 'G9122389', 'G9131604', 'G9291156', 'G9424008', 'G9434860', 'G9434909', 'G9468781',
                 'G9871484', 'G956005X', 'G9864742', 'G0007382', 'G9862644', 'G9862641', 'G9858700', 'G9858697',
                 'G3843408', 'G0924156', 'G2966263', 'G1553429', 'G1552376', 'G941150X', 'G9131639', 'G9365257',
                 'G9276521', 'G9343830', 'G9379371', 'G9379398', 'G9447822', 'G9447881', 'G9468838', 'G9529233',
                 'G9888292', 'G9864741', 'G9864735', 'G9849346', 'G9847068', 'G9847063', 'G9843013', 'G9842940',
                 'G9842518', 'G9839025', 'G9839023', 'G9838965', 'G9837824', 'G9837822', 'G9837821', 'G9837820',
                 'G1572466', 'G9834956']
    sd['Scope'].loc[(sd['Pay_Number'].isin(vent_serv))] = 0
    esp_081020 = ['G3867919', 'G3902595', 'G0649228', 'G0004909', 'G5884705', 'G5861985', 'G4927583', 'G491984X', 'G7077777',
          'G7144091', 'C113521X', 'C1006908', 'C1085271', 'G7119968', 'G9498206', 'G9179348', 'G9431608', 'G2363496',
          'G9882243', 'G0463728', 'G716405X', 'G9842553', 'G9841051', 'G5880483']
    sd['Scope'].loc[(sd['Pay_Number'].isin(esp_081020))] = 0

    claire_goodfellow_20201028 = ['C1009753', 'G1514660', 'G1567268', 'C1167103', 'G1523902', 'G9880908',
                                 'G1572539', 'G9491783', 'G1538233', 'G1544799', 'G9328181', 'G1537261',
                                 'G9833179', 'G9873505', 'G9881777', 'G9869350', 'G110974X', 'G0916943',
                                 'G0000527', 'G9146318', 'G9857320', 'G1542109', 'G9888296', 'G1544829',
                                 'G1534920', 'G0002843', 'G9866842', 'G9870418', 'G9374698', 'G9268308',
                                 'G1514245', 'G1529625', 'G9490949', 'G9853525', 'G1540564', 'C9549250',
                                 'G0001276', 'G9410759', 'G0878200', 'G9837403', 'G9154108', 'G0738646',
                                 'G0923486', 'C1174843', 'G108643X', 'G0773646', 'G9885915', 'G086949X',
                                 'G1573934', 'G1533363', 'G0773735', 'G1529188', 'G1531816', 'G9484280',
                                 'G9870467', 'G9855547', 'G1543873', 'G0927228', 'G1511130', 'G152089X',
                                 'G0921432', 'G9491198', 'G1523953', 'G1551655', 'G1512048', 'G9869580',
                                 'G1503006', 'G9830620', 'G1503502', 'G9867000', 'G1557262', 'G9358110',
                                 'G0871729', 'G9840499', 'G1511211', 'G0878812', 'G9859232', 'G1551140',
                                 'G9839385', 'G1573535', 'G9375104', 'G1506935', 'G1521500', 'G0927112',
                                 'G1078836', 'G9873543', 'G0924113', 'G9358153', 'G3008622', 'G1512099',
                                 'G9843174', 'G9874735', 'G0769940', 'G1509004', 'G1565915', 'G9871035'
                                 'G0007040', 'G9871222']
    sd['Scope'].loc[sd['Pay_Number'].isin(claire_goodfellow_20201028)] = 0

    # Jean Still
    sd['Scope'].loc[sd['Cost_Centre'] == 'G67062'] = 0
    sd['Scope'].loc[sd['Pay_Number'].isin(['G7132980', 'G7160968'])] = 0

    # Isobel Cairney
    sd['Scope'].loc[sd['Pay_Number'].isin(['G0000205', 'G9197079', 'G0950858'])] = 0


    # in response to query from Aline Williams (19/10/2020)
    sd['Scope'].loc[sd['department'] == 'Rah - Pain Service'] = 0

    # in response to Val Cuthill - 28-10-20
    if learnpro_date > pd.to_datetime('01-01-2021', dayfirst=True):
        sd['Scope'].loc[sd['Pay_Number'] == 'G9832554'] = 0

    # Heather Richardson
    sd['Scope'].loc[sd['Pay_Number'] == 'G9566090'] = 0

    # Leigh-Ann Leggat query 28-10-20
    sd['Scope'].loc[sd['Pay_Number'] == 'C9547606'] = 0

    radiogs = ['G5887283', 'G0799173', 'G5907004', 'G9172343', 'G0013919', 'G9314970', 'G108531X', 'C1080385', 'C1148281']
    sd['Scope'].loc[(sd['Pay_Number'].isin(radiogs))] = 0

    outreach = ['G9844683', 'G9845361', 'G9845366', 'G9844969', 'G9843857', 'G9843860', 'G9845728', 'G9837889',
                'G9832929', 'G9830408']
    sd['Scope'].loc[sd['Pay_Number'].isin(outreach)] = 0

    sd['Scope'].loc[sd['Pay_Number'].isin(['G708546X', 'G0007609', 'G0665614'])] = 0

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
off_work = mat + susp + secondment + long_abs


def eESS():
    filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data',
                               title="Choose the relevant M&H Training Extract file"
                               )
    eESS_data = pd.read_csv(filename, encoding='utf-16', sep='\t')
    lookup = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data',
                               title="Choose the pay number lookup file"
                               )
    pay_number_lookup = pd.read_csv(lookup, encoding='utf-16', sep='\t')
    pay_number_lookup.rename(columns={'Payroll Number':'Pay Number'}, inplace=True)
    eESS_data['Assignment Number'] = eESS_data['Assignment Number'].str.slice(0, 8)
    pay_number_lookup['Assignment Number'] = pay_number_lookup['Assignment Number'].str.slice(0, 8)
    eESS_data = eESS_data.merge(pay_number_lookup, how='left', on='Assignment Number')

    eESS_data = eESS_data.drop_duplicates(subset=['Pay Number', 'Course Name'])
    eESS_data = defaultdict(lambda: 0, dict(eESS_data.groupby('Pay Number')['# Enrolment Places'].sum()))

    # eESS_data = defaultdict(lambda:0,dict(eESS_data['Pay Number'].value_counts()))
    print(eESS_data)

    filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/M&H',
                               title="Choose the relevant eESS Assessors file"
                               )
    eESS_assessors = pd.read_excel(filename)
    print(eESS_assessors.columns)
    eESS_assessors = defaultdict(lambda: 0, dict(eESS_assessors['Payroll Number'].value_counts()))
    print(eESS_assessors)
    return eESS_data, eESS_assessors

def ni_fix(df):
    """This little function does a lot - it gets all NI numbers with duplicates, then creates a dataframe for each dupe,
    checks the compliance for each course across all relevant pay numbers, then gets the most recent course completion
    date then applies these to all of the pay numbers"""

    # find dupes and put in list
    dups = df[df.duplicated(subset='NI_Number')]['NI_Number'].drop_duplicates().to_list()

    # this is to make df.update() be able to compare indices and amend data
    df.set_index('Pay_Number')

    courses = ['GGC', 'NES']
    catlookup = {'Out of scope': 0, '0': 1, '1': 2, '2': 3, '3 or more': 4}
    inv_lookup = {v: k for k, v in catlookup.items()}
    # loop through courses, first making all compliant if necessary, then editing expiry dates
    for number, i in enumerate(dups):
        dfx = df[df['NI_Number'] == i]

        dfx['Value'] = df['Assessed Category'].map(catlookup)
        max_val = dfx['Value'].max()
        if max_val > 1:
            max_val = inv_lookup.get(max_val)
            dfx.loc[:, 'Assessed Category'] = max_val
            dfx.drop(columns=['Value'], inplace=True)
            df.update(dfx)
    df.reset_index()
    return df

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
    sd['Assessed Category'].cat.add_categories('Not at work ≥ 28 days', inplace=True)
    sd['Assessed Category'].loc[sd['Scope'] == 0] = 'Out of scope'

    print(sd.columns)
    sd = ni_fix(sd)
    sd['Assessed Category'].loc[
        (sd['Assessed Category'] == '0') & (sd['Pay_Number'].isin(off_work))] = 'Not at work ≥ 28 days'
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
    assessor_piv = pd.pivot_table(sd, index=['Area', 'Sector/Directorate/HSCP'],
                                  columns='Assessor - Both',
                                  values='Pay_Number', aggfunc='count', fill_value=0, margins=True,
                                  margins_name='All Staff')
    with pd.ExcelWriter('W:/Learnpro/M&H/MH Extract Paynums '+learnpro_date.strftime('%Y%m%d')+'.xlsx') as writer:
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
