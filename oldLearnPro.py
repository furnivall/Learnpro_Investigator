import os
import pandas as pd
import numpy as np
#read in all files
path = 'C:/Learnpro_Extracts/path/'
def sd():
    df = pd.read_excel('W:/Staff Downloads/2020-04 - Staff Download.xlsx')
    df = df[['Pay_Number', 'NI_Number']]
    df.columns = ['ID Number', 'NI_Number']
    return df
sd = sd()

# courses = [line.decode('utf-8-sig').strip() for line in open('W:/Learnpro/PyComp/courses.txt', 'rb')]
# eesslookup = {}
# with open('W:/Learnpro/Pycomp/eesslookup.txt', 'rb') as f:
#     for line in f:
#         line = line.decode('utf-8-sig')
#         (key, val) = line.strip().split('=')
#         eesslookup[key] = val

courses = {"Fire Emergency within the Ward":'fire1',"Fire Fighting Equipment":'fire2',"Fire Prevention":'fire3',
  "Introduction and General Fire Safety":'fire4',"Specialist Roles":'fire5','GGC: 001 Fire Safety':'fire6',"Health and Safety Awareness":'hs1',
  "GGC: Health and Safety, an Introduction":'hs2',"SM-Health And Safety, An Introduction GGC002":'hs3',
  "Violence and Aggression":'va1',"GGC: 003 Reducing Risks of Violence & Aggression":'va2',
  "SM-Reducing Risks Of Violence & Aggression GGC003":'va3',"Introduction to Equality and Diversity":'ed1',
  "GGC: Equality, Diversity and Human Rights":'ed2',"SM-Equality, Diversity & Human Rights GGC004":'ed3',
  "Manual Handling (Non Patient) - Efficient Movement":'mh1',"Manual Handling (Non Patient) - Ergonomics":'mh2',
  "Manual Handling (Non Patient) - Legislation":'mh3',"Manual Handling (Non Patient) â€“ Anatomy":'mh4',
  "Manual Handling (Non Patient) â€“ Causes of Injury":'mh5',"Manual Handling (Patient) - Efficient Movement":'mh6',
  "Manual Handling (Patient) - Ergonomics":'mh7',"Manual Handling (Patient) - Legislation":'mh8',
  "Manual Handling (Patient) â€“ Anatomy":'mh9',"Manual Handling (Patient) â€“ Causes of Injury":'mh10',
  "SM-Manual Handling Theory, GGC005":'mh11',"GGC: Manual Handling Theory":'mh12',"GGC: Child Protection - Level one":'pp1',
  "Child Protection - Level 1":'pp2',"Adult Support and Protection Act":'pp3',"Adult Support & Protection":'pp4',
  "SM-Public Protection (Adult Support & Protection And Child":'pp5',"GGC: Standard Infection Control Precautions":'ic1',
  "GGC: Standard Infection Control Precautions ":'ic2',"SM-Standard Infection Control Precautions, GGC007":'ic3',
  "GGC: 008 Security & Threat":'st1',"SM-Security & Threat GGC008":'st2',"Safe Information Handling":'ig1',
  "SM-Safe Information Handling - Foundation (Information Gov":'ig2'}

eesslookup = {
    'GGC E&F StatMand - Equality, Diversity & Human Rights (face to face session)': 'SM-Equality, Diversity & Human Rights GGC004',
    'GGC E&F StatMand - Reducing Risks of Violence & Aggression(face to face session)': 'SM-Reducing Risks Of Violence & Aggression GGC003',
    'GGC E&F Sharps - Disposal of Sharps (Toolbox Talks)': 'Toolbox Talks',
    'GGC E&F StatMand - General Awareness Fire Safety Training (face to face session)': 'FS-General Awareness Fire Safety Training/DVD',
    'GGC E&F Sharps - Inappropriate Disposal of Sharps': 'Toolbox Talks',
    'GGC E&F Sharps - Management of Injuries (Toolbox Talks)': 'Toolbox Talks',
    'GGC E&F StatMand - Health & Safety an Induction (face to face session)': 'SM-Health And Safety, An Introduction GGC002',
    'GGC E&F StatMand - Public Protection (face to face session)': 'SM-Public Protection (Adult Support & Protection And Child',
    'GGC E&F StatMand - Security & Threat (face to face session)': 'SM-Security & Threat GGC008',
    'GGC E&F StatMand - Standard Infection Control Precautions (face to face session)': 'SM-Standard Infection Control Precautions, GGC007',
    'GGC E&F StatMand - Manual Handling Theory (face to face session)': 'SM-Manual Handling Theory, GGC005'}


def eessread():
    df = pd.read_csv(path + 'CompliancePro Extract.csv', encoding='utf-16', sep='\t')
    df2 = pd.read_csv('W:/LearnPro/Data/20190808/Pay Number to Assignment Number.csv', encoding='utf-16', sep='\t')
    # print (df.columns)
    df['Module'] = df['Course Name'].map(eesslookup)
    df['Assessment Date'] = pd.to_datetime('01/01/2019')
    # print(df2.columns)
    df2.columns = ['Assignment Number', 'ID Number']

    df = df.astype(str).merge(df2.astype(str), on='Assignment Number', how='left')
    df = df[['ID Number', 'Module', 'Assessment Date']]
    df = df[pd.notnull(df['ID Number'])]

    df['Module'] = df['Module'].astype('category')  # hack to make fast
    df.to_csv(path + 'eesstest.csv', index=False)
    #df['Assessment Date'] = pd.to_datetime('01/01/2019')
    print('eESS reading complete')
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


    return master

def lpread():
    learnprofiles = ['LEARNPRO 20160731-20170131.xlsx', 'LEARNPRO 20170131-20170731.xlsx',
                     'LEARNPRO 20170731-20180131.xlsx', 'LEARNPRO 20180131-20180731.xlsx',
                     'LEARNPRO 20180731-20190131.xlsx', 'LEARNPRO 20190131-20190731.xlsx'
                    ]

    df = pd.DataFrame(columns=['ID Number','Module', 'Assessment Date'])
    for i in learnprofiles:
        x = pd.read_excel('W:/LearnPro/Data/20190808/'+i, skiprows=14)
        x = x[x['Passed'] == 'Yes'] #remove fails
        x = x[['ID Number', 'Module', 'Assessment Date']] #remove useless cols
        #print(x.dtypes)
        df['Assessment Date'] = df['Assessment Date'].astype('datetime64[ns]')
        #print(df.dtypes)
        x = x[x['Module'].isin(courses)] #remove out of scope mods
        print('Current import = '+i+', len = '+str(len(x)))
        df = pd.concat([df, x])
        print('Concat length:'+str(len(df)))
    df['Module'] = df['Module'].astype('category') #hack to make fast
    df['Assessment Date'] = pd.to_datetime(df['Assessment Date'], format='%d-%b-%y %H:%M:%S').dt.normalize()
    df = df.sort_values(['ID Number','Module','Assessment Date']).drop_duplicates(['ID Number','Module'], keep='last')

    return df
if os.path.exists('C:/Learnpro/PyComp/1 - datadump.csv'):
    df = pd.read_csv('C:/Learnpro/PyComp/1 - datadump.csv')
    df['Module'] = df['Module'].astype('category')  # hack to make fast
    df['Assessment Date'] = pd.to_datetime(df['Assessment Date'], format='%Y-%m-%d').dt.normalize()
else:
    df1 = eessread()
    print(df1.dtypes)
    df2 = take_in_dir(courses)
    print(df2.dtypes)
    df = pd.concat([df1, df2])
    df.to_csv('W:/Learnpro/PyComp/1 - datadump.csv', index=False)

users = df['ID Number'].drop_duplicates().sort_values().to_frame()
for course in courses:
        print(course)
        df1 = df[df["Module"]==course]
        df1 = df1.drop(columns="Module")
        df1 = df1.rename(columns={"Assessment Date": str(course)})
        print(df1.dtypes)
        users = users.merge(df1,on="ID Number", how="left")
        print(users.dtypes)
print(len(users))



extdate = pd.to_datetime('2019-08-08')  # edit this to the extract date
firedate = extdate - pd.DateOffset(years=1)
print(firedate)
courseall = {

    'mh1': ["Manual Handling (Non Patient) - Efficient Movement",
            "Manual Handling (Non Patient) - Ergonomics",
            "Manual Handling (Non Patient) - Legislation",
            "Manual Handling (Non Patient) â€“ Anatomy",
            "Manual Handling (Non Patient) â€“ Causes of Injury"],
    'mh2': ["Manual Handling (Patient) - Efficient Movement",
            "Manual Handling (Patient) - Ergonomics",
            "Manual Handling (Patient) - Legislation",
            "Manual Handling (Patient) â€“ Anatomy",
            "Manual Handling (Patient) â€“ Causes of Injury"]

}
courseany = {
    'Health and Safety': ["Health and Safety Awareness",
                          "GGC: Health and Safety, an Introduction",
                          "SM-Health And Safety, An Introduction GGC002"],
    'Violence & Aggression': ["Violence and Aggression",
                              "GGC: 003 Reducing Risks of Violence & Aggression",
                              "SM-Reducing Risks Of Violence & Aggression GGC003"],
    'Equality and Diversity': ["Introduction to Equality and Diversity",
                               "GGC: Equality, Diversity and Human Rights",
                               "SM-Equality, Diversity & Human Rights GGC004"],
    'Information Governance': ["Safe Information Handling",
                               "SM-Safe Information Handling - Foundation (Information Gov"],
    'Security and Threat': ["GGC: 008 Security & Threat", "SM-Security & Threat GGC008"],
    'Infection Control': ["GGC: Standard Infection Control Precautions",
                          "GGC: Standard Infection Control Precautions ",
                          "SM-Standard Infection Control Precautions, GGC007"],
    'pp1': ["GGC: Child Protection - Level one", "Child Protection - Level 1"],
    'pp2': ["Adult Support and Protection Act", "Adult Support & Protection",
            "SM-Public Protection (Adult Support & Protection And Child"],
    'mh3': ["SM-Manual Handling Theory, GGC005", "GGC: Manual Handling Theory"]

}


for i in courseall:
    users[i] = np.where(users[courseall[i]].notnull().all(axis='columns'), "Compliant", "Not Compliant")

for i in courseany:
    users[i] = np.where(users[courseany[i]].notnull().any(axis='columns'), "Compliant", "Not Compliant")
fire = ["Fire Emergency within the Ward", "Fire Fighting Equipment", "Fire Prevention",
        "Introduction and General Fire Safety", "Specialist Roles"]
fire2 = 'GGC: 001 Fire Safety'
fired = ["Fire Emergency within the Ward", "Fire Fighting Equipment", "Fire Prevention",
         "Introduction and General Fire Safety", "Specialist Roles", 'GGC: 001 Fire Safety']
users['Fire Safety'] = np.where(users[fire].notnull().all(axis='columns'), "Compliant",
                                np.where(users[fire2].notnull(), "Compliant-New Course", "Not Compliant"))
users['Fire Safety Expiry'] = np.where(users[fire2].notnull(), users[fire2], users[fire].min(axis=1))
users['Health and Safety Expiry'] = users[courseany['Health and Safety']].min(axis=1)
users['Equality and Diversity Expiry'] = users[courseany['Equality and Diversity']].min(axis=1)
users['Violence & Aggression Expiry'] = users[courseany['Violence & Aggression']].min(axis=1)
users['Infection Control Expiry'] = users[courseany['Infection Control']].min(axis=1)
users['Information Governance Expiry'] = users[courseany['Information Governance']].min(axis=1)
users['Security and Threat Expiry'] = users[courseany['Security and Threat']].min(axis=1)
users['Manual Handling'] = np.where((users['mh1'] == 'Compliant') |
                                    (users['mh2'] == 'Compliant') |
                                    (users['mh3'] == 'Compliant'), 'Compliant', 'Non-Compliant')
users['Public Protection'] = np.where((users['pp1'] == 'Compliant') &
                                      (users['pp2'] == 'Compliant'), 'Compliant', 'Non-Compliant')
users.loc[users['Fire Safety Expiry'] < firedate, 'Fire Safety'] = 'Expired'
users.to_csv
users = users[['ID Number', 'Fire Safety', 'Health and Safety', 'Violence & Aggression',
               'Equality and Diversity', 'Manual Handling', 'Public Protection',
               'Infection Control', 'Security and Threat', 'Information Governance',
               'Fire Safety Expiry', 'Health and Safety Expiry', 'Equality and Diversity Expiry',
               'Violence & Aggression Expiry', 'Infection Control Expiry',
               'Security and Threat Expiry', 'Information Governance Expiry']]

users.to_csv('C:/LearnPro_Extracts/2 - test-users.csv', index=False)
sd = sd.merge(users, on='ID Number', how='left')
sd['nullcols'] = sd.isnull().sum(axis=1)
sd['account?'] = np.where(sd['nullcols'] == 16, "No account detected", 1)
cols = sd.columns.tolist()
cols = cols[-1:] + cols[:-1]
sd = sd[cols]
ni = sd['NI_Number']
dfdupz = sd[ni.isin(ni[ni.duplicated()])]
dfdupz = dfdupz['NI_Number'].drop_duplicates(keep='first')
dfdupz.to_csv('C:/LearnPro_Extracts/4 - NI Duplicates.csv', index=False)
matchlist = {'Fire Safety Expiry':'Fire Safety',
            'Health and Safety Expiry':'Health and Safety',
            'Equality and Diversity Expiry':'Equality and Diversity',
            'Violence & Aggression Expiry':'Violence & Aggression',
            'Infection Control Expiry':'Infection Control',
            'Security and Threat Expiry':'Security and Threat',
            'Information Governance Expiry':'Information Governance'}
for i in dfdupz:
    x = sd[sd['NI_Number'] == i]
    for col in matchlist:
        x[col] = pd.to_datetime(x[col])
        xmax = (x[col].max(),x[col].idxmax())
        if pd.notnull(xmax[0]):
            sd.loc[sd['NI_Number'] == i, col] = xmax[0]
            sd.loc[sd['NI_Number'] == i, matchlist[col]] = sd[matchlist[col]].iloc[xmax[1]]


sd.to_csv('C:/LearnPro_Extracts/3 - SD.csv', index=False)

#null_pns = sd.isnull().sum(axis=1)


#null_pns.to_csv('W:/LearnPro/PyComp/4 - NIcorrection.csv', index=False)