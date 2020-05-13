import pandas as pd


#read in data
df = pd.read_csv('C:/Learnpro_extracts/data.xls', skiprows=14, sep='\t')


print(df.columns)

#print(df['Course'].unique())

modules = ['GGC: 001 Fire Safety',
           'GGC: Health and Safety, an Introduction',
           'GGC: 003 Reducing Risks of Violence & Aggression',
           'GGC: Equality, Diversity and Human Rights',
           'GGC: Manual Handling Theory',
           'Child Protection - Level 1',
           'Adult Support & Protection',
           'GGC: Standard Infection Control Precautions ',
           'GGC: 008 Security & Threat',
           'GGC: 009 Safe Information Handling',
           'An Introduction to Falls',
           'The Falls Bundle of Care',
           'Risk Factors for Falls (Part 1)',
           'Risk Factors for Falls (Part 2)',
           'What to do when your patient falls',
           'Falls - Bedrails ',
           'GGC: Management of Needlestick & Similar Injuries',
           'NES: Prevention and Management of Occ. Exposure',
           'NES: Prev. and Mgmt. of Occ. Exposure (Assessment)',
           'Infection Prevention and Control Covid-19',
           'Pressure Ulcer Prevention',
           'Medicines Administration',
           'Food Fluid and Nutrition',
           'Falls Prevention and Management',
           'Moving & Handling Quick Guide 2020',
           'NEWS and the Deteriorating Patient',
           'Infection Prevention and Control Covid-19',
           'Management of Transfused Patient ',
           'Blood Group Serology',
           'Administration Procedure',
           'Collection Procedure',
           'Haemovigilance in the UK and ROI',
           'Sampling Procedures',
           'Requesting Procedure',
           'Massive Transfusion',
           'Plasma Components',
           'Adverse Effects of Transfusion',
           'Platelets',
           'Plasma Derivatives',
           'Red Blood Cells',
           'Transfusion Governance and Appropriate Use',
           'Blood Group Serology',
           'Safe Blood Sampling for Transfusion'
           ]

print(len(df))
df = df[(df['Module'].isin(modules)) & (df['Passed'] == "Yes")]
print(len(df))

for i in modules:
    print(i + str(len(df[df['Module'] == i])))

users = df['ID Number'].unique().tolist()
print(type(users))

data = pd.DataFrame()
# list of courses outside of the stat/mand ones
courses = ['GGC: 234 RN Update – COVID-19 Contingencies', 'LBT: Blood Components and Indications for Use',
           'LBT: Safe Blood Sampling for Transfusion Video', 'LBT: Safe Transfusion Practice',
           'GGC: 210 FS Precision Pro - Glucose']


# code to get all modules contained in a course
for i in courses:
    print(i + '\n' + df[df['Course'] == i]['Module'].unique().tolist())


print(df['Course'].unique())




for j in modules:
    temp_frame = df[df['Module'] == j]['ID Number'].unique()
    print(j + str(len(users)))
    for i in users:
        if i not in temp_frame:
            users.remove(i)

data = pd.DataFrame(users)
data.to_csv('C:/Learnpro_extracts/compliant_users.csv', index=False)





