'''preparing to add to the main frankenlearnpro script - needs to deal with empower training'''
import pandas as pd
import os

master_data = pd.DataFrame()
for file in os.listdir('C:/Learnpro_Extracts/LP_testData_31-05-20'):
    if 'EMPOWER' in file:
        tempdf = pd.read_excel('C:/Learnpro_Extracts/LP_testData_31-05-20/'+file)
        rename_cols = {'SM-Health And Safety, An Introduction GGC002': 'GGC: Health and Safety, an Introduction',
                       'SM-Reducing Risks Of Violence & Aggression GGC003': 'GGC: 003 Reducing Risks of Violence & Aggression',
                       'SM-Equality, Diversity & Human Rights GGC004': 'GGC: Equality, Diversity and Human Rights',
                       'SM-Manual Handling Theory, GGC005': 'GGC: Manual Handling Theory',
                       'SM-Public Protection (Adult Support & Protection And Child': 'Child Protection - Level 1',
                       'SM-Standard Infection Control Precautions, GGC007': 'SM-Standard Infection Control Precautions, GGC007',
                       'SM-Security & Threat GGC008': 'GGC: 008 Security & Threat',
                       'Information Handling': 'GGC: 009 Safe Information Handling'}
        for i in rename_cols.keys():
            x = tempdf[tempdf['course_name'] == i]
            print("****"+file+'****')
            print(i + ": " + str(len(x)))
        master_data = master_data.append(tempdf, ignore_index=True)

for i in master_data['course_name'].unique():
    x = master_data[master_data['course_name'] == i]
    print(i + ": " + str(len(x)))

master_data.to_excel('C:/Learnpro_Extracts/LP_testData_31-05-20/empower_final.xlsx')
exit()

print(df.columns)
print(df['course_name'].unique())

rename_cols = {'SM-Health And Safety, An Introduction GGC002':'GGC: Health and Safety, an Introduction',
               'SM-Reducing Risks Of Violence & Aggression GGC003':'GGC: 003 Reducing Risks of Violence & Aggression',
               'SM-Equality, Diversity & Human Rights GGC004':'GGC: Equality, Diversity and Human Rights',
               'SM-Manual Handling Theory, GGC005':'GGC: Manual Handling Theory',
               'SM-Public Protection (Adult Support & Protection And Child':'Child Protection - Level 1',
               'SM-Standard Infection Control Precautions, GGC007':'SM-Standard Infection Control Precautions, GGC007',
               'SM-Security & Threat GGC008':'GGC: 008 Security & Threat',
               'Information Handling':'GGC: 009 Safe Information Handling'}
for i in rename_cols.keys():
    x = df[df['course_name'] == i]
    print(i + ": " + str(len(x)))









print(list(rename_cols.keys()))