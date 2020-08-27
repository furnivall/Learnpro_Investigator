'''This file makes specialised empower extracts for given sets of courses. It uses the final Empower data from the
20190405 data folder in W:/Learnpro/Data to pull out only relevant courses'''

import pandas as pd
import os
pd.set_option('display.width', 420)
pd.set_option('display.max_columns', 10)

master = pd.DataFrame()
sharps_courses = ['Toolbox Talks - Disposal Of Sharps', 'FAC- Inappropriate Disposal Of Sharps',
                  'Sharps - Managment Of Injuries (Toolbox Talks)', 'Sharps Disposal (Toolbox Talks)']

for file in os.listdir('C:/Learnpro_Extracts/Empower_Data'):
    print(file)
    dfx = pd.read_excel('C:/Learnpro_Extracts/Empower_Data/'+file)
    print(dfx.columns)
    master = master.append(dfx, ignore_index=False)

master = master[['payroll_number','course_name', 'start_date']]
master.rename(columns={'payroll_number': 'ID Number', 'course_name':'Module', 'start_date':'Assessment Date'},inplace=True)

print(len(master))


sharps = master[master['Module'].isin(sharps_courses)]
print(len(sharps))
print(sharps['Module'].value_counts())
sharps.loc[:, 'Module'] = 'GGC: Management of Needlestick & Similar Injuries'
sharps['GGC Source'] = 'Empower'

sharps.to_excel('C:/Learnpro_Extracts/Empower_Data/final_sharps.xlsx', index=False)