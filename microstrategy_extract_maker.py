import pandas as pd
from tkinter.filedialog import askopenfilename
import numpy as np

current_month = input("What's the current reporting (i.e. staff download) month - format = mm-yy")
current_month = pd.to_datetime(current_month, format='%m-%y')



filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/Named Lists',
                           title="Choose the relevant stat/mand file - "
                                 "it should be the 'GGC Pay Nums' file for the relevant month."
                           )
df = pd.read_excel(filename, sheet_name='data')
df.rename(columns={'At work?':'Headcount'}, inplace=True)
print(df.columns)
df.to_csv('C:/Tong/testdata.csv')

df = df[df['Headcount']==1]
df['Headcount'] = df['Headcount'].astype(int)
piv = pd.pivot_table(df, index=['Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1',
                                                    'Sub-Directorate 2', 'department', 'Cost_Centre'],
                     values=['SM1', 'SM2', 'SM3', 'SM4','SM5', 'SM6', 'SM7', 'SM8', 'SM9', 'Headcount'],
                     aggfunc=np.sum
                     )
piv.to_excel('C:/Tong/test_ms.xlsx')
piv.reset_index(inplace=True)
print(piv.columns)


rename_cols = {'SM1':'Equality Diversity', 'SM2':'Fire Safety', 'SM3':'Health Safety', 'SM4':'Infection Control',
               'SM5':'Information Gov.', 'SM6':'Manual Handling', 'SM7':'Public Protection', 'SM8':'Security and Threat',
               'SM9':'Violence Aggression'}
piv.rename(columns=rename_cols, inplace=True)
piv['Date'] = current_month.strftime('%m/%d/%Y')

piv = piv[['Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1',
       'Sub-Directorate 2', 'department', 'Cost_Centre', 'Headcount',
       'Equality Diversity', 'Fire Safety', 'Health Safety',
       'Infection Control', 'Information Gov.', 'Manual Handling',
       'Public Protection', 'Security and Threat', 'Violence Aggression',
       'Date']]
filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Microstrategy/Data/Statutory Mandatory Training',
                           title="Choose the most recent microstrategy file"
                           )
old_data = pd.read_excel(filename)
old_data = old_data[old_data['Date']!=current_month]




old_data = old_data.append(piv, ignore_index=True)
#old_data['Date'] = old_data['Date'].dt.strftime('%m/%d/%Y')
old_data = old_data.to_excel('W:/Microstrategy/Data/Statutory Mandatory Training/SM_Upload_'
                             +pd.Timestamp.now().strftime('%Y%m%d')+'.xlsx', index=False)