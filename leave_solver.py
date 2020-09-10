import pandas as pd
from tkinter.filedialog import askopenfilename
df=pd.read_excel('W:/Daily_Absence/2020-09-03.xls', skiprows=4)
print(df.columns)
print(f'Total absences: {len(df)}')
mat = df[df['AbsenceReason Description'] == 'Maternity Leave']['Pay No'].tolist()
df = df[~df['Pay No'].isin(mat)]
susp = df[df['Absence Type'] == 'Suspended']['Pay No'].tolist()
df = df[~df['Pay No'].isin(susp)]
secondment = df[df['AbsenceReason Description'] == 'Secondment']['Pay No'].tolist()
df = df[~df['Pay No'].isin(secondment)]

df['Absence Episode Start Date'] = pd.to_datetime(df['Absence Episode Start Date'], format='%Y-%m-%d')

day28 = pd.Timestamp.now() - pd.DateOffset(days = 28)
long_abs = df[df['Absence Episode Start Date'] < day28]['Pay No'].tolist()
print(f'Seconded = {len(secondment)}, Mat Leave = {len(mat)}, Suspended = {len(susp)}, Over 28 days = {len(long_abs)}')

df_statmand = pd.read_excel('C:/Learnpro_Extracts/namedList.xlsx', sheet_name='data')
print(df_statmand.columns)


#Stat Mand
df_statmand['Not at work'] = ''

df_statmand.loc[df_statmand['ID Number'].isin(long_abs), 'Not at work'] = 'Not at work ≥ 28 days'
df_statmand.loc[df_statmand['ID Number'].isin(mat), 'Not at work'] = 'Not at work ≥ 28 days'
df_statmand.loc[df_statmand['ID Number'].isin(secondment), 'Not at work'] = 'Not at work ≥ 28 days'
df_statmand.loc[df_statmand['ID Number'].isin(susp), 'Not at work'] = 'Not at work ≥ 28 days'
df_statmand.to_excel('C:/Learnpro_Extracts/statmand_exemptions.xlsx', index=False)


#
# #Falls
# df_falls = pd.read_excel('C:/Learnpro_Extracts/falls/20200904 - HSE Falls.xlsx', sheet_name='data')
#
# df_falls['Not at work'] = ''
# df_falls.loc[((df_falls['Pay_Number'].isin(long_abs)) & (~df_falls['Falls Compliant'].isin(['Complete', 'Out of scope']))), 'Falls Compliant'] = '>=28 days Absence'
# df_falls.loc[((df_falls['Pay_Number'].isin(mat)) & (~df_falls['Falls Compliant'].isin(['Complete','Out of scope']))), 'Falls Compliant'] = 'Maternity Leave'
# df_falls.loc[((df_falls['Pay_Number'].isin(secondment)) & (~df_falls['Falls Compliant'].isin(['Complete', 'Out of scope']))), 'Falls Compliant'] = 'Secondment'
# df_falls.to_excel('C:/Learnpro_Extracts/falls/falls_notatwork.xlsx', index=False)
#
#
# #Sharps
# df_sharps = pd.read_excel('C:/Learnpro_Extracts/sharps/namedList-sharps.xlsx', sheet_name='data')
# df_sharps['Not at work'] = ''
# df_sharps.loc[((df_sharps['Pay_Number'].isin(long_abs)) & (~df_sharps['GGC Module'].isin(['Complete', 'Out Of Scope']))), 'GGC Module'] = '>=28 days Absence'
# df_sharps.loc[((df_sharps['Pay_Number'].isin(mat)) & (~df_sharps['GGC Module'].isin(['Complete', 'Out Of Scope']))), 'GGC Module'] = 'Maternity Leave'
# df_sharps.loc[((df_sharps['Pay_Number'].isin(secondment)) & (~df_sharps['GGC Module'].isin(['Complete', 'Out Of Scope']))), 'GGC Module'] = 'Secondment'
#
# df_sharps.to_excel('C:/Learnpro_Extracts/sharps/sharps_notatwork.xlsx', index=False)