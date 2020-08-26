import pandas as pd
from tkinter.filedialog import askopenfilename
df=pd.read_excel('W:/Daily_Absence/2020-08-19.xls', skiprows=4)
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
long_abs = df[df['Absence Episode Start Date'] < day28]
print(f'Seconded = {len(secondment)}, Mat Leave = {len(mat)}, Suspended = {len(susp)}, Over 28 days = {len(long_abs)}')
