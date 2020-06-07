import pandas as pd

# df = pd.read_excel('W:/1 - Sandbox Downloads/eESS staff download.xlsx')
# print(len(df))
# print(df.columns)
#
# x = df[['Payroll Number', 'Employee Number']]
# lookup_dic = {row[1]: row[0] for row in x.values}
# with open("C:/Learnpro_Extracts/eesslookup.txt", "w") as file:
#     print(lookup_dic, file=file)

df = pd.read_excel("C:/Learnpro_Extracts/eESS Learning.xlsx")
with open("C:/Learnpro_Extracts/eesslookup.txt", "r") as file:
    data = file.read()
eesslookup = eval(data)


statmand = ['GGC E&F StatMand - Equality, Diversity & Human Rights (face to face session)',
            'GGC E&F StatMand - General Awareness Fire Safety Training (face to face session)',
            'GGC E&F StatMand - Health & Safety an Induction (face to face session)',
            'GGC E&F StatMand - Manual Handling Theory (face to face session)',
            'GGC E&F StatMand - Public Protection (face to face session)',
            'GGC E&F StatMand - Reducing Risks of Violence & Aggression(face to face session)',
            'GGC E&F StatMand - Safe Information Handling Foundation (face to face session)',
            'GGC E&F StatMand - Security & Threat (face to face session)',
            'GGC E&F StatMand - Standard Infection Control Precautions (face to face session)']


df['Pay Number'] = df['Employee Number'].map(eesslookup)
print(df.columns)
print(df['Course Name'].unique())
print(df.head)
df = df[df['Course Name'].isin(statmand)]
print(len(df))
print(df['Enrolment Status'].value_counts())
df = df.rename(columns={'Employee Number': 'ID Number', 'Course Name': 'Module', 'Course End Date': 'Assessment Date'})

df = df[['ID Number', 'Module', 'Assessment Date']]

df.to_csv('C:/Learnpro_Extracts/eESS_test.csv')

