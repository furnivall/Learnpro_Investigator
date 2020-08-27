import pandas as pd
pd.set_option('display.width', 420)
pd.set_option('display.max_columns', 10)
def eESS(file):
    df = pd.read_excel(file)

    eess_courses = ['GGC E&F Sharps - Disposal of Sharps (Toolbox Talks)', 'GGC E&F Sharps - Inappropriate Disposal of Sharps',
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
    df['GGC Source'] = 'eESS'
    return df

print(eESS('C:/Learnpro_Extracts/20200813-auto/eESS.xlsx').head(5))