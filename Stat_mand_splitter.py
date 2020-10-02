import pandas as pd
import numpy as np
from tkinter.filedialog import askopenfilename

filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/Named Lists',
                           title="Choose the relevant stat/mand file - "
                                 "it should be the 'GGC Pay Nums' file for the relevant month."
                           )
df = pd.read_excel(filename, sheet_name='data')
print(df.columns)

sectors = df['Sector/Directorate/HSCP'].unique()
print(sectors)
coldic = {'SM1':'Equality & Diversity', 'SM2':'Fire Safety', 'SM3':'Health & Safety', 'SM4':'Infec. Control', 'SM5':'Info Governance', 'SM6':'Manual Handling', 'SM7':'Public Protection', 'SM8':'Security & Threat', 'SM9':'Violence & Aggression'}
for area in df['Area'].unique():
    dfw = df[df['Area'] == area]
    piv = pd.pivot_table(dfw, index='Sector/Directorate/HSCP',
                         values=['SM1', 'SM2', 'SM3', 'SM4', 'SM5', 'SM6', 'SM7', 'SM8', 'SM9'],
                         aggfunc=np.sum, margins=True)
    piv.loc['% Compliance'] = round(piv.iloc[-1] / len(dfw) * 100, 1)
    piv.rename(columns=coldic)
    dfw.drop(columns=['ID Number'], inplace=True)
    with pd.ExcelWriter('W:/Learnpro/Named Lists/2020-08/2020-08 '+area+'.xlsx') as writer:
        dfw.to_excel(writer, sheet_name='data', index=False)
        piv.to_excel(writer, sheet_name='pivot')
for sector in sectors:
    dfx = df[df['Sector/Directorate/HSCP'] == sector]
    piv = pd.pivot_table(dfx, index='Sub-Directorate 1',
                         values=['SM1', 'SM2', 'SM3', 'SM4', 'SM5', 'SM6', 'SM7', 'SM8', 'SM9'],
                         aggfunc=np.sum, margins=True)
    # produce compliance percentage row
    piv.loc['% Compliance'] = round(piv.iloc[-1] / len(dfx) * 100, 1)
    dfx.drop(columns=['ID Number'], inplace=True)
    piv.rename(columns=coldic)
    # write to book
    with pd.ExcelWriter('W:/Learnpro/Named Lists/2020-08/2020-08 '+sector+'.xlsx') as writer:
        dfx.to_excel(writer, sheet_name='data', index=False)
        piv.to_excel(writer, sheet_name='pivot')
    writer.save()
    subdirs = dfx['Sub-Directorate 1'].unique()
    if len(subdirs) > 1:
        for subdir in subdirs:
            dfy = dfx[dfx['Sub-Directorate 1'] == subdir]
            piv = pd.pivot_table(dfx, index='Sub-Directorate 2',
                                 values=['SM1', 'SM2', 'SM3', 'SM4', 'SM5', 'SM6', 'SM7', 'SM8', 'SM9'],
                                 aggfunc=np.sum, margins=True)
            piv.loc['% Compliance'] = round(piv.iloc[-1] / len(dfx) * 100, 1)
            piv.rename(columns=coldic)
            with pd.ExcelWriter('W:/Learnpro/Named Lists/2020-08/2020-08 ' + sector + ' - ' + subdir.replace('/',
                                                                                                             "'").replace('&',"'") + '.xlsx') as writer:
                dfy.to_excel(writer, sheet_name='data', index=False)
                piv.to_excel(writer, sheet_name='pivot')
                writer.save()



