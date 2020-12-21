"""This file captures NMAHP data from the HSE and stat mand data that's already produced for Ben Law"""

import pandas as pd
from tkinter.filedialog import askopenfilename

def staff_download_bit():
    """
    Captures the staff download from W:/Staff Downloads, cuts down the columns and filters it to only N&M + AHPs
    """
    sd = askopenfilename(initialdir='W:/Staff Downloads/', title='Choose relevant staff download file')
    sd = pd.read_excel(sd)
    sd = sd[['Pay_Number', 'Area', 'Sector/Directorate/HSCP_Code',
           'Sector/Directorate/HSCP', 'Sub-Directorate 1', 'Sub-Directorate 2',
          'department', 'Cost_Centre', 'Surname', 'Forename','Job_Family','Sub_Job_Family','Pay_Band']]

    print(f"Length before filtering = {len(sd)}")
    sd = sd[sd['Job_Family'].isin(['Nursing and Midwifery', 'Allied Health Profession'])]
    print(f"Length after filtering = {len(sd)}")
    return sd

def stat_mand_data():
    """
    Pulls the stat mand data and reduces it down to relevant columns for the merge
    """
    df = askopenfilename(initialdir='W:/Learnpro/Named Lists', title='Choose GGC paynums file for the relevant month')
    df = pd.read_excel(df)
    print(df.columns)
    print(f'Length of stat mand before job fam split: {len(df)}')
    df = df[df['Job_Family'].isin(['Nursing and Midwifery', 'Allied Health Profession'])]
    print(f'length of stat mand after job fam split: {len(df)}')
    df = df[['ID Number', 'SM1', 'SM2', 'SM3', 'SM4','SM5', 'SM6', 'SM7', 'SM8', 'SM9', 'At work?',
             'Equality, Diversity and Human Rights',
             'Fire Awareness', 'Health, Safety & Welfare', 'Infection Control',
             'Information Governance', 'Manual Handling', 'Public Protection',
             'Security and Threat', 'Violence and Aggression'
             ]]

    df = df.rename(columns={'ID Number':'Pay_Number',
                            'Equality, Diversity and Human Rights':'Equality, Diversity and Human Rights status',
                            'Fire Awareness':'Fire Awareness status',
                            'Health, Safety & Welfare':'Health, Safety & Welfare status',
                            'Infection Control':'Infection Control status',
                            'Information Governance':'Information Governance status',
                            'Manual Handling':'Manual Handling status',
                            'Public Protection':'Public Protection status',
                            'Security and Threat':'Security and Threat status',
                            'Violence and Aggression':'Violence and Aggression status',
                            'SM1':'Equality, Diversity and Human Rights',
                            'SM2':'Fire Awareness',
                            'SM3':'Health, Safety & Welfare',
                            'SM4':'Infection Control',
                            'SM5':'Information Governance',
                            'SM6':'Manual Handling',
                            'SM7':'Public Protection',
                            'SM8':'Security and Threat',
                            'SM9':'Violence and Aggression',
                            })
    return df

def get_sharps_data():
    """
    Gets sharps data and reduces to relevant columns
    """
    df = askopenfilename(initialdir = 'W:/Learnpro/HSE Sharps and Skins')
    df = pd.read_excel(df)

    print(f'Length of sharps before job fam split: {len(df)}')
    df = df[df['Job_Family'].isin(['Nursing and Midwifery', 'Allied Health Profession'])]
    print(f'length of sharps after job fam split: {len(df)}')
    df = df[['Pay_Number', 'GGC Module', 'NES Module']]
    print(df['GGC Module'].unique())
    headcount_map = {'Complete':1, 'Expired':1, 'Out Of Scope':0, 'Not undertaken':1, 'Not at work ≥ 28 days':1,
                     'No Account':1}
    complete_map = {'Complete':1, 'Expired':0, 'Out Of Scope':0, 'Not undertaken':0, 'Not at work ≥ 28 days':0,
                     'No Account':0}
    df['SharpsGGC'] = df['GGC Module'].map(complete_map)
    df['SharpsNES'] = df['NES Module'].map(complete_map)
    df['SharpsGGCHeads'] = df['GGC Module'].map(headcount_map)
    df['SharpsNESHeads'] = df['NES Module'].map(headcount_map)
    df = df.rename(columns={'GGC Module':'SharpsGGC status', 'NES Module':'SharpsNES status'})
    df = df[['Pay_Number', 'SharpsGGC', 'SharpsNES', 'SharpsGGCHeads', 'SharpsNESHeads', 'SharpsNES status', 'SharpsGGC status']]
    return df

def falls_data():
    """
    This takes in falls data and reduces it to relevant cols
    """
    df = askopenfilename(initialdir='W:/Learnpro/HSE Falls/', title='Choose the relevant falls data file')
    df = pd.read_excel(df)

    print(f'Length of falls before job fam split: {len(df)}')
    df = df[df['Job_Family'].isin(['Nursing and Midwifery', 'Allied Health Profession'])]
    print(f'length of falls after job fam split: {len(df)}')
    print(df['Falls Compliant'].unique())
    df = df[['Pay_Number', 'Falls Compliant']]
    falls_compliance_map = {'Out of scope':0,'Not Complete':0, 'Complete':1, 'Not at work ≥ 28 days':0, 'No Account':0,
                            'Not Undertaken':0}
    falls_headcount_map = {'Out of scope':0, 'Not Complete':1, 'Complete':1, 'Not at work ≥ 28 days':1, 'No Account':1,
                           'Not Undertaken':1}
    df['Falls'] = df['Falls Compliant'].map(falls_compliance_map)
    df['FallsHeads'] = df['Falls Compliant'].map(falls_headcount_map)
    df.rename(columns={'Falls Compliant':'Falls status'}, inplace=True)
    df = df[['Pay_Number', 'Falls', 'FallsHeads', 'Falls status']]
    return df

def MH():
    df = askopenfilename(initialdir = 'W:/Learnpro/M&H/', title='Choose the relevant M&H Pay Number file')
    df = pd.read_excel(df)

    print(f'Length of M&H before job fam split: {len(df)}')
    df = df[df['Job_Family'].isin(['Nursing and Midwifery', 'Allied Health Profession'])]
    print(f'length of M&H after job fam split: {len(df)}')
    df = df[['Pay_Number', 'Assessed Category']]
    print(df['Assessed Category'].unique())
    mh_map = {'3 or more':1, '0':0, '2':1, '1':1, 'Not at work ≥ 28 days':0, 'Out of scope':0}
    mh_hc = {'3 or more':1, '0':1, '2':1, '1':1, 'Not at work ≥ 28 days':1, 'Out of scope':0}
    df['Moving & Handling'] = df['Assessed Category'].map(mh_map)
    df['MHHeads'] = df['Assessed Category'].map(mh_hc)
    df = df.rename(columns={'Assessed Category':'Moving & Handling status'})
    df = df[['Pay_Number', 'Moving & Handling', 'MHHeads', 'Moving & Handling status']]
    return df



sd = staff_download_bit()
statmand = stat_mand_data()
#exit()
sharps = get_sharps_data()
falls = falls_data()
mh = MH()

for i in [statmand, sharps, falls, mh]:
    sd = sd.merge(i, on='Pay_Number', how='left')

columns = ['Area', 'Sector/Directorate/HSCP', 'Sub-Directorate 1',
       'Sub-Directorate 2', 'department', 'Cost_Centre', 'Pay_Number',
       'Surname', 'Forename', 'Sub_Job_Family', 'Pay_Band',
       'Equality, Diversity and Human Rights status', 'Fire Awareness status',
       'Health, Safety & Welfare status', 'Infection Control status',
       'Information Governance status', 'Manual Handling status',
       'Public Protection status', 'Security and Threat status',
       'Violence and Aggression status', 'SharpsGGC status',
       'SharpsNES status', 'Falls status', 'Moving & Handling status',
       'Equality, Diversity and Human Rights', 'Fire Awareness',
       'Health, Safety & Welfare', 'Infection Control',
       'Information Governance', 'Manual Handling', 'Public Protection',
       'Security and Threat', 'Violence and Aggression', 'SharpsGGC',
       'SharpsNES', 'SharpsGGCHeads', 'SharpsNESHeads', 'Falls', 'FallsHeads',
       'Moving & Handling', 'MHHeads', 'Headcount', 'Job_Family', 'At work?']

sd['Headcount'] = 1
sd = sd[columns]
sd.to_excel('W:/Learnpro/Nursing and Midwifery/NMAHP data-'+pd.Timestamp.now().strftime('%Y%m%d')+'.xlsx',index=False)




