import pandas as pd
from tkinter.filedialog import askopenfilename






# list of modules required for compliance
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
           'NES: Prev. and Mgmt. of Occ. Exposure (Assessment)']
extra_mods=[
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

def read_file():

    # read in data
    filename = askopenfilename(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data', title="Choose a learnpro file",
                           )

    #gui_code.loading_label.configure(text = "Loading...")
    df = pd.read_csv(filename, skiprows=14, sep='\t')

    print(df.columns)

    # remove extraneous modules
    df = df[(df['Module'].isin(modules)) & (df['Passed'] == "Yes")]

    # check number of passes per module
    for i in modules:
        print(i + str(len(df[df['Module'] == i])))


    # get list of all user ids within this extract
    users = df['Email / Username'].unique().tolist()

    # list of courses outside of the stat/mand ones
    courses = ['GGC: 234 RN Update – COVID-19 Contingencies', 'LBT: Blood Components and Indications for Use',
               'LBT: Safe Blood Sampling for Transfusion Video', 'LBT: Safe Transfusion Practice',
               'GGC: 210 FS Precision Pro - Glucose']


    # code to get all modules contained in a course
    for i in courses:
        print(df[df['Course'] == i]['Module'].unique().tolist())


    # iterate through each module, then list of users. Delete user if they don't have required module pass.
    for module in modules:
        temp_frame = df[df['Module'] == module]['Email / Username'].unique()
        for user in users:
            if user not in temp_frame:
                users.remove(user)
        # print current length of users list
        print(module + " - " + str(len(users)))

    data = pd.DataFrame(users)
    data.to_csv('C:/Learnpro_extracts/compliant_users.csv', index=False)


read_file()


