'''
This takes in a directory with several learnpro files and concatenates them into one master file with some cleaning and
dupe removal.
'''


import pandas as pd
from tkinter.filedialog import askdirectory
import os
import numpy as np
time = pd.Timestamp.now()

with open("C:/Learnpro_Extracts/eesslookup.txt", "r") as file:
    line = file.read()
eesslookup = eval(line)

staff_download = pd.read_excel('W:/Staff Downloads/2020-04 - Staff Download.xlsx')
staff_ids = staff_download['Pay_Number'].unique().tolist()

stat_mand = ['GGC: 001 Fire Safety',
             'GGC: Health and Safety, an Introduction',
             'GGC: 003 Reducing Risks of Violence & Aggression',
             'GGC: Equality, Diversity and Human Rights',
             'GGC: Manual Handling Theory',
             'Child Protection - Level 1',
             'Adult Support & Protection',
             'GGC: Standard Infection Control Precautions ',
             'GGC: 008 Security & Threat',
             'GGC: 009 Safe Information Handling',
             'Safe Information Handling',
             ]


def take_in_dir(list_of_modules):
    """Takes in directory with prompt then selects all learnpro files within that folder"""

    # counters for lp/nonlp files within dir
    lp_count = 0
    non_lp_count = 0

    # prompt for dir
    dirname = askdirectory(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data',
                           title="Choose a directory full of learnpro files."
                           )

    # initialise master dataframe
    master = pd.DataFrame()

    # iterate through directory
    for file in os.listdir(dirname):
        # operate on the LEARNPRO files first
        if 'LEARNPRO' in file:
            print(file)
            lp_count += 1
            # read file into df
            df = pd.read_csv(dirname + "/" + file, skiprows=14, sep="\t")
            print("Initial length: " + str(len(df)))
            # drop fails
            df = df[df['Passed'] == "Yes"]
            print("Removed fails - new length: " + str(len(df)))
            # TODO this would be a good place to add any other cleaning steps
            # drop modules not in list
            df = df[df['Module'].isin(list_of_modules)]
            print("Removed extra modules - new length: " + str(len(df)))
            # add data to master file
            master = master.append(df, ignore_index=True)
            print(file + " added to master, current size= " + str(len(master)))
        else:
            # TODO deal with empower, eESS etc
            print("not lp")
            non_lp_count += 1
    # log outputs below:
    print(str(lp_count + non_lp_count) + " files read. " + str(lp_count) + " contained learnpro data, " + str(
        non_lp_count) +
          " did not.")
    df_users = master['ID Number'].unique().tolist()
    return master, df_users


df, learnpro_users = take_in_dir(stat_mand)

print(df.columns)
print(str(len(staff_ids)) + " employees in Staff Download")
print(str(len(learnpro_users)) + " learnpro users found.")

print(str(len(df['ID Number'].isin(staff_ids))) + " learnpro accounts with matching pay number")
print(str(len(staff_download[staff_download['Pay_Number'].isin(learnpro_users)])) + " pay numbers with matching learnpro account")

df_temps = df[df['ID Number'].str.contains("T", na=False)]
print(str(len(df_temps)) + " temporary accounts found.")

module_dfs = {}
for module in stat_mand:
    module_dfs[module] = df[df['Module'] == module]

for module in module_dfs:
    print(module, str(len(module_dfs.get(module))))

# initialise df

named_list_df = pd.DataFrame()

named_list_df = named_list_df.assign(ID = staff_ids)
named_list_df.set_index("ID", inplace=True)
for module in module_dfs:
    mod = module_dfs.get(module)

    named_list_df[module] = np.where(named_list_df.index.isin(mod["ID Number"]), 1, 0)
    print(mod.columns)
    named_list_df[module+' - date'] = np.where(named_list_df.index.isin(mod["ID Number"]), mod['Expiry Date'], 0)


#
# #iterative method
# for user in staff_ids:
#     data_dic = {}
#     print(user)
#     data_dic['ID Number'] = user
#     for module in module_dfs:
#         #print(module_dfs.get(module)['ID Number'].head(10))
#         mod = module_dfs.get(module)
#         if mod[mod['ID Number'].isin([user])].empty:  # - works, but very slow
#         # if user in mod['ID Number']: does not work, and slow
#             data_dic[module] = 0
#         else:
#             data_dic[module] = 1
#     named_list_df = named_list_df.append(data_dic, ignore_index=True)
# named_list_df.set_index('ID Number', inplace=True)

named_list_df.to_excel('C:/Learnpro_Extracts/final.xlsx')
print(pd.Timestamp.now() - time)
