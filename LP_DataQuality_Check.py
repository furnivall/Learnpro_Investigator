"""Aim of this is to look at specific weird quirks in Learnpro data and report on them as necessary"""

import pandas as pd
from tkinter.filedialog import askdirectory
import os

sharps_courses = ['GGC: Management of Needlestick & Similar Injuries',
                  'NES: Prev. and Mgmt. of Occ. Exposure (Assessment)']

def take_in_dir(list_of_modules):
    """Takes in directory with prompt then selects all learnpro files within that folder"""

    # counters for lp/nonlp files within dir
    lp_count = 0
    non_lp_count = 0

    # prompt for dir

    # TODO swap this back when testing is done
    dirname = askdirectory(initialdir='C:/Learnpro_Extracts',
                           title="Choose a directory full of learnpro files."
                           )

    # initialise master dataframe
    master = pd.DataFrame()
    modules = []
    # read in historic empower data

    user_ids = pd.DataFrame()

    # iterate through directory
    for file in os.listdir(dirname):
        # operate on the LEARNPRO files first
        if 'LEARNPRO' in file:

            print(file)
            lp_count += 1
            # read file into df
            df = pd.read_csv(dirname + "/" + file, skiprows=14, sep="\t")
            print("Initial length: " + str(len(df)))
            modules.append(df['Module'].unique().tolist())
            # drop fails
            df = df[df['Passed'] == "Yes"]
            print("Removed fails - new length: " + str(len(df)))
            df = df[['ID Number', 'Module', 'Assessment Date']]
            # TODO this would be a good place to add any other cleaning steps
            # drop modules not in list
            df = df[df['Module'].isin(list_of_modules)]
            print("Removed extra modules - new length: " + str(len(df)))

            # add data to master file
            master = master.append(df, ignore_index=True)

            print(file + " added to master, current size= " + str(len(master)))
        elif 'users' in file:
            df = pd.read_csv(dirname + "/" + file, skiprows=11, sep="\t")
            user_ids = df

        else:
            print("not lp")
            non_lp_count += 1

    # log outputs below:
    print(str(lp_count + non_lp_count) + " files read. " + str(lp_count) + " contained learnpro data, " + str(
        non_lp_count) +
          " did not.")
    master['Module'] = master['Module'].astype('category')

    master['Assessment Date'] = pd.to_datetime(master['Assessment Date'], format='%d/%m/%y %H:%M')

    return master, user_ids

df, users = take_in_dir(list_of_modules=sharps_courses)


print(len(users))

print(f"{len(users[users['ID Number'].str.match(r'^[gcGC][0-9xX]+$') == True])}")

users[users['ID Number'].str.len() > 8].to_excel('C:/Learnpro_Extracts/lp_data_quality/too_long_ids.xlsx', index=False)

users[users['ID Number'].str.match(r'^[gcGC][0-9]+$') == False].to_excel('C:/Learnpro_Extracts/lp_data_quality/non_standard_ids.xlsx', index=False)
