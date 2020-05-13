'''
This takes in a directory with several learnpro files and concatenates them into one master file with some cleaning and
dupe removal.
'''
import pandas as pd
from tkinter import messagebox
from tkinter.filedialog import askdirectory
import os


def take_in_dir():
    """Takes in directory with prompt then selects all learnpro files within that folder"""

    #counters for lp/nonlp files within dir
    lp_count = 0
    non_lp_count = 0

    #prompt for dir
    dirname = askdirectory(initialdir='//ntserver5/generalDB/WorkforceDB/Learnpro/data',
                           title="Choose the relevant absence extract."
                           )

    #initialise master dataframe
    master = pd.DataFrame()


    for file in os.listdir(dirname):
        if 'LEARNPRO' in file:
            print(file)
            lp_count += 1
            df = pd.read_excel(dirname+"/"+file, skiprows=14)
            master = master.append(df, ignore_index=True)
            print(file + "added to master, current size= " + str(len(master)))
        else:
            print("not lp")
            non_lp_count += 1

    print(str(lp_count+non_lp_count) + " files read. " + str(lp_count) +" contained learnpro data, "+str(non_lp_count) +
          " did not.")


take_in_dir()

