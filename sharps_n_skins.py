"""
This file builds on the work already done with the statutory/mandatory portion of the learnpro process and produces a
further report on Sharps & Skins data.
"""

import pandas as pd
from tkinter.filedialog import askdirectory
import os
import datetime as dt
import numpy as np

sd = pd.read_excel('W:/Staff Downloads/2020-05 - Staff Download.xlsx')
print(len(sd['Cost_Centre'].unique()))


def NESScope(df):
    df['NES Module'] = "In scope"
    # print(len(df[df['department'].isin(['Rad-Chaplains', 'Community Engagement Team'])]))

    # bws_excl = df['department'].isin(['Rad-Chaplains', 'Community Engagement Team'])
    # sector_excl = df['Sector/Directorate/HSCP'].isin(['Estates and Facilities', 'eHealth']) & \
    #               df['Area'] == 'Board-Wide Services'
    #
    # jobfam_excl = df['Area'] == 'Partnership' & df['Job_Family'].isin(['Personal And Social Care', 'Support Services', 'Personal and Social Care'])


    # codified list of out of scope people for NES modules:
    # All estates & facilities staff
    est_fac = df[df['Sector/Directorate/HSCP'] == 'Estates and Facilities']
    # All support services staff
    support_services = df[df['Job_Family'] == 'Support Services']
    print(len(est_fac))
    print(len(support_services))
    # All eHealth staff
    ehealth = df[df['Sector/Directorate/HSCP'] == 'eHealth']
    # All personal & social care staff
    # All band 5 health visitors and school nurses
    # All AHPs in the "Orthoptics", "Dietetics", "Occupational Therapy", "Speech and Language Therapy", "Physiotherapy"
    # and "Prosthetics" sub job families
    # All staff within the "Chief Operating Officer" sub job family
    # All staff in "Health Promotion" or "Occupational Therapy" sub job families
    # All Specialist Nurses in Bands 8a or 8b
    # All staff within the "Hscp-Inv: Child Psychiat", "Inv Com Nurs Chld Serv Fund", "Inv Hscp Child Smile"
    # sub-directorates
    # All staff within the "Rad-Chaplains" or "Community Engagement Team" departments

    # out of scope list for GGC modules
    # All staff in Administrative Services, Personal and Social Care, Executives, Other Therapeutic job families
    # All staff in "Optometry", "Speech and Language Therapy", "Arts Therapies", "Dietetics" and "Health Promotion"
    # sub job families
    # All staff in the "Acute Directors", "Board Medical Director", "Board Administration",
    # "Centre for Population Health", "Corporate Communications", "eHealth", "Finance" and "Public Health" sectors
    # All staff in the "Occupational Health" subdirectorate of HR&OD
    # All staff in the "Nurse Director" subdirectorate of Support Services
    # All band 5 Health Visitors and School nurses.
    # All staff within the "Rhc Eeg" department









    #df.loc[bws_excl, 'NES Module'] = "Out of scope"


    #df2 = df[df['NES Module'] == 'Out of scope']

    #print(df2['department'].value_counts())
    return df

def get_sharps_scope():
    pass


def read_in_data():
    """This function will read in the data for the relevant period"""

    pass

def crunch_data():
    """This function will be the parent function for all steps of the data processing stage"""
    pass

def outputs():
    """This function will produce all outputs as required."""
    pass

sd = NESScope(sd)
#print(sd['NES Module'].value_counts())