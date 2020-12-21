"""This file looks for staff within the staff download based on the concatenated info produced by L&E"""

import pandas as pd
from tkinter.filedialog import askopenfilename
import numpy as np

extract_filename = askopenfilename(initialdir='W:/Learnpro/')
df = pd.read_excel(extract_filename)

sd_filename = askopenfilename(initialdir='W:/Staff Downloads')
sd = pd.read_excel(sd_filename)


print(df.columns)

usable_folk = {}

print(sd['Date_of_Birth'])

for i in df['ID'].unique().tolist():


    forename, surname, day, month, year = str.split(i, sep='_')
    date = pd.to_datetime(day + "/" + month + "/" + year, format='%d/%m/%Y')
    curr_search = sd[(sd['Forename'] == forename) & (sd['Surname'] == surname) & (sd['Date_of_Birth'] == date)]
    if len(curr_search.index) == 1:
        usable_folk[i] = curr_search['Pay_Number'].iloc[0]

print(usable_folk)
output = pd.DataFrame.from_dict(usable_folk, orient='index', columns=['Pay_Number'])
output.index.name='Old LP ID'
output.to_csv('W:/Learnpro/LE_found/'+pd.Timestamp.now().strftime('%Y%m%d')+'.csv')