"""This is to test whether we can read files from external sources"""
import urllib
import pandas as pd
import requests
from requests_ntlm import HttpNtlmAuth

df = pd.read_excel('W:/Staff Downloads/2020-07 - Staff Download.xlsx')
print(df.columns)

def example_func(df):
    df = df[df['department'] == 'Hr-Workforce Information']
    return df['Surname'].unique().tolist()



print(f'Workforce info - {example_func(df)}')


exit()

response = requests.get('http://teams.staffnet.ggc.scot.nhs.uk/teams/CorpSvc/HR/WorkInfo/HSE%20Resources/Staff%20Scope%20List.xlsx', auth=HttpNtlmAuth(r'username', 'password'))

with open('W:/LearnPro/HSEScope-current.xlsx', 'wb') as f:
    f.write(response.content)

df = pd.read_excel("W:/LearnPro/HSEScope-current.xlsx")
print(df.columns)
