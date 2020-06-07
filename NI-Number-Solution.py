"""
This was made to help with implementing an NI number check for the new learnpro solution
"""
import pandas as pd

df = pd.read_excel('C:/Learnpro_Extracts/namedList.xlsx')
print(df.columns)

print(len(df))

#get list of all NI numbers with multiple roles.
dups = df[df.duplicated(subset='NI_Number')]['NI_Number'].drop_duplicates().to_list()

print(dups)
df.set_index('ID Number')

# test for individual edit
# dfx = df[df['NI_Number'] == 'NE824972C']
# # print(dfx.columns)
# print(dfx['Fire Awareness'])
# dfx.loc[:, 'Fire Awareness'] = 'Complete'
# print(dfx['Fire Awareness'])
#
# print("Initial fire compliance: "+ df[df['NI_Number'] == 'NE824972C']['Fire Awareness'])
#
# df.update(dfx)

print("Ending fire compliance: "+ df[df['NI_Number'] == 'NE824972C']['Fire Awareness'])

# get list of all courses to be altered
start_time = pd.Timestamp.now()
courses = ['Equality, Diversity and Human Rights', 'Fire Awareness', 'Health, Safety & Welfare', 'Infection Control',
           'Information Governance', 'Manual Handling', 'Public Protection','Security and Threat', 'Violence and Aggression']

    # dfx.to_csv('C:/Learnpro_Extracts/2.csv')
    # if number == 2:
    #     exit()

    print(number)


# for i in dups:
#     x = df[df['NI_Number'] == i]
#     #max_fire = x['']


end_time = pd.Timestamp.now()

diff = end_time - start_time


print("time to run: " + str(diff))

df.to_csv('C:/Learnpro_Extracts/NIchecker.csv')

