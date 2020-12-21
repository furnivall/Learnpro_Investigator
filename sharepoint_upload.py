import os
import shutil
print(os.listdir('Z:'))

sectors = ['eHealth','Diagnostic Services1', 'East Renfrewshire HSCP', 'North Sector1', 'Other Functions1',
  'East Dunbartonshire HSCP', 'Inverclyde HSCP', 'Women  Childrens1', 'GGC2', 'Regional Services1',
  'Clyde Sector', 'PPFM','West Dun HSCP', 'Glasgow City HSCP', 'Renfrewshire HSCP', 'South Sector1']



def check_for_HSE_file():
    date = input('what HSE date are you looking at? (format = YYYY-MM-DD)')
    while(os.path.isdir('W:/Learnpro/Named Lists HSE/'+date)==False):
        date = input("Error - a file matching that date doesn't seem to exist yet. Please try again (format = YYYY-MM-DD)")
    print("Checking sectors for most recent upload")
    count = 0
    incomplete = []
    # if you want to replace existing files in a day, uncomment the below:
    # for j in sectors:
    #     print('Copying files to ' + j)
    #     shutil.rmtree('Z:/' + j + '/HSE Reports/' + date)
    #     print("done.")
    for i in sectors:
        if (os.path.isdir('Z:/'+i+'/HSE Reports/'+date)):
            print(f'{i} - up to date')
            count += 1
        else:
            print(f'{i} - file not found, most recent file is: {os.listdir("Z:/"+i+"/HSE Reports")[-1]}')
            incomplete.append(i)
    print(f'incomplete: {incomplete}')
    if count < len(sectors):
        updateBool = input(f"Complete - all dirs scanned, {count}/{len(sectors)} up to date - do you want to update the remaining files?")
        if updateBool.lower() == 'y':
            for j in incomplete:
               print('Copying files to '+j)
               shutil.copytree('W:/Learnpro/Named Lists HSE/'+date,'Z:/'+j+'/HSE Reports/'+date)
               print("done.")

        else:
            print("Thanks, closing program now.")




check_for_HSE_file()
exit()

for i in sectors:
    f = open("Z:/"+i+"/HSE Reports/test.txt", "w+")
    f.close()

#list files
for i in sectors:
    list = os.listdir('Z:/'+i+'/HSE Reports')
    list.sort()
    print(list[-1:])

for i in sectors:
    os.remove("Z:/"+i+"/HSE Reports/test.txt")

#list files
for i in sectors:
    list = os.listdir('Z:/'+i+'/HSE Reports')
    list.sort()
    print(list[-1:])