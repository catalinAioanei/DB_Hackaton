import xlrd
import csv

def csv_from_excel():
    wb = xlrd.open_workbook('Hackathon Dummy Data.xlsx')
    sh = wb.sheet_by_name('Team Preferences')
    your_csv_file = open('Team.csv', 'w')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()

# runs the csv_from_excel function:
csv_from_excel()

import pandas as pd
grad = pd.read_csv("Grad.csv")
team = pd.read_csv("Team.csv")
team.head()

grad_preferences = {}
for index, row in grad.iterrows():
    values = []
    values.append(row['Team Preference 1'])
    values.append(row['Team Preference 2'])
    values.append(row['Team Preference 3'])
    values.append(row['Team Preference 4'])
    values.append(row['Team Preference 5'])
    grad_preferences[row['Graduates']] = values
    
team_preferences = {}
for index, row in team.iterrows():
    values = []
    values.append(row['Grad Preference 1'])
    values.append(row['Grad Preference 2'])
    values.append(row['Grad Preference 3'])
    values.append(row['Grad Preference 4'])
    values.append(row['Grad Preference 5'])
    team_preferences[row['Teams']] = values

data = {'Grads' : grad['Graduates']}
matrix = pd.DataFrame(data)
teamList = list(team['Teams'])
for i in range(0,len(teamList)):
    matrix[teamList[i]] = ""
    
for i, row in matrix.iterrows():
    gradName = row['Grads']
#     print(teamList)
    gradPreferences = grad_preferences[row['Grads']]
    for i in range(0,len(gradPreferences)):
        if gradPreferences[i] in teamList:
#             print(gradPreferences[i])
            row[gradPreferences[i]] = i + 1 
#         print("-")
        
#         if gradPreferences[i] in teamList:
#             row[gradPreferences] = i

matrix.to_csv('Matrix.csv')

for i in grad_preferences:
    print (i)
    print (grad_preferences[i])

