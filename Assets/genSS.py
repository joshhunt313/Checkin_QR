from openpyxl import Workbook
import json

# Read info from fake names file
people = open('samplePpl.json').read()
pplJson = json.loads(people)

# Add people to list
pplList = []
for p in pplJson:
    newPerson = [p['id'], p['age'], p['firstName'], p['lastName'], p['gender'], p['email'], p['phone'], p['address'], p['registered']]
    pplList.append(newPerson)

wb = Workbook()
dest = '../people.xlsx'
ws = wb.active
ws.title = 'Hackers'

# Column titles
ws.append(['id', 'age', 'first', 'last', 'gender', 'email', 'phone', 'address', 'timestamp'])

# Put everybody on the spreadsheet
for p in pplList:
    ws.append(p)

wb.save(dest)