import openpyxl
from datetime import datetime
import urllib.request
import os.path

# Write data to file
filename = "PERM_Disclosure_Data_FY2019.xlsx"

if not os.path.isfile(filename):
    response = urllib.request.urlopen('https://www.foreignlaborcert.doleta.gov/pdf/PerformanceData/2019/PERM_Disclosure_Data_FY2019.xlsx')
    data = response.read()
    file_ = open(filename, 'wb')
    file_.write(data)
    file_.close()

book = openpyxl.load_workbook(filename, data_only=True)

sheet = book.active

rows = sheet.rows

values = []

cova = {}
mdata = datetime(2000, 1, 1)

for row in rows:
    if (row[0].value == 'CASE_NUMBER'):
        continue
    #datetime_object = datetime.strptime(row[3].value, '%m/%d/%Y').strftime('%m/%Y')
    datetime_object = row[3].value.strftime('%Y-%m')
    # check maxprocess data
    if (mdata < row[3].value) and (row[2].value == "Certified"):
        mdata = row[3].value

    if cova.get(datetime_object):
        cova[datetime_object][row[2].value] += 1
    else:
        cova[datetime_object] = dict({"Certified": 0, "Withdrawn": 0, "Denied": 0, "Certified-Expired": 0})
        cova[datetime_object][row[2].value] = 1

totalitems = 0
for key in sorted(cova):
    for typo in cova[key]:
        print(" Month: {0} CaseStatus: \"{1}\" items: {2} ".format(key, typo, cova[key][typo]))
        totalitems += cova[key][typo]
    print("x" * 50)

print("Total items: {0}".format(totalitems))
print("Last processed date: {0}".format(mdata.strftime('%m/%d/%Y')))
