from openpyxl import load_workbook
from shareplum import Site
from requests_ntlm import HttpNtlmAuth

excelPath = r"C:\Users\zoran.gligoric\Documents\zatvaranjeTEST.xlsx"
wb = load_workbook(excelPath, read_only=True)

username = 'zoran.gligoric'
pwd = 'HawHaw13!'
sajt = "tsp"

if "test" in excelPath.lower():
    sajt = "tsp"

auth = HttpNtlmAuth("PEXIMBG\\" + username, pwd)
site = Site('http://'+sajt+'/', auth=auth)
sp_list = site.List('Customer Requests')

sheet = wb['Sheet1']

i = 2
v = sheet['A'+str(i)]
n = {}
while v.value:
    n[i-2] = str(v.value)
    i = i + 1
    v = sheet['A' + str(i)]

print(n)

sp_update = []

for r in n.values():
    fields = ['ID', 'Request Status']
    query = {'Where': [('Eq', 'ID', str(r))]}
    sp_data = sp_list.GetListItems(fields=fields, query=query)
    reqstat = ""
    for dt in sp_data:
        reqstat = dt.get('Request Status')
        reqstat = reqstat.lower()
        if "completed" not in reqstat and "reject" not in reqstat and "remove" not in reqstat:
            sp_update.append({'ID': r, 'Request Status': 'BSW-Delivered, Test successful'})

#sp_list.UpdateListItems(data=sp_update, kind='Update')
