from openpyxl import load_workbook
from shareplum import Site
from requests_ntlm import HttpNtlmAuth
import getpass

excelPath = r"C:\Users\zoran.gligoric\Documents\Stari taskovi za zatvaranje DIREKTNA 12022020.xlsx"
wb = load_workbook(excelPath, read_only=True)

username = 'zoran.gligoric'
pwd = 'HawHaw16!'#getpass.getpass('Password: ')

auth = HttpNtlmAuth("PEXIMBG\\" + username, pwd)
site = Site('http://tsp/', auth=auth)
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
    fields = ['ID', 'Test specialist comments']
    query = {'Where': [('Eq', 'ID', str(r))]}
    sp_data = sp_list.GetListItems(fields=fields, query=query)
    for dt in sp_data:
        comment = dt.get('Test specialist comments')
        if comment==None:
            comment=''
        comment = comment+'\n\n[12.02.2020. Aleksandra Milisavljevic]\nPayment statusi: „For payment“ i „For invoicing“  azurirani sa „Not To Be Invoiced“. ' \
                                                     'Task pripada grupi taskova iz perioda postojanja Credy i KBM banke do spajanja sa Findomestik bankom, ' \
                                                     'tacnije od 2012. g. zakljucvno sa 30.06.2017.g. za koje banka nije prihvatila placanje.\n'
        #comment = ''
        sp_update.append({'ID': r, 'Payment Status': 'Not to be invoiced', 'Test specialist comments': comment})

#sp_list.UpdateListItems(data=sp_update, kind='Update')