import os
import sys
import openpyxl
from collections import namedtuple
from operator import itemgetter

Entry = namedtuple('Entry', 'id name')
# FILE = 'test.xlsx'
# FILE = 'clean-data-cscntc.v14.xlsx'
# FILE = 'checking.xlsx'
FILE = 'Copy of clean-data-cscntc.v16.xlsx'
COL_NAME_1 = 'A'
COL_NAME_2 = 'C'
# DELIM_CHAR = ','


if len(sys.argv) > 1:
    try:
        FILE = sys.argv[1]
        COL_NAME_1 = sys.argv[2]
        DELIM_CHAR = sys.argv[3]
    except:
        pass

wb = openpyxl.load_workbook(FILE)
sheets = wb.get_sheet_names()
worksheet = wb.get_sheet_by_name(sheets[0])
names = set()
unique = set()
col_1 = []
col_2 = [] 
new_vals = []
new_entries = []

for entry_1, entry_2 in zip(worksheet[COL_NAME_1], worksheet[chr(ord(COL_NAME_1)+1)]):
  col_1.append(Entry(entry_1, entry_2))
  names.add(entry_1.value) 
  unique.add(Entry(entry_1.value, entry_2.value))

for  entry_1, entry_2 in zip(worksheet[COL_NAME_2], worksheet[chr(ord(COL_NAME_2)+1)]):
  col_2.append(Entry(entry_1, entry_2))
  names.add(entry_1.value) 
  unique.add(Entry(entry_1.value, entry_2.value))

col_1_ids = set([entry.id.value for entry in col_1])
col_2_ids = set([entry.id.value for entry in col_2])

new_vals = list(col_2_ids - col_1_ids)
for id in new_vals:
    name = [item.name.value for item in col_2 if item[0].value == id][0]
    new_entries.append(Entry(id,name)) 
    

# print('{}: {}'.format(entry.id.value, entry.name.value))

# unique_list = list(unique)
# unique_list.remove(Entry(None,None))
# unique_list.remove(Entry('Account Number','Account Name'))
# unique_list = sorted(unique_list, key=itemgetter(0))
# for i in range(2, len(unique_list)+2):
    # worksheet['G{}'.format(i)].value = unique_list[i-2].id
    # worksheet['H{}'.format(i)].value = unique_list[i-2].name

wb.save('results.xlsx')
sys.exit(0)
