#! /usr/env/bin python3

import openpyxl as xl
from collections import Counter

CATAGORY_COL = 'C'

def summarize_event_attendance(workbook_path):
    wb = xl.load_workbook(workbook_path)
    stats  = {k:None for k in wb.sheetnames}
    for sheet in wb:
        stats[sheet.title] = Counter(cell.value for cell in sheet[CATAGORY_COL])

    with open('event_results.csv','w') as file:
        for key,val in stats.items():
            file.write('{}\n'.format(key))
            for k,v in val.items():
                file.write(',{},{}\n'.format(k,v))
    wb.close()


if __name__ == '__main__':
    import sys
    summarize_event_attendance(sys.argv[1])
    sys.exit(0)