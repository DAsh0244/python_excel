'''
 attendance_information.py "Attendance_2\February 2018 Attendance with Demographics.xlsx" "Dallas" "Tarrant" "Collin"
'''
    
import os.path as osp
from common import *
from collections import (
                                        OrderedDict, 
                                        Counter
                                      )


            
if __name__ == '__main__':
    import sys
    path = sys.argv[1]
    col_of_interest = sys.argv[2]
    try:
        col_title = sys.argv[3]
    except IndexError:
        col_title = None
    wb= get_workbook(FakeArgs(osp.join(TEST_DIR,path),None))[0]
    for ws in wb.worksheets:
        if not col_title :
            col_title = ws['{}3'.format(col_of_interest)].value 
        from pprint import pprint
        info = {k:val for k,val in  Counter(row_vals(ws[col_of_interest])).items() if k not in {None,'',col_title}}
        with open(osp.join(TEST_DIR,osp.dirname(path),'{}_{}.csv'.format(ws.title,col_title.replace(' ','_') if col_title else col_of_interest)), 'w') as file:
            file.write('{}:,Num Entries:\n'.format(col_title if col_title else col_of_interest))
            for k,v in info.items():
                file.write('{}:,{}\n'.format(k,v))
            # worksheet = wb.get_sheet_by_name(ws)
    
    wb.close()
    sys.exit(0)