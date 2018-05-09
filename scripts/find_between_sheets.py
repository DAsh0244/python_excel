# find name and earliest month found
#
# name: 
# - in  A col
# - Names start at A19 downwards
# - [<last>, <first>]
# - already unique per sheet
#
# month:
# - in sheet name
# - only want up to june
#
# count only if the corresponding cols [G-O] shows a visit (not empty cells)
#
# dict format: {name: [months]}
#
#
# C:\Users\Danyal\Documents\GitHub\python-excel>scripts\find_between_sheets.py "C:\Users\Danyal\Documents\GitHub\python-excel\test_docs\find_between\Tarrant County General Stats 2017 - June.xlsx" "C:\Users\Danyal\Documents\GitHub\python-excel\test_docs\find_between\General Stats 2017 collin.xlsx" -o test_docs\find_between\results.txt

import os
import sys
import glob
import argparse
import openpyxl
import common
import itertools
from string import whitespace
from calendar import month_name


MONTH_MAPPING = {v: k for k,v in enumerate(month_name)}
ACTIVE_MONTHS = set(month_name[1:8]) # [Jan, Aug)
CLUBHOUSES = {'tarrant', 'dallas', 'collin'}
BAD_VALS = set().union((0, None),('name', 'Name') ,*[('name'+char , 'Name'+char, char) for char in whitespace])

# <last> | <first> | <clubhouse> | <Earliest> | <Latest> 
FMT_WIDTHS = (20,20,10,10,10)
FMT_STR = ('|'.join(itertools.repeat('{{:^{}}}', len(FMT_WIDTHS))) + '\n').format(*FMT_WIDTHS).format
SPACING_STR = '+'.join(['-'*width for width in FMT_WIDTHS]) + '\n'

def remove_bad_cells(cells, excludes=(None, '')):
    s = set(excludes)
    return (x for x in cells if x.value not in s)

def get_clubhouse(book):
    name = os.path.basename(book).lower()
    for location in CLUBHOUSES:
        if location in name:
            return location
    return 'N/A'
    
def main(workbooks, output=None):
    if output is None:
        output = sys.stdout
    else:
        global FMT_STR, SPACING_STR
        if '.csv' in output:
            FMT_STR = (','.join(itertools.repeat('{}', len(FMT_WIDTHS))) + '\n').format
            SPACING_STR = ''
        # elif '.txt' in output:
        #     FMT_STR = '\t'.join(itertools.repeat('{}', len(FMT_WIDTHS))) + '\n'
        #     SPACING_STR = ''
        output = open(output, 'w')
    
    visitors = dict()
    for book in workbooks:
        wb = openpyxl.load_workbook(book)
        clubhouse = get_clubhouse(book)
        active_sheets = {sheet for sheet in wb if sheet.title[:-5] in ACTIVE_MONTHS}
        for sheet in active_sheets:
            month = sheet.title[:-5]
            names = remove_bad_cells(sheet['A'][18:], excludes=BAD_VALS) # A19 onward
            
            for name in names:
                if list(common.filter_list(common.row_vals(sheet['G{}'.format(name.row): 'O{}'.format(name.row)][0]),
                                           excludes=BAD_VALS)
                        ):
                    key = '{},{}'.format(name.value.strip(), clubhouse)
                    try:
                        visitors[key].append(month) 
                    except KeyError: # new person
                        visitors[key] = [month]
        wb.close()
    try:
        output.write(FMT_STR('Lastname', 'Firstname', 'Clubhouse', 'Earliest', 'Latest'))  # write header row
        output.write(SPACING_STR)
        for name in visitors:
            visits = sorted(visitors[name], key=lambda x : MONTH_MAPPING[x])
            split_key = [item.strip() for item in name.split(',')]
            output.write(FMT_STR(*split_key, visits[0], visits[-1]))
        output.write(SPACING_STR)
    except BrokenPipeError: # for redirection piping.
        pass
    
    if not output.isatty():
        output.close()
    
    parser = argparse.ArgumentParser()
    parser.add_argument('workbooks', nargs='+', type=str, action='store')
    parser.add_argument('--output', '-o', type=str, action='store', default=None)
    parser.add_argument('--file', '-f', action='store_true')
        
if __name__ == '__main__':
    args = parser.parse_args()
    if args.file:
        files = []
        for path in args.workbooks:
            files.append(glob.iglob(os.path.join(path, '*.xlsx')))
        args.workbooks = itertools.chain.from_iterable(files)
    main(args.workbooks, args.output)
    sys.exit(0)
