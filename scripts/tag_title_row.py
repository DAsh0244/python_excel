import os
import sys
import openpyxl
import argparse
import re

from common import *

__version__ = '0.0.1'

# cli parser definition
parser = argparse.ArgumentParser('tag_title_row')
parser.add_argument('--version', action='version', version=__version__)
parser.add_argument('file', type=str, help='file name to be read in.')
parser.add_argument('start_col', type=str, help='Column to begin adding tags to.')
parser.add_argument('stop_col', type=str, help='Column to end adding tags to.')
parser.add_argument('--sheet', '-s', type=str, help='sheet name to pull data from')
parser.add_argument('--output', '-o', type=str, help='file name to save as')
parser.add_argument('--inclusive', '-i', action="store_false", # default=True, 
                    help='whether to use end column as inclusive or exclusive endpoint')
parser.add_argument('--direct_col', '-d', action="store_true", # default=False, 
                    help='use direct column name instead of name defined in first row of column')


    
# main function 
@time_execution
def main():
    args = parser.parse_args()
    
    wb, ws = get_workbook(args)
    title_mapping = excel_mappings(ws[1])

    # cols = parse_cols(args.col_name, title_mapping)
    if args.direct_col:
        cols = get_cols_by_name(args.start_col, args.stop_col, inclusive=args.inclusive)
    else:
        cols = get_cols_by_name(title_mapping[args.start_col], title_mapping[args.stop_col], inclusive=args.inclusive)
    
    format_title_row(ws[1], cols)
    write_workbook(wb, args)

    
if __name__ == '__main__':
    main()
    sys.exit(0)
