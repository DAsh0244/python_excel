'''
python excel_concat.py JournalExport.xlsx JournalExportCopy-JB.xlsx -H -k -o journal_outputs\test_out.xlsx -c 20000
'''


import os
import sys
import openpyxl
import argparse
import re
import itertools

from common import *

__version__ = '0.0.1'

# cli parser definition
parser = argparse.ArgumentParser()
parser.add_argument('file', type=str, nargs='+', help='file(s) to be read in.')
parser.add_argument('--chunk_size', '-c', type=int, default='100', help='How many entries to write to new file(s). A size less than zero will write to a single file')
parser.add_argument('--output', '-o', type=str, help='base file name to save as. Resulting files will be incremented on the end. If the path given doesnt exist, it will be made.')
parser.add_argument('--h_concat', '-H', action="store_true", help='Join files in a sytle of horizontal row appending')
parser.add_argument('--keep_title', '-k', action="store_true", help='Whether to keep title row in each chunked file or not.')
parser.add_argument('--no_overwrite', '-n', action="store_true", help='If set, will not overwrite files')
parser.add_argument('--version', action='version', version=__version__)
# parser.add_argument('file', type=argparse.FileType('r'), nargs='+', help='file(s) name to be read in.')
# parser.add_argument('--sheet', '-s', type=str, nargs='+', help='sheet name to pull data from')

@time_execution(alert='loading workbooks')
def load_wbs(params):
    wbs = []
    wss = []
    for file in params.file:
        book = openpyxl.load_workbook(file, read_only=True)
        ws = book.worksheets[0] 
        if ws.calculate_dimension() == 'A1:A1':
            print('worksheet dimensions incorrect, recalculating...')
            ws.max_row = ws.max_column = None
            print('recalcualted dimesions as {}'.format(ws.calculate_dimension(force=True)))
            print('press <Ctrl+c> to cancel if dimensions are incorrect')
            # tmp_rows = 0
            # cols = []
            # while ws.rows:
                # cols.append(len(next(ws.rows)))
                # tmp_rows += 1
            # cols = max(cols)
            # ws.max_column = colnum_string(cols)
            # ws.max_row = tmp_rows
            # print('{}'.format(rows, cols, colnum_string))
        wbs.append(book)
        wss.append(book.worksheets[0])
        print('loaded {}'.format(file))
    return wbs, wss
    
@time_execution(alert='writing chunks')
def write_chunked_wbs(wss, chunk_size, concat=False, output=None, keep_title=True):
    chunk_num = 0
    if output is None:
        name = 'chunked_workbook'
        ext = 'xlsx'
    else:
        name, ext = output.split('.')
    output = '{}_{}.{}'.format(name, chunk_num, ext)
    if concat:
        # print('concat')
        # for wb1, wb2 in grouped(wbs, 2):
        for ws1, ws2 in pairwise(wss):
            # new_wb = openpyxl.Workbook(write_only = True)
            # print(new_wb)
            # new_ws = new_wb.create_sheet()
            try:
                rows1, rows2 =  ws1.rows, ws2.rows 
                if keep_title:
                    title_1 = row_vals(next(rows1))
                    title_2 = row_vals(next(rows2))[2:]  # first two cols are duplicate of first two cols of first file
                    # title_2 = list(filter_list(row_vals(next(rows2)), title_1))
                    title_row = flatten([title_1, title_2])
                    # print(title_row)
                while rows1 or rows2:
                    new_wb = openpyxl.Workbook(write_only=True)
                    new_ws = new_wb.create_sheet()
                    if keep_title:
                        new_ws.append(title_row)
                    if chunk_size < 0:
                        while True:
                            new_ws.append(flatten([row_vals(next(rows1)), row_vals(next(rows2)[2:])]))  # first two cols are duplicate of first two cols of first file
                            # new_ws.append(flatten([row_vals(next(rows1)), row_vals(next(rows2))]))
                    else:
                        for line in range(chunk_size):
                            new_ws.append(flatten([row_vals(next(rows1)), row_vals(next(rows2)[2:])]))  # first two cols are duplicate of first two cols of first file
                            # new_ws.append(flatten([row_vals(next(rows1)), row_vals(next(rows2))]))
                    new_wb.save(output)
                    print('wrote {}'.format(output))
                    chunk_num += 1
                    output = '{}_{}.{}'.format(name, chunk_num, ext)
            except StopIteration:
                new_wb.save(output)    
    else: 
        for ws in wss:
            # new_wb = openpyxl.Workbook(write_only=True)
            # new_ws = new_wb.create_sheet()
            try:
                rows1 =  ws1.rows 
                if keep_title:
                    title_row = row_vals(next(rows1))
                    print(title_row)
                while rows1:
                    new_wb = openpyxl.Workbook(write_only=True)
                    new_ws = new_wb.create_sheet()
                    if keep_title:
                        new_ws.append(title_row)
                    if chunk_size < 0:
                        while True:
                            new_ws.append(row_vals(next(rows1)))
                    else:
                        for line in range(chunk_size):
                            new_ws.append(row_vals(next(rows1)))
                    new_wb.save(output)
                    print('wrote {}'.format(output))
                    chunk_num += 1
                    output = '{}_{}.{}'.format(name, chunk_num, ext)
            except StopIteration:
                new_wb.save(output)    

def ensure_dir(file_path):
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)
                
                
@time_execution('Total operation')
def main():
    args = parser.parse_args()
    # print(vars(args))
    for file in iter(args.file):
        if not os.path.exists(file):
            raise OSError
    if args.output:
        os.makedirs(os.path.dirname(args.output), exist_ok=True)
    workbooks, worksheets = load_wbs(args)
    write_chunked_wbs(worksheets, args.chunk_size, concat=args.h_concat, 
                      output=args.output, keep_title=args.keep_title)
    
if __name__ == '__main__':
    main()
    sys.exit(0)
