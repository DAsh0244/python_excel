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
parser.add_argument('--version', action='version', version=__version__)
parser.add_argument('file', type=str, nargs='+', help='file(s) name to be read in.')
parser.add_argument('--h_concat', '-H', action="store_true", help='join files in a sytle of horizontal row appending')
parser.add_argument('--output', '-o', type=str, help='file name to save as. Will save the next file as file_name_01.extension')
parser.add_argument('--keep_title', '-k', action="store_true", help='whether to keep title row in each chunkned file or not.')
parser.add_argument('--dont_overwrite', '-d', action="store_true", help='If set, will not overwrite files')
parser.add_argument('--chunk_size', '-c', type=int, default='100', help='how many entries to write to new file(s). A size less than zero will write to a single file')
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
            new_wb = openpyxl.Workbook()
            new_ws = new_wb.active
            try:
                rows1, rows2 =  ws1.rows, ws2.rows 
                if keep_title:
                    title_1 = row_vals(next(rows1))
                    title_2 = row_vals(next(rows2))[2:]
                    # title_2 = list(filter_list(row_vals(next(rows2)), title_1))
                    title_row = flatten([title_1, title_2])
                    # print(title_row)
                while rows1 or rows2:
                    new_wb = openpyxl.Workbook()
                    new_ws = new_wb.active
                    if keep_title:
                        new_ws.append(title_row)
                    if chunk_size < 0:
                        while True:
                            new_ws.append(flatten([row_vals(next(rows1)), row_vals(next(rows2)[2:])]))
                    else:
                        for line in range(chunk_size):
                            new_ws.append(flatten([row_vals(next(rows1)), row_vals(next(rows2)[2:])]))
                            # new_ws.append(flatten([row_vals(next(rows1)), row_vals(next(rows2))]))
                    new_wb.save(output)
                    print('wrote {}'.format(output))
                    chunk_num += 1
                    output = '{}_{}.{}'.format(name, chunk_num, ext)
            except StopIteration:
                new_wb.save(output)    
    else: 
        # raise NotImplementedError('Straight concat is not yet supported.')
        for ws in wss:
            new_wb = openpyxl.Workbook()
            new_ws = new_wb.active
            try:
                rows1 =  ws1.rows 
                if keep_title:
                    title_row = row_vals(next(rows1))
                    print(title_row)
                while rows1:
                    new_wb = openpyxl.Workbook()
                    new_ws = new_wb.active
                    if keep_title:
                        new_ws.append(title_row)
                    for line in range(chunk_size):
                        new_ws.append(row_vals(next(rows1)))
                    new_wb.save(output)
                    print('wrote {}'.format(output))
                    chunk_num += 1
                    output = '{}_{}.{}'.format(name, chunk_num, ext)
            except StopIteration:
                new_wb.save(output)    

                
@time_execution('Total operation')
def main():
    args = parser.parse_args()
    # print(vars(args))
    workbooks, worksheets = load_wbs(args)
    write_chunked_wbs(worksheets, args.chunk_size, concat=args.h_concat, 
                      output=args.output, keep_title=args.keep_title)
    
if __name__ == '__main__':
    main()
    sys.exit(0)

