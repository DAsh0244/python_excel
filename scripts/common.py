import time as _time
import openpyxl as _openpyxl
import string as _string
import itertools as _itertools


def pairwise(iterable):
    "s -> (s0,s1), (s1,s2), (s2, s3), ..."
    a, b = _itertools.tee(iterable)
    next(b, None)
    return zip(a, b)

def grouped(iterable, n):
    "s -> (s0,s1,s2,...sn-1), (sn,sn+1,sn+2,...s2n-1), (s2n,s2n+1,s2n+2,...s3n-1), ..."
    return zip(*[iter(iterable)]*n)

def filter_list(full_list, excludes):
    s = set(excludes)
    return (x for x in full_list if x not in s)
    
def flatten(iterable, n=1):
    """
    assumes uniform nesting
    """
    tmp = iterable
    for i in range(n):
        tmp = [item for sublist in tmp for item in sublist]
    return tmp
    
def _sec2time(sec, n_msec=5):
    ''' 
    Convert seconds to "D days, HH:MM:SS.FFF"    
    '''
    if hasattr(sec,'__len__'):
        return [_sec2time(s) for s in sec]
    m, s = divmod(sec, 60)
    h, m = divmod(m, 60)
    d, h = divmod(h, 24)
    d, h, m = int(d),int(h), int(m) 
    if n_msec > 0:
        pattern = '{{:02d}}:{{:02d}}:{{:0{space}.{msec}f}}'.format(space=n_msec+3, msec=n_msec)
        # pattern = '%%02d:%%02d:%%0%d.%df' %(n_msec+3, n_msec)
    else:
        pattern = '{:02d}:{:02d}:{:02d}'
        # pattern = r'%02d:%02d:%02d'
    if d == 0:
        # return pattern % (h, m, s)
        return pattern.format(h, m, s)
    return ('{:02d} days, ' + pattern).format(d, h, m, s)

def time_execution(desc=None, alert=None):
    def decorator(method):
        def wrapper(*args, **kw):
            if alert:
                print('\n{}:'.format(alert.title()))
            t1 = _time.perf_counter()
            result = method(*args, **kw)
            t2 = _time.perf_counter()
            if desc:
                msg = '{} took {}'.format(desc, _sec2time(t2 - t1))
            else:
                msg = '"{}": Operation took {}'.format(method.__name__, _sec2time(t2 - t1))
            print(msg)
            return result
        return wrapper
    return decorator

def write_row(row_num, row_data, sheet):
    for col,cell in enumerate(row_data, start=1):
        sheet.cell(row=row_num, column=col).value = cell.value
    
def excel_mappings(title_row):
    return {entry.value.strip(): entry.column for entry in title_row}
    
def inv_map(dictionary: dict):
    return {v: k for k, v in dictionary.items()}

def increment_char(c):
    """
    Increment an uppercase character, returning 'A' if 'Z' is given
    """
    return chr(ord(c) + 1) if c != 'Z' else 'A'

def increment_str(s):
    lpart = s.rstrip('Z')
    num_replacements = len(s) - len(lpart)
    new_s = lpart[:-1] + increment_char(lpart[-1]) if lpart else 'A'
    new_s += 'A' * num_replacements
    return new_s
    
def col2num(col):
    num = 0
    for c in col:
        if c in _string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num    

def colnum_string(n):
    div=n
    string=""
    temp=0
    while div>0:
        module=(div-1)%26
        string=chr(65+module)+string
        div=int((div-module)/26)
    return string
    
def row_vals(row):
    return [cell.value for cell in row]
    
@time_execution('getting workbook')
def get_workbook(args):
    wb = _openpyxl.load_workbook(args.file)
    if not args.sheet:
        ws = wb.worksheets[0]
    else:
        ws = wb.get_sheet_by_name(args.sheet)
    return wb, ws
    
@time_execution('writing workbook')
def write_workbook(wb, args):
    if not args.output:
        wb.save(args.file)
    else:
        wb.save(args.output)
           
def get_cols_by_name(start, stop, inclusive=True):
    if inclusive:
        cols = list(map(colnum_string, range(col2num(start), col2num(stop)+1)))
    else:
        cols = list(map(colnum_string, range(col2num(start), col2num(stop))))
    return cols

def format_title_row(row, cols, tag='Interest', tag_label=':'):
    leading_tag = '{}{} '.format(tag, tag_label)
    for cell in row:
        if cell.column in cols and not cell.value.startswith(leading_tag):
            cell.value = leading_tag + cell.value


''' EOF '''
