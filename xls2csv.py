#!python3

usage = '''
Парсит файл отчета .xls из БКС (broker.ru) в .csv файлы согласно таблицам найденным в файле отчета. Создает несколько файлов .csv рядом с файлом .xls с таким же названием, но без расширения .xls.

Использование:

    bcsxls2csvs.py <путь к файлу .xls>

'''

from xlrd import open_workbook
from sys import argv
from csv import writer
from re import compile
from os import path, listdir
from txt2csv import convert2csv



class Iterator(object):
    def __init__(self):
        self._files = []
        self._table_rows = []
        self._buffer = []
        self._part_name = None
        self._part_table_count = 0
        self._part_table_name = None
        self._state = self._state_init

    def next(self, line):
        def clean(e):
            if isinstance(e, str):
                return e.replace('\n', ' ').replace('\t', ' ')
            return e
        line = [clean(e) for e in line]
        self._state(line)

    def files(self):
        yield from self._files

    def _state_init(self, line):
        if line_ispartheader(line):
            self._state_part_header(line)
        # else ignore this line

    def _state_part_header(self, line):
        assert line_ispartheader(line)
        self._buffer.append(line)
        self._state = self._state_part_start

    def _state_part_start(self, line):
        if line_isempty(line):
            self._part_name = clean_str(self._buffer.pop())
            self._part_table_count = 0
            self._state = self._state_start_table
        else:  # Not a part
            self._buffer.clear()
            self._state = self._state_init
            self._state_init(line)

    def _state_start_table(self, line):
        if line_isempty(line):
            pass  # Wait table or next part.
        elif line_ispartheader(line):
            self._state_part_header(line)
        else:
            assert not self._table_rows, f'Failed to start table: has previous rows. Was previous flushed?'
            self._part_table_count += 1
            self._part_table_name = gettablefilename(self._part_name, self._part_table_count)
            self._state = self._state_parse_table
            self._state_parse_table(line)

    def _state_parse_table(self, line):
        if line_isempty(line):  # Table end.
            self._state_end_table()
        else:
            self._table_rows.append(line)
            assert self._state == self._state_parse_table, f'Should continue parsing table but: {self._state}'

    def _state_end_table(self):
        rows = self._table_rows
        MINIMUM_TABLE_ROWS = 3
        isgood = (rows
                    and len(rows) >= MINIMUM_TABLE_ROWS
                    and not isbottomasterix(rows)
        )
        if isgood:
            assert self._part_table_name
            self._files.append((self._part_table_name, rows.copy()))
        else:
            self._part_table_count -= 1
        self._part_table_name = None
        self._table_rows.clear()
        self._state = self._state_start_table

part_header_re = compile(r'\s*(\d\.(\d\.)*\s+\w+.*)')
word_re = compile(r'((\d\.(\d\.)*)|\w+)')

def clean_str(line):
    return ' '.join(m.group(0) for m in word_re.finditer(ljoin(line)))

def gettablefilename(name, tablecount):
    MAXCHARS = 50
    name = name.replace(' ', '_')[:MAXCHARS]
    return f'{name}_T{tablecount}'

def line_ispartheader(line):
    return part_header_re.match(ljoin(line))

def line_isempty(line):
    return not isinstance(line, list) or all(not e or not str(e).strip() for e in line)

def ljoin(line):
    return ' '.join(str(e) for e in line)

asterixre = compile(r'^\s*\(\d+\*\)\s*-\s*.*')
def isbottomasterix(rows):
    if not rows or not len(rows)>1 or not isinstance(rows[0], list):
        return False
    r = ''.join(str(e) for e in rows[0])
    return asterixre.match(r)

basefilenamere = compile(r'(.*)(\.(csv|xls))?')

def getbasefilename(filename):
    m = basefilenamere.search(filename)
    assert m, f"could not get base file name for '{filename}'"
    return m.group(1)

def parser(xlsfilepath):
    iterator = Iterator()
    # Input:
    book = open_workbook(xlsfilepath)
    sh = book.sheet_by_index(0)
    for rindex in range(sh.nrows):
        iterator.next(sh.row_values(rindex))
    # Output:
    for fname, lines in iterator.files():
        bname = path.basename(xlsfilepath).replace('.xls', '')
        outfile = f"{bname}_{getbasefilename(fname)}.txt"
        outfile = path.join(path.dirname(xlsfilepath), outfile)
        with open(outfile, 'w') as fhandler:
            w = writer(fhandler, dialect='excel-tab')
            for l in lines:
                w.writerow(l)

if __name__ == "__main__":
    if len(argv) < 2:
        print(usage)
    else:
        xlsf = argv[1]
        try:
            parser(xlsf)
        except:
            print(f'Failed to transform "{xlsf}" to txt.')
            raise
        xlsfiledir = path.dirname(xlsf)
        for txtf in listdir(xlsfiledir):
            if not txtf.endswith('.txt'):
                continue
            try:
                convert2csv(path.join(xlsfiledir, txtf))
            except:
                print(f'Failed to transform "{txtf}" to csv.')
                raise
