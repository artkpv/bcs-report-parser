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
        MINIMUM_TABLE_ROWS = 3
        if self._table_rows and len(self._table_rows) >= MINIMUM_TABLE_ROWS:
            assert self._part_table_name
            self._files.append((self._part_table_name, self._table_rows.copy()))
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
    return all(not e.strip() for e in line)

def ljoin(line):
    return ' '.join(str(e) for e in line)

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
        fname = f"{getbasefilename(xlsfilepath)}_{getbasefilename(fname)}.csv"
        with open(fname, 'w') as fhandler:
            w = writer(fhandler)
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
            print(f'Failed to transform "{xlsf}" to csv.')
            raise
