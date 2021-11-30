
from os import path
from sys import argv
from csv import reader as CSVReader, writer as CSVWriter, DictReader, DictWriter, get_dialect

CSVDIALECT = get_dialect('excel-tab')

class TXT2CSV(object):
    def __init__(self, fileinfo=None):
        assert isinstance(fileinfo, dict)
        self._fileinfo = fileinfo

    def convert_txt2csv(self, filename, fileinfo=None):
        if not path.exists(filename):
            raise Exception(f"not found: {filename}")
        if '.txt' not in filename:
            raise Exception(f"expect .txt in {filename}")
        outfile = filename.replace('.txt', '.csv')
        assert outfile != filename
        if '_Движение_денежных_средств_' in filename:
            self._transactions(filename, outfile)
        elif ('_Сделки_' in filename):
            self._deals(filename, outfile)

    def _deals(self, filename, outfile):
        fileType = None
        with open(filename, mode="r") as readf:
            first = readf.readline()
            if 'Акция' in first:
                fileType = self._deals_instruments
            elif 'Объём в валюте лота (в ед. валюты)' in first:
                fileType = self._deals_forex
        if fileType:
            with open(filename, mode="r") as readf, open(outfile, mode="w") as writef:
                fileType(readf, writef)

    def _deals_instruments(self, readf, writef):
        allreader = CSVReader(readf, dialect=CSVDIALECT)
        lines = list(allreader)
        header = _trim_last_empty(lines[2])
        assert 'Сумма платежа' in header
        myfields = header + ['Ticker', 'ISIN'] 
        myfields += self._fileinfo.keys()
        self._fix_duplicate_names_in_header(myfields)
        wcsv = DictWriter(writef, fieldnames=myfields, dialect=CSVDIALECT)
        wcsv.writeheader()
        n = len(lines)
        i = 3
        while i < n:
            if 'ISIN:' not in lines[i]:
                i += 1
            else:
                isin = lines[i][lines[i].index('ISIN:')+1]
                ticker = lines[i][1]
                end = next(j for j, l in enumerate(lines[i:]) if any('Итого по' in f for f in l)) + i 
                for row in lines[i+1: end]:
                    row = _trim_last_empty(row)
                    cells = {myfields[inx]: el for inx, el in enumerate(row)}
                    cells['Ticker'] = ticker
                    cells['ISIN'] = isin
                    for k in self._fileinfo:
                        cells[k] = self._fileinfo[k]
                    wcsv.writerow(cells)
                i = end+1


    def _fix_duplicate_names_in_header(self, header):
        for i,f in enumerate(header[::-1]):
            samenum = sum(1 for f2 in header if f2 == f)
            if samenum > 1:
                header[-i-1] = f"{f}{samenum}"


    def _deals_forex(self, readf, writef):
        allreader = CSVReader(readf, dialect=CSVDIALECT)
        lines = list(allreader)
        header = _trim_last_empty(lines[0])
        n = len(lines)
        myfields = header + [ 'From', 'To' ]
        myfields += self._fileinfo.keys()
        self._fix_duplicate_names_in_header(myfields)
        wcsv = DictWriter(writef, fieldnames=myfields, dialect=CSVDIALECT)
        wcsv.writeheader()
        i = 1
        def _getfield(from_, line):
            inx = next(i for i, e in enumerate(line) if from_ in e)
            return next(f for f in line[inx+1:] if f)  # First non empty
        while i < n:
            if 'Валюта лота:' not in lines[i]:
                i += 1
            else:
                assert 'Валюта лота:' in lines[i]
                tocurrency = _getfield('Валюта лота', lines[i])
                fromcurrency = _getfield('Сопряж. валюта', lines[i])
                end = next(j for j, l in enumerate(lines[i:]) if any('Итого по' in f for f in l)) + i 
                for row in lines[i+1: end]:
                    row = _trim_last_empty(row)
                    cells = {myfields[inx]: el for inx, el in enumerate(row)}
                    cells['From'] = tocurrency
                    cells['To'] = fromcurrency
                    for k in self._fileinfo:
                        cells[k] = self._fileinfo[k]
                    wcsv.writerow(cells)
                i = end

    def _transactions(self, filename, outfile):
        with open(filename, mode="r") as readf:
            lines = readf.readlines()
            currency = lines[0].split(CSVDIALECT.delimiter)[1]
            rcsv = DictReader(lines[1:-1], dialect=CSVDIALECT)
            myfields = ['Дата', 'Операция', 'Сумма зачисления', 'Сумма списания', 'Валюта']
            myfields += self._fileinfo.keys()
            isvalid = 'Операция' in lines[1]
            if isvalid:
                with open(outfile, mode="w") as writef:
                    wcsv = DictWriter(writef, fieldnames=myfields, dialect=CSVDIALECT)
                    wcsv.writeheader()
                    for row in rcsv:
                        if 'Итого' in row['Операция']:
                            continue
                        cells = dict((k, row[k]) for k in row if k in myfields)
                        cells['Валюта'] = currency
                        for k in self._fileinfo:
                            cells[k] = self._fileinfo[k]
                        if float(cells['Сумма зачисления']) > 0 and float(cells['Сумма списания']) > 0:
                            t = cells['Сумма зачисления']
                            cells['Сумма зачисления'] = None
                            wcsv.writerow(cells)
                            cells['Сумма зачисления'] = t
                            cells['Сумма списания'] = None
                            wcsv.writerow(cells)
                        else:
                            wcsv.writerow(cells)

def _trim_last_empty(arr):
    if not arr:
        return arr
    i = next(i for (i,e) in enumerate(arr[::-1]) if e)
    return arr[:-i-1]


if __name__ == '__main__':
    c = TXT2CSV()
    for f in argv[1:]:
        c.convert_txt2csv(f)
    input()
