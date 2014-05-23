#*-* coding: gbk *-*
from xlwt import *
mark = '\t'


def main():
    print 'write data to excel'
    namelist = open('sheetlist.txt', 'r')
    wb = Workbook()

    name = namelist.readline()
    print len(name.split(mark))
    for i in range(len(name.split(mark))):
        print i
        print name
        print name.split(mark)[i]
        ws0 = wb.add_sheet(name.split(mark)[i].decode('gbk'))
    ws0.write_merge(0, 2, 0, 1, 'test1')

    wb.save('test.xls')

if __name__ == '__main__':
    main()