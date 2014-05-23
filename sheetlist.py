__author__ = 'willard'
#-*- coding: utf-8 -*-
import os

#功能：从模版文件中获取sheet列表和人名列表
#sheet列表用于生成目标文件的各个sheet
#人名用户文件查找
#如果模版文件有更新，需要重新运行本程序，更新列表文件

from xlrd import open_workbook

def run_check(file_name):
    if not os.path.isfile(file_name):
        print 'The file "'+file_name+'" is not exist, please check it again.'
        return 0


def main():
    print '********************'
    print 'program is running...'
    if run_check('model.xls') == 0:
        return 0
    if run_check('gpsjicheng.xls') == 0:
        return 0

    BauModel = open_workbook('model.xls')
    merge_file = open_workbook('gpsjicheng.xls')
    dest = open('sheetlist.txt', 'w')
    dest1 = open('namelist.txt', 'w')
    dest2 = open('sheetlist_jicheng.txt', 'w')
    number = 0
    print 'getting information at now, please wait a moment...'

    for s in BauModel.sheets():
        #最后三个字是“日行程”
        print s.name.encode('utf-8')[-9:]
        if number > 0:
            if number > 1:
                dest1.write('\n')
            dest1.write(s.name.encode('utf-8')[:-9])

        if number >= 1:
            dest.write('\t')
        dest.write(s.name.encode('utf-8'))

        number += 1

    dest.close()
    dest1.close()

    number = 0
    for s in merge_file.sheets():
        if number > 0:
            dest2.write('\t')
        dest2.write(s.name.encode('utf-8'))
        number += 1
    dest2.close()
    print 'sheets hava been saved in file "sheetlist.txt"'
    print 'The names have been saved in file "namelist.txt"'
    print 'The jicheng_sheets have been saved in file "sheetlist_jicheng.txt"'
    print '**********software will be closed**********'

if __name__ == '__main__':
    main()