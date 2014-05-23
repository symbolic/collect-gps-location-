__author__ = 'willard'
#-*- coding: gbk -*-

#功能：从模版文件中获取sheet列表和人名列表
#sheet列表用于生成目标文件的各个sheet
#人名用户文件查找
#如果模版文件有更新，需要重新运行本程序，更新列表文件

from xlrd import open_workbook

def main():
    print '请确保模版文件名为model.xls，并和本文件在同一个文件夹下'
    print '**********程序已启动**********'
    print '初始化中...'

    BauModel = open_workbook('model.xls')
    dest = open('sheetlist.txt', 'w')
    dest1 = open('namelist.txt', 'w')
    Number = 0
    print '正在获取信息，请稍候...'

    for s in BauModel.sheets():
        if ((Number > 0) & (Number%2 == 0)):
            if (Number != 2):
                dest1.write('\n')
            dest1.write(s.name.encode('gbk')[:-6])


        if (Number%2 ==1):
            if (Number >= 2):
                dest.write('\t')
            dest.write(s.name.encode('gbk'))


        Number += 1

    dest.close()
    dest1.close()
    print 'sheet已保存在sheetlist.txt中'
    print '人名已经保存在namelist.txt文件中'
    print '**********程序已退出**********'

if __name__ == '__main__':
    main()